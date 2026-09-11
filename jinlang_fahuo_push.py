# -*- coding: utf-8 -*-
"""锦浪办公发货汇总推送工具 - 持续化 Tk 版"""
from __future__ import annotations
import json, os, re, threading, time, urllib.error, urllib.request
from collections import defaultdict
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from openpyxl import load_workbook

APP_NAME = "锦浪办公发货助手"
APP_DIR_NAME = "JinlangOfficeFahuo"
DEFAULT_WEBHOOK = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=f5dd8950-08ad-4c4d-b157-4c9fc7fae2b8"
MODEL_KEYS = ("型号", "产品型号", "物料型号", "物料名称", "产品名称", "物料", "产品")
REMAIN_KEYS = ("未发货数量", "未发数量", "剩余数量", "剩余未发", "待发数量", "未发货", "待发")
SALES_KEYS = ("销售数量", "订单数量", "订购数量", "订单量", "销售量", "数量")
SHIPPED_KEYS = ("已发货数量", "已发数量", "出库数量", "发货数量", "累计发货", "已出库数量")

def app_config_path():
    base = os.getenv("APPDATA")
    folder = Path(base) / APP_DIR_NAME if base else Path.home() / f".{APP_DIR_NAME}"
    folder.mkdir(parents=True, exist_ok=True)
    return folder / "config.json"

def load_config():
    data = {"webhook": DEFAULT_WEBHOOK, "unit_mode": "自动识别"}
    try:
        p = app_config_path()
        if p.exists():
            saved = json.loads(p.read_text(encoding="utf-8"))
            if isinstance(saved, dict): data.update({k:v for k,v in saved.items() if isinstance(v,str)})
    except Exception: pass
    return data

def save_config(data):
    app_config_path().write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

def norm(v): return re.sub(r"\s+", "", str(v or "")).lower()
def match(header, keys):
    h = norm(header)
    return any(h == norm(k) or norm(k) in h for k in keys)

def number(v):
    if v is None or v == "": return None
    if isinstance(v, (int,float)): return float(v)
    s = str(v).strip()
    for x in (",","，","吨","kg","KG","Kg","公斤","千克"): s = s.replace(x, "")
    try: return float(s)
    except ValueError: return None

def detect_unit(header):
    h = str(header or "").lower()
    if any(x in h for x in ("kg","千克","公斤")): return "kg"
    if "吨" in h or "(t)" in h or "（t）" in h: return "ton"
    return None

def read_table(path):
    wb = load_workbook(path, data_only=True, read_only=True)
    try:
        for ws in wb.worksheets:
            rows = list(ws.iter_rows(values_only=True))
            for hi, row in enumerate(rows[:30]):
                headers = {i:str(v or "").strip() for i,v in enumerate(row)}
                model_col = next((i for i,h in headers.items() if match(h, MODEL_KEYS)), None)
                remain_col = next((i for i,h in headers.items() if match(h, REMAIN_KEYS)), None)
                sales_col = next((i for i,h in headers.items() if match(h, SALES_KEYS)), None)
                shipped_col = next((i for i,h in headers.items() if match(h, SHIPPED_KEYS)), None)
                qty_col = remain_col if remain_col is not None else sales_col
                if model_col is None or qty_col is None: continue
                mode = "remaining" if remain_col is not None else ("sales_minus_shipped" if shipped_col is not None else "sales_as_remaining")
                data = []
                for r in rows[hi+1:]:
                    model = str(r[model_col] or "").strip() if model_col < len(r) else ""
                    if not model: continue
                    q = number(r[qty_col] if qty_col < len(r) else None)
                    if q is None: continue
                    if mode == "sales_minus_shipped": q -= number(r[shipped_col] if shipped_col < len(r) else None) or 0
                    if q > 0: data.append((model, q))
                if data:
                    return {"rows":data,"mode":mode,"qty_header":headers[qty_col],"model_header":headers[model_col],"sheet":ws.title}
    finally: wb.close()
    raise ValueError("没有找到可识别的数据。至少需要‘型号/产品’和‘剩余数量/销售数量/订单数量’等列。")

def fmt_qty(v): return f"{v:.3f}".rstrip("0").rstrip(".")
def mode_text(mode):
    return {"remaining":"直接汇总剩余/未发数量","sales_minus_shipped":"订单/销售数量 - 已发货/出库数量","sales_as_remaining":"未检测到已发货列，暂按销售/订单数量作为未发数量"}.get(mode, mode)
def summarize(rows, is_kg):
    totals = defaultdict(float)
    for model, qty in rows: totals[model] += qty/1000.0 if is_kg else qty
    return sorted(((m,q) for m,q in totals.items() if q>0), key=lambda x:(-x[1],x[0]))

def build_markdown(path, totals, rule, unit_text):
    lines = ["### 发货待办汇总", f"> 文件：{os.path.basename(path)}", f"> 统计时间：{time.strftime('%Y-%m-%d %H:%M')}", f"> 规则：{rule}；单位：{unit_text}", ""]
    total = 0.0
    for model, qty in totals:
        lines.append(f"- **{model}**：剩余 **{fmt_qty(qty)} 吨**未发货")
        total += qty
    lines += ["", f"> 合计：**{fmt_qty(total)} 吨**"]
    return "\n".join(lines)

def send_wecom(webhook, content):
    webhook = webhook.strip()
    if not webhook.startswith("https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key="):
        raise ValueError("Webhook 地址格式不正确，应为企业微信群机器人 webhook 地址。")
    body = json.dumps({"msgtype":"markdown","markdown":{"content":content}}, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(webhook, data=body, headers={"Content-Type":"application/json; charset=utf-8"}, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=15) as resp: raw = resp.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as e:
        raw = e.read().decode("utf-8", errors="replace"); raise RuntimeError(f"HTTP {e.code}：{raw}") from e
    except urllib.error.URLError as e: raise RuntimeError(f"网络请求失败：{e.reason}") from e
    try: result = json.loads(raw)
    except json.JSONDecodeError as e: raise RuntimeError(f"企业微信返回无法解析：{raw[:300]}") from e
    if result.get("errcode") != 0: raise RuntimeError(f"企业微信错误 {result.get('errcode')}：{result.get('errmsg')}")

class App(tk.Tk):
    def __init__(self):
        super().__init__(); self.title(APP_NAME); self.geometry("850x590"); self.minsize(760,520)
        self.config_data = load_config(); self.selected_file=""; self.current_totals=[]; self.current_rule=""; self.current_unit_text=""; self.current_table=None; self.sending=False
        self.file_var=tk.StringVar(value="尚未选择文件"); self.status_var=tk.StringVar(value="等待选择 Excel 文件")
        self.webhook_var=tk.StringVar(value=self.config_data.get("webhook", DEFAULT_WEBHOOK)); self.unit_var=tk.StringVar(value=self.config_data.get("unit_mode","自动识别"))
        self._build_ui(); self._log("程序已启动。请选择销售/发货 Excel 文件。")
    def _build_ui(self):
        outer=ttk.Frame(self,padding=10); outer.pack(fill="both",expand=True)
        nb=ttk.Notebook(outer); nb.pack(fill="both",expand=True)
        self.main_tab=ttk.Frame(nb,padding=12); self.settings_tab=ttk.Frame(nb,padding=12); nb.add(self.main_tab,text="发货推送"); nb.add(self.settings_tab,text="设置")
        top=ttk.LabelFrame(self.main_tab,text="Excel 文件",padding=10); top.pack(fill="x")
        ttk.Entry(top,textvariable=self.file_var,state="readonly").pack(side="left",fill="x",expand=True,padx=(0,8)); ttk.Button(top,text="选择 Excel 文件",command=self.choose_file).pack(side="left")
        options=ttk.Frame(self.main_tab); options.pack(fill="x",pady=(10,6)); ttk.Label(options,text="数量单位：").pack(side="left")
        self.unit_combo=ttk.Combobox(options,textvariable=self.unit_var,values=("自动识别","kg（自动换算为吨）","吨"),state="readonly",width=20); self.unit_combo.pack(side="left"); self.unit_combo.bind("<<ComboboxSelected>>",lambda _e:self.recalculate_if_possible())
        ttk.Label(options,textvariable=self.status_var).pack(side="right")
        log_box=ttk.LabelFrame(self.main_tab,text="日志 / 推送预览",padding=8); log_box.pack(fill="both",expand=True,pady=(6,10))
        self.log=ScrolledText(log_box,wrap="word",font=("Microsoft YaHei UI",10),state="disabled"); self.log.pack(fill="both",expand=True)
        bottom=ttk.Frame(self.main_tab); bottom.pack(fill="x"); self.reject_btn=ttk.Button(bottom,text="拒绝 / 清空",command=self.reject_current); self.reject_btn.pack(side="right")
        self.confirm_btn=ttk.Button(bottom,text="确定并推送",command=self.confirm_push,state="disabled"); self.confirm_btn.pack(side="right",padx=(0,8))
        wb=ttk.LabelFrame(self.settings_tab,text="企业微信机器人 Webhook",padding=12); wb.pack(fill="x")
        ttk.Label(wb,text="默认已填写当前企业微信群机器人地址。修改后点击‘保存设置’，下次打开会继续使用。",wraplength=720).pack(anchor="w",pady=(0,8)); ttk.Entry(wb,textvariable=self.webhook_var).pack(fill="x")
        btns=ttk.Frame(wb); btns.pack(fill="x",pady=(10,0)); ttk.Button(btns,text="恢复默认 Webhook",command=self.restore_default_webhook).pack(side="left"); ttk.Button(btns,text="保存设置",command=self.save_settings).pack(side="right")
        ib=ttk.LabelFrame(self.settings_tab,text="本地配置",padding=12); ib.pack(fill="x",pady=(12,0)); ttk.Label(ib,text=f"配置文件：{app_config_path()}",wraplength=720).pack(anchor="w"); ttk.Label(ib,text="Webhook 只保存在本机配置文件中；不要把包含真实 key 的配置文件上传到公开仓库。",wraplength=720).pack(anchor="w",pady=(6,0))
    def _log(self,text):
        self.log.configure(state="normal"); self.log.insert("end",f"[{time.strftime('%H:%M:%S')}] {text}\n"); self.log.see("end"); self.log.configure(state="disabled")
    def _clear_log(self): self.log.configure(state="normal"); self.log.delete("1.0","end"); self.log.configure(state="disabled")
    def choose_file(self):
        path=filedialog.askopenfilename(title="选择销售/发货 Excel 文件",filetypes=[("Excel 文件","*.xlsx *.xlsm"),("所有文件","*.*")])
        if not path: self._log("已取消选择文件。"); return
        self.selected_file=path; self.file_var.set(path); self._clear_log(); self._log(f"已选择：{os.path.basename(path)}"); self.status_var.set("正在读取…"); self.confirm_btn.configure(state="disabled"); self.after(10,self._parse_selected)
    def _parse_selected(self):
        try:
            table=read_table(self.selected_file); self.current_table=table; self._log(f"工作表：{table['sheet']}"); self._log(f"识别型号列：{table['model_header']}"); self._log(f"识别数量列：{table['qty_header']}"); self._log(f"统计规则：{mode_text(table['mode'])}")
            if table["mode"]=="sales_as_remaining": self._log("警告：没有检测到‘已发货/出库数量’列，目前把销售/订单数量视为未发数量。")
            self.recalculate_if_possible()
        except Exception as e:
            self.current_table=None; self.current_totals=[]; self.status_var.set("读取失败"); self._log(f"读取失败：{e}"); messagebox.showerror("读取失败",str(e),parent=self)
    def recalculate_if_possible(self):
        if not self.current_table: return
        table=self.current_table; choice=self.unit_var.get(); explicit=detect_unit(table["qty_header"])
        if choice=="kg（自动换算为吨）": is_kg,unit_text=True,"kg（已换算为吨）"
        elif choice=="吨": is_kg,unit_text=False,"吨"
        elif explicit=="kg": is_kg,unit_text=True,"kg（表头识别，已换算为吨）"
        elif explicit=="ton": is_kg,unit_text=False,"吨（表头识别）"
        else: is_kg,unit_text=True,"kg（表头未注明，自动模式默认按 kg）"; self._log("数量列未注明单位：自动模式暂按 kg 读取。若原表单位是吨，请在上方切换为‘吨’。")
        self.current_totals=summarize(table["rows"],is_kg); self.current_rule=mode_text(table["mode"]); self.current_unit_text=unit_text
        self._log("—— 当前待发汇总 ——"); total=0.0
        for model,qty in self.current_totals: self._log(f"{model}：剩余 {fmt_qty(qty)} 吨未发货"); total+=qty
        self._log(f"合计：{fmt_qty(total)} 吨")
        if self.current_totals: self.status_var.set(f"已读取 {len(self.current_totals)} 个型号，合计 {fmt_qty(total)} 吨"); self.confirm_btn.configure(state="normal")
        else: self.status_var.set("没有大于 0 的待发数量"); self.confirm_btn.configure(state="disabled")
    def reject_current(self):
        if self.sending: return
        self.selected_file=""; self.current_table=None; self.current_totals=[]; self.current_rule=""; self.current_unit_text=""; self.file_var.set("尚未选择文件"); self.status_var.set("已拒绝本次数据，等待重新选择"); self.confirm_btn.configure(state="disabled"); self._clear_log(); self._log("已拒绝并清空当前数据，未进行推送。")
    def confirm_push(self):
        if self.sending or not self.selected_file or not self.current_totals: return
        webhook=self.webhook_var.get().strip()
        if not webhook: messagebox.showwarning("未设置 Webhook","请先在‘设置’标签页填写企业微信 Webhook。",parent=self); return
        total=sum(q for _,q in self.current_totals)
        if not messagebox.askokcancel("确认推送",f"将向企业微信推送 {len(self.current_totals)} 个型号，合计 {fmt_qty(total)} 吨。\n\n是否继续？",parent=self): self._log("用户取消了本次推送。"); return
        content=build_markdown(self.selected_file,self.current_totals,self.current_rule,self.current_unit_text); self.sending=True; self.confirm_btn.configure(state="disabled"); self.reject_btn.configure(state="disabled"); self.status_var.set("正在推送…"); self._log("正在推送到企业微信…")
        def worker():
            try: send_wecom(webhook,content)
            except Exception as e: self.after(0,lambda:self._push_done(False,str(e)))
            else: self.after(0,lambda:self._push_done(True,""))
        threading.Thread(target=worker,daemon=True).start()
    def _push_done(self,success,error):
        self.sending=False; self.reject_btn.configure(state="normal"); self.confirm_btn.configure(state="normal" if self.current_totals else "disabled")
        if success: self.status_var.set("推送成功，可继续选择下一份文件"); self._log("推送成功。窗口保持打开，可继续处理下一份 Excel。"); messagebox.showinfo("推送成功","已成功推送到企业微信群机器人。",parent=self)
        else: self.status_var.set("推送失败"); self._log(f"推送失败：{error}"); messagebox.showerror("推送失败",error,parent=self)
    def save_settings(self):
        webhook=self.webhook_var.get().strip()
        if not webhook: messagebox.showwarning("设置","Webhook 不能为空。",parent=self); return
        data={"webhook":webhook,"unit_mode":self.unit_var.get()}
        try: save_config(data); self.config_data=data; messagebox.showinfo("设置","设置已保存。",parent=self)
        except Exception as e: messagebox.showerror("保存失败",str(e),parent=self)
    def restore_default_webhook(self): self.webhook_var.set(DEFAULT_WEBHOOK)

def main(): App().mainloop()
if __name__ == "__main__": main()
