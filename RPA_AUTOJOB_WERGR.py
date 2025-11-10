# ================================== GR RPA (FINAL, ROBUST EXPORT + TEMP ATTACH) ==================================
import os, time, smtplib, traceback, subprocess, psutil
import win32com.client
from datetime import datetime, timedelta
from pathlib import Path
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.utils import formataddr
import re, sys
# ------------------------------ CONFIG ------------------------------
SAP_SERVER   = "03.SAP S/4 HANA - PRD"
SAP_USER     = "MCP_ADMIN2"
SAP_PASS     = "P@SsWord_MCP_ADMIN2"
SAP_CLIENT   = "900"
SAP_LANGUAGE = "EN"
SAP_EXE_PATH = r"C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe"

TABLE_NAME   = "ZTMCPMM0113"
VARIANT_NAME = "IS_JOB"
VARIANT_USER = "MCP_ADMIN2"

TEMP_DIR = Path(os.getenv("TEMP_DIR", r"C:\TEMP"))
TEMP_DIR.mkdir(parents=True, exist_ok=True)

MAX_RETRIES      = 4
RETRY_SLEEP_SECS = 30 * 60  # 30 นาที

# ------------------------------ EMAIL ------------------------------
SMTP_HOST    = os.getenv("SMTP_HOST", "10.236.36.206")
SMTP_PORT    = int(os.getenv("SMTP_PORT", "25"))
SMTP_USE_TLS = True
SMTP_USER    = os.getenv("SMTP_USER", "")
SMTP_PASS    = os.getenv("SMTP_PASS", "")

MAIL_FROM = os.getenv("MAIL_FROM", "suppakit.s@mcp.meap.com")
MAIL_TO   = [x.strip() for x in os.getenv("MAIL_TO", "suppakit.s@mcp.meap.com").split(",") if x.strip()]
# MAIL_CC   = [x.strip() for x in os.getenv("MAIL_CC", "").split(",") if x.strip()]


# def email_success_html(elapsed, files):
#     lis = "".join(f"<li>{Path(p).name}</li>" for p in files)
#     return f"""<html><body style="font-family:Segoe UI,Arial">
# <div style="background:#21a366;color:#fff;padding:10px;font-weight:700">✅ GR RPA — Export Success</div>
# <table><tr><td><b>⏱ Duration</b></td><td>{elapsed:.1f} s</td></tr>
# <tr><td><b>🕒 Time</b></td><td>{datetime.now():%Y-%m-%d %H:%M:%S}</td></tr></table>
# <p><b>Files:</b></p><ul>{lis}</ul><hr><small>🤖 RPA_GR Bot</small></body></html>"""

# def email_error_html(err, elapsed, tries, reason=None):
#     return f"""<html><body style="font-family:Segoe UI,Arial">
# <div style="background:#d9534f;color:#fff;padding:10px;font-weight:700">❌ GR RPA — Failed</div>
# <table><tr><td><b>⏱ Duration</b></td><td>{elapsed:.1f} s</td></tr>
# <tr><td><b>🧪 Attempts</b></td><td>{tries}/{MAX_RETRIES}</td></tr>
# <tr><td><b>🕒 Time</b></td><td>{datetime.now():%Y-%m-%d %H:%M:%S}</td></tr></table>
# {f"<p><b>Reason:</b> {reason}</p>" if reason else ""}
# <pre style="background:#fee;border:1px solid #d88;padding:8px;white-space:pre-wrap">{err}</pre>
# <hr><small>🤖 RPA_GR Bot</small></body></html>"""

def _exe_dir() -> Path:
    """ตำแหน่งโฟลเดอร์ที่รันอยู่ (รองรับ .py และ .exe onefile)"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent

def _dedup_preserve_case(emails: list[str]) -> list[str]:
    out, seen = [], set()
    for e in emails:
        e2 = e.strip()
        if not e2:
            continue
        k = e2.lower()
        if k not in seen:
            seen.add(k)
            out.append(e2)
    return out

def _load_cc_from_file(filename="CC_Mail_list.txt") -> list[str]:
    """อ่าน CC จากไฟล์ที่วาง 'ที่เดียวกับ .exe/.py' รองรับคั่นด้วย , ; เว้นวรรค/ขึ้นบรรทัด"""
    path = _exe_dir() / filename
    if not path.exists():
        return []
    txt = path.read_text(encoding="utf-8", errors="ignore")
    # แยกด้วย , ; เว้นวรรค และขึ้นบรรทัด
    parts = re.split(r"[,\s;]+", txt)
    # กรองเฉพาะที่มี '@'
    emails = [p.strip() for p in parts if "@" in p]
    return _dedup_preserve_case(emails)

    
MAIL_CC = _load_cc_from_file("CC_Mail_list.txt")

def email_success_html(elapsed, files):
    # ใช้ข้อความตามต้องการ + ใส่ชื่อไฟล์แนบไว้เผื่ออ้างอิง
    lis = "".join(f"<li>{Path(p).name}</li>" for p in (files or []))
    return f"""<html><body style="font-family:Segoe UI,Arial; font-size:14px; color:#1f2937">
<p>Dear All Concerned,</p>

<p>These are GR result of SAP job.<br>
Please find the detail in attachment.</p>

{("<p><b>Attachment:</b></p><ul>"+lis+"</ul>") if lis else ""}

<p>Best Regards,<br>
RPA_AutoJob</p>
</body></html>"""


def email_error_html(err, elapsed, tries, reason=None):
    # เวอร์ชันแจ้งล้มเหลว (ภาษาตามแบบเดียวกัน)
    return f"""<html><body style="font-family:Segoe UI,Arial; font-size:14px; color:#1f2937">
<p>Dear All Concerned,</p>

<p><b>GR result export of SAP job failed.</b><br>
Please check the details below.</p>

<p><b>Attempts:</b> {tries} / {MAX_RETRIES}<br>
<b>Reason:</b> {reason or "N/A"}<br>
<b>Time:</b> {datetime.now():%Y-%m-%d %H:%M:%S}</p>

<pre style="background:#f8f8f8; border:1px solid #e5e7eb; padding:10px; white-space:pre-wrap">{err}</pre>

<p>Best Regards,<br>
RPA_AutoJob</p>
</body></html>"""

def send_mail(subject: str, body_html: str, attachments: list[str] | None):
    # 1) รวม To/CC จาก ENV + ไฟล์ แล้ว "ตัดซ้ำ"
    to_list  = _dedup_preserve_case(MAIL_TO)
    cc_env   = _dedup_preserve_case(MAIL_CC)
    cc_file  = _load_cc_from_file("CC_Mail_list.txt")
    cc_all   = _dedup_preserve_case(cc_env + cc_file)

    # 2) ไม่ให้ CC ซ้ำกับ To (ทั้ง header และ rcpts)
    to_set = {e.lower() for e in to_list}
    cc_list = [e for e in cc_all if e.lower() not in to_set]

    # 3) สร้างข้อความ
    msg = MIMEMultipart("mixed")
    alt = MIMEMultipart("alternative")
    msg.attach(alt)
    msg["From"] = formataddr(("RPA_GR", MAIL_FROM))
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg["Subject"] = subject
    alt.attach(MIMEText(body_html, "html", "utf-8"))

    # 4) แนบไฟล์ (ถ้ามี)
    for p in (attachments or []):
        try:
            with open(p, "rb") as f:
                part = MIMEApplication(f.read())
            part.add_header("Content-Disposition", "attachment", filename=Path(p).name)
            msg.attach(part)
        except Exception as e:
            print(f"⚠️ แนบไฟล์ไม่ได้: {p} : {e}")

    # 5) ปลายทางส่งจริง = To ∪ Cc (ตัดซ้ำ)
    rcpts = _dedup_preserve_case(to_list + cc_list)

    # 6) ส่ง
    s = smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=60)
    try:
        if SMTP_USE_TLS:
            s.starttls()
        if SMTP_USER:
            s.login(SMTP_USER, SMTP_PASS)
        s.sendmail(MAIL_FROM, rcpts, msg.as_string())
        print("📨 ส่งเมลเรียบร้อย")
    finally:
        s.quit()



# ------------------------------ HELPERS ------------------------------
def ensure_sap_running():
    for p in psutil.process_iter(['name']):
        n = (p.info.get('name') or "").lower()
        if n.startswith("saplogon"):
            print("✅ SAP GUI already running."); return
    print("🚀 Starting SAP GUI...")
    subprocess.Popen(SAP_EXE_PATH); time.sleep(6); print("✅ SAP GUI started.")

def get_session():
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    app = SapGuiAuto.GetScriptingEngine
    sess = None
    if app.Children.Count > 0:
        conn = app.Children(0)
        if conn.Children.Count > 0: sess = conn.Children(0)
    if sess is None:
        conn = app.OpenConnection(SAP_SERVER, True); time.sleep(3)
        sess = conn.Children(0)
    sess.findById("wnd[0]").maximize()
    return sess

def handle_multiple_logon_popup(session):
    try:
        w = session.findById("wnd[1]", False)
        if w and "Multiple Logon" in w.Text:
            w.findById("usr/radMULTI_LOGON_OPT2").select()
            w.findById("tbar[0]/btn[0]").press()
            time.sleep(1.0)
    except Exception:
        pass

def login_if_needed(session):
    try:
        if session.findById("wnd[0]/usr/txtRSYST-BNAME").Text == "":
            session.findById("wnd[0]/usr/txtRSYST-MANDT").Text = SAP_CLIENT
            session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = SAP_USER
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = SAP_PASS
            session.findById("wnd[0]/usr/txtRSYST-LANGU").Text = SAP_LANGUAGE
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(2.5)
            handle_multiple_logon_popup(session)
        else:
            print("✅ Session already logged in")
    except Exception as e:
        print("⚠️ Skip login:", e)

def wait_until(session, path, timeout=6.0, step=0.1, optional=False):
    end = time.time() + timeout
    while time.time() < end:
        try: return session.findById(path)
        except Exception: time.sleep(step)
    if optional: return None
    raise TimeoutError(f"Timeout waiting for {path}")

def _set_text_safe(ctrl, text:str):
    try: ctrl.Text = text; return True
    except Exception: pass
    try: ctrl.SetFocus(); ctrl.Text = text; return True
    except Exception: return False

def status_text(session) -> str:
    try: return session.findById("wnd[0]/sbar").Text.strip()
    except Exception: return ""

def close_sap_all(session=None):
    try:
        if session:
            try:
                session.findById("wnd[0]/tbar[0]/okcd").Text = "/nex"
                session.findById("wnd[0]").sendVKey(0)
            except Exception: pass
        for p in psutil.process_iter(['pid','name']):
            n = (p.info.get('name') or "").lower()
            if n.startswith("excel") or n.startswith("saplogon"):
                try: p.terminate()
                except Exception: pass
    except Exception as e:
        print("⚠️ close_sap_all:", e)

# ------------------------------ SE16N STEPS ------------------------------
def _set_table_name(session, table_name:str):
    fld = wait_until(session, "wnd[0]/usr/ctxtGD-TAB", 3.0, optional=True)
    if not fld: raise RuntimeError("ไม่พบช่อง 'Table' (ctxtGD-TAB)")
    fld.Text = table_name

def choose_variant(session, variant_name, user="*"):
    session.findById("wnd[0]").sendVKey(6)  # F6
    wait_until(session, "wnd[1]", 6)
    vf = session.findById("wnd[1]/usr/ctxtGS_SE16N_LT-NAME")
    uf = wait_until(session, "wnd[1]/usr/txtGS_SE16N_LT-UNAME", 1.0, optional=True)
    vf.Text = variant_name
    if uf: uf.Text = user or "*"
    try:
        vf.SetFocus(); session.findById("wnd[1]").sendVKey(4)
        _hitlist_select_variant(session, variant_name, (user if user!="*" else None))
    except Exception:
        pass
    _press_ok_popup(session)

def _press_ok_popup(session):
    try: session.findById("wnd[1]").sendVKey(0); return
    except Exception: pass
    for b in ("wnd[1]/tbar[0]/btn[0]","wnd[1]/tbar[0]/btn[2]"):
        btn = wait_until(session, b, 1.0, optional=True)
        if btn: btn.press(); return
    raise RuntimeError("ไม่พบปุ่ม OK ใน popup")

def _hitlist_select_variant(session, vname, user=None):
    win = None
    for w in ("wnd[2]","wnd[1]"):
        try: win = session.findById(w); break
        except Exception: pass
    if not win: return
    candidates = [
        "usr/cntlALV_CONTAINER_1/shellcont/shell",
        "usr/cntlGRID1/shellcont/shell",
        "usr/tblSAPLALDB","usr/tbl*"
    ]
    grid = None
    for p in candidates:
        try:
            if p.endswith("*"):
                ch = win.Children
                for i in range(ch.Count):
                    c = ch.Item(i)
                    if c.Id.startswith(f"{win.Id}/{p[:-1]}"):
                        grid = c; break
            else:
                grid = win.findById(f"{win.Id}/{p}")
            if grid: break
        except Exception: pass
    if not grid:
        try: win.sendVKey(0); return
        except Exception: return

    def get(r, cols):
        for c in cols:
            try: return str(grid.GetCellValue(r, c)).strip()
            except Exception: pass
        return ""

    var_cols  = ["VARIANT","VARNAME","NAME","LTNAME",0,1]
    user_cols = ["USER","UNAME","AENAM",3]
    try: rows = grid.RowCount
    except Exception: rows = 1500

    pick = None
    for r in range(rows):
        val = get(r, var_cols)
        if val and val.lower() == vname.lower():
            u = get(r, user_cols) if user else ""
            if (not user) or (u.lower()==user.lower()):
                pick = r; break
    if pick is not None:
        grid.currentCellRow = pick
        grid.selectedRows   = str(pick)
        try: grid.doubleClickCurrentCell()
        except Exception: win.sendVKey(0)
    else:
        win.sendVKey(0)

def fill_running_datetime(session, minus_hours=1):
    now = datetime.now()
    run_date = now.strftime("%d.%m.%Y")  # ใช้วันที่จริง
    # run_date = "09.11.2025"  # ใช้วันที่จริง

    run_time = (now - timedelta(hours=minus_hours)).strftime("%H:%M:%S")
    # run_time = "20:10:00"

    base = "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC"

    date_edit = wait_until(session, f"{base}/ctxtGS_SELFIELDS-LOW[2,1]", 3.0, optional=True)
    if not date_edit: raise RuntimeError("ไม่พบช่อง Running Date (LOW[2,1])")
    if not _set_text_safe(date_edit, run_date):
        if not _f4_fill(session, date_edit, run_date):
            raise RuntimeError("ใส่ค่า Running Date ไม่สำเร็จ")

    time_edit = wait_until(session, f"{base}/ctxtGS_SELFIELDS-LOW[2,2]", 3.0, optional=True)
    if not time_edit: raise RuntimeError("ไม่พบช่อง Running Time (LOW[2,2])")
    if not _set_text_safe(time_edit, run_time):
        if not _f4_fill(session, time_edit, run_time):
            raise RuntimeError("ใส่ค่า Running Time ไม่สำเร็จ")

    print(f"📝 Set Running Date={run_date} | Running Time={run_time}")

def _f4_fill(session, ctrl, value):
    try: ctrl.SetFocus()
    except Exception: pass
    try: session.findById("wnd[0]").sendVKey(4)
    except Exception: pass
    win = None
    for w in ("wnd[2]","wnd[1]"):
        try: win = session.findById(w); break
        except Exception: pass
    if not win: return False

    def first_input(node):
        try:
            ch = node.Children
            for i in range(ch.Count):
                c = ch.Item(i)
                t = getattr(c,"Type","")
                if t.endswith("GuiCTextField") or t.endswith("GuiTextField"): return c
                sub = first_input(c)
                if sub: return sub
        except Exception: return None
        return None

    inp = first_input(win)
    if not inp: return False
    if not _set_text_safe(inp, value): return False
    try: win.sendVKey(0)
    except Exception:
        for bb in ("tbar[0]/btn[0]","tbar[0]/btn[2]"):
            b = wait_until(session, f"{win.Id}/{bb}", 0.8, optional=True)
            if b: b.press(); break
    return True

def execute_report(session):
    try: session.findById("wnd[0]/tbar[1]/btn[8]").press()
    except Exception: session.findById("wnd[0]").sendVKey(8)

# ------------------------------ ALV helpers ------------------------------
def _is_alv_like(node) -> bool:
    t = (getattr(node, "Type", "") or "")
    return (t.endswith(("GuiShell","GuiGridView","GuiTableControl","GuiTree"))
            or hasattr(node, "contextMenu")
            or hasattr(node, "pressToolbarContextButton")
            or hasattr(node, "RowCount"))

def _deep_find_alv_node(root):
    try:
        ch = root.Children
        for i in range(ch.Count):
            c = ch.Item(i)
            if _is_alv_like(c): return c
            got = _deep_find_alv_node(c)
            if got: return got
    except Exception:
        return None
    return None

def _find_alv_grid(session):
    paths = [
        "wnd[0]/usr/cntlALV_CONTAINER_1/shellcont/shell",
        "wnd[0]/usr/cntlGRID1/shellcont/shell",
        "wnd[0]/usr/cntlGRID/shellcont/shell",
        "wnd[0]/usr/cntlCONTAINER/shellcont/shell",
        "wnd[0]/usr/cntlALV_CONTAINER/shellcont/shell",
    ]
    for p in paths:
        try:
            g = session.findById(p)
            if _is_alv_like(g): return g
        except Exception: pass
    try:
        usr = session.findById("wnd[0]/usr")
        return _deep_find_alv_node(usr)
    except Exception:
        return None

def _open_export_via_grid_context(session) -> bool:
    grid = _find_alv_grid(session)
    if not grid: return False
    try:
        try:
            grid.pressToolbarContextButton("&MB_EXPORT")
            for item in ("&XXL","&PC","SPREADSHEET"):
                try:
                    grid.selectContextMenuItem(item); return True
                except Exception: pass
        except Exception: pass
        try:
            try: grid.setCurrentCell(0,0)
            except Exception: pass
            grid.contextMenu()
            for item in ("&XXL","&PC","SPREADSHEET"):
                try:
                    grid.selectContextMenuItem(item); return True
                except Exception: pass
        except Exception: pass
    except Exception:
        return False
    return False

# ------------------------------ POPUP/WARN ------------------------------
def _press_continue_if_popup(session) -> bool:
    for w in ("wnd[2]","wnd[1]"):
        try:
            win = session.findById(w)
            for btn in ("tbar[0]/btn[0]","tbar[0]/btn[2]"):
                try:
                    win.findById(btn).press(); time.sleep(0.4); return True
                except Exception: pass
        except Exception: pass
    return False

def wait_for_alv_or_continue(session, timeout=30.0, step=0.5):
    end = time.time() + timeout
    while time.time() < end:
        g = _find_alv_grid(session)
        if g: return g
        if _press_continue_if_popup(session):
            time.sleep(0.5); continue
        try:
            txt = status_text(session)
            if "Numerous rows" in txt or "performance" in txt.lower():
                session.findById("wnd[0]").sendVKey(0); time.sleep(0.5); continue
        except Exception: pass
        time.sleep(step)
    print("⚠️ ไม่พบ ALV grid ชัดเจน แต่จะดำเนินการ Export ต่อไป")

# ------------------------------ FILE helpers ------------------------------
def _wait_file_stable(path: Path, min_size=2048, stable_secs=1.5, timeout=30.0) -> bool:
    path = Path(path); end = time.time() + timeout; last_mtime = None
    while time.time() < end:
        try:
            st = path.stat()
            if st.st_size >= min_size:
                m = st.st_mtime
                if last_mtime is not None and (time.time() - m) >= stable_secs: return True
                last_mtime = m
        except FileNotFoundError:
            pass
        time.sleep(0.25)
    return False

def _wait_for_new_xlsx(folder: Path|str, since_ts: float, timeout=60.0, stable_secs=1.5, min_size=1024) -> Path|None:
    folder = Path(folder); end = time.time() + timeout; last_seen = None
    while time.time() < end:
        try:
            cands = []
            for f in folder.glob("*.xlsx"):
                if f.name.startswith("~$"): continue
                st = f.stat()
                if st.st_mtime >= since_ts and st.st_size >= min_size:
                    cands.append((st.st_mtime, f))
            if cands:
                cands.sort(key=lambda x: x[0], reverse=True)
                newest = cands[0][1]
                st2 = newest.stat()
                if (time.time() - st2.st_mtime) >= stable_secs and st2.st_size >= min_size:
                    return newest
                last_seen = newest
        except FileNotFoundError:
            pass
        time.sleep(0.25)
    return last_seen

def _rename_with_retry(src: Path, dst: Path, tries=15, sleep=0.3) -> Path:
    src = Path(src); dst = Path(dst)
    for _ in range(tries):
        try:
            if src.resolve() == dst.resolve(): return dst
            dst.parent.mkdir(parents=True, exist_ok=True)
            os.replace(str(src), str(dst))
            return dst
        except Exception:
            time.sleep(sleep)
    return src

# ------------------------------ EXPORT ------------------------------
def export_alv_to_excel_and_return_paths(session) -> list[str]:
    time.sleep(1.0)
    _press_continue_if_popup(session)

    opened = False
    if _find_alv_grid(session) and _open_export_via_grid_context(session):
        opened = True; print("✅ Export via ALV context menu")

    if not opened:
        for b in ("wnd[0]/tbar[1]/btn[33]","wnd[0]/tbar[1]/btn[45]","wnd[0]/tbar[1]/btn[44]","wnd[0]/tbar[1]/btn[32]"):
            try: session.findById(b).press(); opened=True; print(f"✅ Export via toolbar: {b}"); break
            except Exception: pass

    if not opened:
        for mp in ("wnd[0]/mbar/menu[1]/menu[3]","wnd[0]/mbar/menu[1]/menu[2]","wnd[0]/mbar/menu[0]/menu[3]/menu[1]"):
            try: session.findById(mp).select(); opened=True; print(f"✅ Export via menu: {mp}"); break
            except Exception: pass

    if not opened:
        raise RuntimeError("❌ ไม่พบคำสั่ง Export")

    time.sleep(0.6)
    _press_continue_if_popup(session)
    _handle_export_format_selection(session)

    t0 = time.time()
    path = _handle_save_file_dialog_and_save(session, str(TEMP_DIR))
    print("💾 Saved (expected):", path)

    p_obj = Path(path)
    if not p_obj.exists():
        p_actual = _wait_for_new_xlsx(TEMP_DIR, since_ts=t0, timeout=60.0, stable_secs=1.5, min_size=1024)
        if p_actual and p_actual.exists():
            p_obj = p_actual
        else:
            lf = _latest_real_xlsx(str(TEMP_DIR))
            if lf: p_obj = Path(lf)
            else: raise RuntimeError("Export เสร็จแต่ไม่พบไฟล์ .xlsx ในโฟลเดอร์")

    _wait_file_stable(p_obj, min_size=2048, stable_secs=1.5, timeout=30.0)

    final_name = f"GR_WEB_Errorlist_{datetime.now():%Y%m%d_%H%M%S}.xlsx"
    p_final = _rename_with_retry(p_obj, Path(TEMP_DIR)/final_name, tries=15, sleep=0.3)
    return [str(p_final)]

def _handle_export_format_selection(session):
    for w in ("wnd[2]","wnd[1]"):
        try: win = session.findById(w)
        except Exception: continue
        for rid in ("usr/radRB_OTHERS","usr/rad-rb_spreadsheet","usr/radSPOPLI-XL"):
            try: win.findById(rid).select()
            except Exception: pass
        for cid in ("usr/cmbG_LISTBOX","usr/cmbSALV_BS_EXPORT-LIST","usr/cmbGD_TYPE"):
            try:
                cb = win.findById(cid)
                try: cb.key = "31"  # XLSX
                except Exception: pass
            except Exception: pass
        for ok in ("tbar[0]/btn[0]","tbar[0]/btn[2]"):
            try: win.findById(ok).press(); return
            except Exception: pass
    # ไม่มี dialog ก็ไปต่อ

def _handle_save_file_dialog_and_save(session, folder: str) -> str:
    """
    Dialog บันทึกไฟล์ (3 แบบ)
    - ไม่เปลี่ยนค่า File Name หากมีอยู่แล้ว
    - กดปุ่ม Generate ให้แน่นอน (ด้วยหลาย fallback)
    """
    def _export_filename(): return f"EXPORT_{datetime.now():%Y%m%d%H%M%S}.xlsx"

    def _click_button_by_text(win, keywords=("Generate","Replace","Extend","Save","OK","Continue","Export to")):
        try:
            ch = win.Children
            for i in range(ch.Count):
                c = ch.Item(i)
                if getattr(c,"Type","").endswith("GuiButton"):
                    txt = (getattr(c,"Text","") or "").strip()
                    if any(k.lower() in txt.lower() for k in keywords):
                        c.press(); return True
                if _click_button_by_text(c, keywords): return True
        except Exception: return False
        return False

    def _press_any_button(win):
        """fallback สุดท้าย: กดปุ่มแรกที่หาได้"""
        try:
            ch = win.Children
            for i in range(ch.Count):
                c = ch.Item(i)
                if getattr(c,"Type","").endswith("GuiButton"):
                    try: c.press(); return True
                    except Exception: pass
                if _press_any_button(c): return True
        except Exception: return False
        return False

    # หา window
    win = None; start = time.time()
    while time.time() - start < 20:
        for w in ("wnd[2]","wnd[1]"):
            try:
                tmp = session.findById(w)
                if hasattr(tmp,"Children") and tmp.Children.Count>0: win = tmp; break
            except Exception: pass
        if win: break
        time.sleep(0.3)
    if not win: raise RuntimeError("ไม่พบ Save File dialog")

    # --- รูปแบบ A: DY_PATH/DY_FILENAME ---
    try:
        p = win.findById(f"{win.Id}/usr/ctxtDY_PATH")
        f = win.findById(f"{win.Id}/usr/ctxtDY_FILENAME")
        p.Text = folder
        if not (getattr(f,"Text","") or "").strip():
            f.Text = _export_filename()
        # กด Generate/OK
        if not _click_button_by_text(win):
            try: win.sendVKey(0)
            except Exception: pass
        if not _click_button_by_text(win):
            for b in ("tbar[0]/btn[11]","tbar[0]/btn[0]"):
                try: win.findById(b).press(); break
                except Exception: pass
        return str(Path(folder) / (getattr(f,"Text","") or _export_filename()))
    except Exception:
        pass

    # สแกน input ทั้งหมด
    def walk_inputs(node):
        res = []
        try:
            ch = node.Children
            for i in range(ch.Count):
                c = ch.Item(i); t = getattr(c,"Type","")
                if t.endswith("GuiCTextField") or t.endswith("GuiTextField"): res.append(c)
                res += walk_inputs(c)
        except Exception: pass
        return res

    inputs = walk_inputs(win)

    # --- รูปแบบ B: 2 ช่อง (Directory + File Name) ---
    if len(inputs) >= 2:
        dir_inp, fn_inp = inputs[0], inputs[1]

        # ไม่ไปแก้ File Name ถ้ามีค่าอยู่แล้ว
        cur_name = (getattr(fn_inp,"Text","") or "").strip()
        if not cur_name:
            cur_name = _export_filename()
            _set_text_safe(fn_inp, cur_name)

        # ใส่ Directory (ถ้าว่าง/ผิด)
        _set_text_safe(dir_inp, folder)

        # พยายามกด Generate ด้วยชุด fallback
        if not _click_button_by_text(win, ("Generate","Save","OK","Continue","Replace","Extend")):
            try: win.sendVKey(0); time.sleep(0.2)
            except Exception: pass
        if not _click_button_by_text(win, ("Generate","Save","OK","Continue")):
            if not _press_any_button(win):
                for b in ("tbar[0]/btn[0]","tbar[0]/btn[11]"):
                    try: win.findById(b).press(); break
                    except Exception: pass

        return str(Path(folder) / Path(cur_name).name)

    # --- รูปแบบ C: 1 ช่อง (ใส่เป็นโฟลเดอร์) ---
    if len(inputs) == 1:
        file_inp = inputs[0]
        if not (getattr(file_inp,"Text","") or "").strip():
            _set_text_safe(file_inp, folder)

        if not _click_button_by_text(win, ("Export to","Save","OK","Continue","Replace")):
            try: win.sendVKey(0)
            except Exception: pass
        if not _click_button_by_text(win):
            for b in ("tbar[0]/btn[0]","tbar[0]/btn[11]"):
                try: win.findById(b).press(); break
                except Exception: pass

        # คืน path เดา (ไปจับจริงภายหลัง)
        return str(Path(folder) / f"EXPORT_{datetime.now():%Y%m%d%H%M%S}.xlsx")

    raise RuntimeError("ระบุพาธสำหรับบันทึกไฟล์ไม่สำเร็จใน Save File dialog")

def _guess_resulting_file(folder:str) -> Path:
    f = _latest_real_xlsx(folder)
    return Path(f) if f else Path(folder)/f"EXPORT_{datetime.now():%Y%m%d%H%M%S}.xlsx"

def _latest_real_xlsx(folder:str, min_size=1024, stable=1.2) -> str|None:
    p = Path(folder); now = time.time(); cands = []
    for f in p.glob("*.xlsx"):
        if f.name.startswith("~$"): continue
        try: st = f.stat()
        except FileNotFoundError: continue
        if st.st_size < min_size: continue
        if (now - st.st_mtime) < stable: continue
        cands.append((st.st_mtime, f))
    if not cands: return None
    cands.sort(key=lambda x: x[0], reverse=True)
    return str(cands[0][1])

def _collect_recent_xlsx(folder:Path, primary:Path|str|None) -> list[str]:
    primary = Path(primary) if primary else None
    files = []; cutoff = time.time() - 120
    for f in folder.glob("*.xlsx"):
        if f.name.startswith("~$"): continue
        try:
            if f.stat().st_mtime >= cutoff: files.append(str(f))
        except FileNotFoundError: pass
    if not files:
        lf = _latest_real_xlsx(str(folder))
        if lf: files = [lf]
    if primary and str(primary) in files:
        files.remove(str(primary)); files.insert(0, str(primary))
    return files

# ------------------------------ MAIN ------------------------------
def main():
    start = datetime.now()
    tries = 0
    success = False
    reason = None
    last_err = None
    attachments: list[str] = []

    while tries < MAX_RETRIES and not success:
        tries += 1
        try:
            print(f"\n===== Attempt {tries}/{MAX_RETRIES} =====")
            ensure_sap_running()
            sess = get_session()
            login_if_needed(sess)

            sess.StartTransaction("SE16N"); time.sleep(1.0)
            _set_table_name(sess, TABLE_NAME)
            choose_variant(sess, VARIANT_NAME, VARIANT_USER)
            fill_running_datetime(sess, minus_hours=1.5)
            execute_report(sess)

            wait_for_alv_or_continue(sess, timeout=45.0)

            time.sleep(0.6)
            sbar = status_text(sess)
            print("ℹ️ Status:", sbar)
            if "No value" in sbar or "No values" in sbar:
                reason = "No values found"
                print("🟥 No values found → ปิด SAP และรอ 30 นาที")
                close_sap_all(sess); time.sleep(RETRY_SLEEP_SECS); continue

            attachments = export_alv_to_excel_and_return_paths(sess)
            success = True
            close_sap_all(sess)

        except Exception as e:
            last_err = f"{e}\n\nTraceback:\n{traceback.format_exc()}"
            print("❌ ERROR:", last_err)
            close_sap_all()
            if tries < MAX_RETRIES:
                print("⏳ รอ 60 วินาทีแล้วลองใหม่..."); time.sleep(60)

    elapsed = (datetime.now() - start).total_seconds()
    try:
        if success:
            html = email_success_html(elapsed, attachments)
            send_mail(f"[SUCCESS] WEB GR Result Export — {datetime.now():%Y-%m-%d}", html, attachments)
        else:
            html = email_error_html(last_err or "Unknown error", elapsed, tries, reason)
            send_mail(f"[FAILED] WEB GR Result Export — {datetime.now():%Y-%m-%d}", html, None)
    finally:
        for p in attachments:
            try:
                Path(p).unlink(missing_ok=True)
                print(f"🧹 Deleted temp: {p}")
            except Exception as e:
                print(f"⚠️ ลบไฟล์ไม่สำเร็จ {p} : {e}")

if __name__ == "__main__":
    main()
