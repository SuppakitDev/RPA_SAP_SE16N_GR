import win32com.client
import time
from datetime import datetime, timedelta
import os
import subprocess
import psutil
import glob
import pyodbc
import smtplib
import pandas as pd
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.utils import formataddr
import traceback
import mimetypes
from pathlib import Path
 
# ------------------------------------------------------------
# CONFIG
# ------------------------------------------------------------
SAP_SERVER   = "03.SAP S/4 HANA - PRD"    # ต้องตรงกับชื่อใน SAP Logon
SAP_USER     = "MCP_ADMIN2"
SAP_PASS     = "P@SsWord_MCP_ADMIN2"
SAP_CLIENT   = "900"
SAP_LANGUAGE = "EN"
BASE_PATH    = r"\\10.236.36.212\FTP_File\MCP\900\Inbound\MM\IF_GR_REF_PO\STOPlan"
SAP_EXE_PATH = r"C:\Program Files\SAP\FrontEnd\SAPgui\saplogon.exe"

# ------------------------------------------------------------
# EMAIL CONFIG
# ------------------------------------------------------------
SMTP_HOST    = os.getenv("SMTP_HOST", "10.236.36.206")
SMTP_PORT    = int(os.getenv("SMTP_PORT", "25"))
SMTP_USE_TLS = True  # ส่วนใหญ่ O365 ใช้ TLS:587
SMTP_USER    = os.getenv("SMTP_USER", "")
SMTP_PASS    = os.getenv("SMTP_PASS", "")

MAIL_FROM    = os.getenv("MAIL_FROM", "suppakit.s@mcp.meap.com")
MAIL_TO      = [addr.strip() for addr in os.getenv("MAIL_TO", "suppakit.s@mcp.meap.com").split(",")]
MAIL_CC      = ["natthadech.r@mcp.meap.com"]  # ใส่เพิ่มได้
# MAIL_CC      = []  # ใส่เพิ่มได้

def email_success_html(elapsed, export_folder):
    return f"""
<html>
<body style="font-family:Segoe UI,Arial;">
<div style="background:#21a366;padding:12px;color:white;font-size:18px;font-weight:bold;">
✅ STO RPA — Export Success
</div>

<p>ระบบ RPA ทำงานเสร็จสมบูรณ์ ✅</p>

<table style="border-collapse:collapse;">
<tr><td><b>📂 Folder:</b></td><td>{export_folder}</td></tr>
<tr><td><b>⏱️ Duration:</b></td><td>{elapsed:.1f} seconds</td></tr>
<tr><td><b>🕒 Time:</b></td><td>{datetime.now():%Y-%m-%d %H:%M:%S}</td></tr>
</table>

<br>

<a href="file:///{export_folder.replace("\\", "/")}" 
style="background:#21a366;color:white;padding:10px 15px;text-decoration:none;border-radius:5px;">
📂 Open Folder
</a>

<hr>
<p style="font-size:12px;color:gray;">
🤖 RPA_STO Bot<br>
Auto-generated email — please do not reply
</p>
</body>
</html>
"""

def email_error_html(error_text, elapsed):
    return f"""
<html>
<body style="font-family:Segoe UI,Arial;">
<div style="background:#d9534f;padding:12px;color:white;font-size:18px;font-weight:bold;">
❌ STO RPA — Export Failed
</div>

<p>ระบบ RPA ล้มเหลวในการทำงาน ❌</p>

<table style="border-collapse:collapse;">
<tr><td><b>⏱️ Duration:</b></td><td>{elapsed:.1f} seconds</td></tr>
<tr><td><b>🕒 Time:</b></td><td>{datetime.now():%Y-%m-%d %H:%M:%S}</td></tr>
</table>

<br>
<b>⚠️ Error Detail:</b>
<pre style="background:#f8d7da;padding:10px;border:1px solid #d9534f;white-space:pre-wrap;">
{error_text}
</pre>

<hr>
<p style="font-size:12px;color:gray;">
🤖 RPA_STO Bot<br>
Auto-generated email — please do not reply
</p>
</body>
</html>
"""


def send_mail(subject: str, body_html: str):
    msg = MIMEMultipart("alternative")
    msg["From"] = formataddr(("RPA_STO", MAIL_FROM))
    msg["To"] = ", ".join(MAIL_TO)
    if MAIL_CC:
        msg["Cc"] = ", ".join(MAIL_CC)
    msg["Subject"] = subject

    # HTML part only
    msg.attach(MIMEText(body_html, "html", "utf-8"))

    rcpts = MAIL_TO + MAIL_CC

    server = smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30)
    try:
        if SMTP_USE_TLS:
            server.starttls()
        if SMTP_USER:
            server.login(SMTP_USER, SMTP_PASS)
        server.sendmail(MAIL_FROM, rcpts, msg.as_string())
        print("📨 ส่งอีเมลเรียบร้อย (no attachment)")
    finally:
        server.quit()


def is_temp_office_file(p: Path) -> bool:
    return p.name.startswith("~$")

def latest_real_xlsx(folder: str | Path, pattern: str = "*.xlsx", min_size_bytes: int = 1024, stable_secs: float = 1.5) -> str | None:
    """คืน path ของไฟล์ .xlsx ล่าสุดที่ 'ไม่ใช่' ~$, ขนาด > min_size และเวลาแก้ไขนิ่งแล้ว stable_secs วินาที"""
    folder = Path(folder)
    cands = []
    now = time.time()
    for f in folder.glob(pattern):
        if not f.is_file():
            continue
        if is_temp_office_file(f):
            continue
        try:
            st = f.stat()
        except FileNotFoundError:
            continue
        # ข้ามไฟล์เล็กจิ๋ว/ยังเขียนอยู่
        if st.st_size < min_size_bytes:
            continue
        # ต้องนิ่งมาสักพัก (กันเคสเพิ่งปิด handle)
        if (now - st.st_mtime) < stable_secs:
            continue
        cands.append((st.st_mtime, f))
    if not cands:
        return None
    cands.sort(key=lambda x: x[0], reverse=True)
    return str(cands[0][1])

def wait_for_real_xlsx(folder: str | Path, timeout: float = 60.0) -> str | None:
    """วนรอจนกว่าจะเจอไฟล์จริง (ไม่ใช่ ~$) ตามเกณฑ์ด้านบน หรือหมดเวลา"""
    end = time.time() + timeout
    while time.time() < end:
        p = latest_real_xlsx(folder)
        if p:
            return p
        time.sleep(0.5)
    return None


# ------------------------------------------------------------
# PREPARE EXPORT PATH (AUTO BY DATE)
# ------------------------------------------------------------
# today_str = datetime.now().strftime("%Y-%m-%d")
# EXPORT_FOLDER = os.path.join(BASE_PATH, today_str)
# os.makedirs(EXPORT_FOLDER, exist_ok=True)
# EXPORT_PATH = os.path.join(EXPORT_FOLDER, "STO_Report.xlsx")
 
# print(f"📁 Export folder ready: {EXPORT_FOLDER}")
# print(f"📄 Export file will be saved as: {EXPORT_PATH}")
# ------------------------------------------------------------
# PREPARE EXPORT PATH (FIXED FOLDER: ...\File)
# ------------------------------------------------------------
EXPORT_FOLDER = os.path.join(BASE_PATH, "File")
os.makedirs(EXPORT_FOLDER, exist_ok=True)

print(f"📁 Export folder ready: {EXPORT_FOLDER}")
print(f"📄 Export file will be saved into this folder (using SAP's default filename).")
 
# ------------------------------------------------------------
# HELPER FUNCTIONS
# ------------------------------------------------------------
def ensure_sap_running():
    for proc in psutil.process_iter(['name']):
        if proc.info['name'] and proc.info['name'].lower().startswith("saplogon"):
            print("✅ SAP GUI already running.")
            return
    print("🚀 Starting SAP GUI...")
    subprocess.Popen(SAP_EXE_PATH)
    time.sleep(6)
    print("✅ SAP GUI started successfully.")
 
def handle_multiple_logon_popup(session):
    """ตรวจจับ popup “Multiple Logon” แล้วเลือก Option 2 (Continue with this login)"""
    try:
        wnd1 = session.findById("wnd[1]", False)
        if wnd1 and "Multiple Logon" in wnd1.text:
            print("⚠️ พบหน้าต่าง Multiple Logon — เลือก Option 2 และดำเนินการต่อ...")
            wnd1.findById("usr/radMULTI_LOGON_OPT2").select()
            wnd1.findById("usr/radMULTI_LOGON_OPT2").setFocus()
            wnd1.findById("tbar[0]/btn[0]").press()
            print("✅ เลือก Continue with this login สำเร็จ")
            time.sleep(2)
    except Exception:
        pass
 
# ------------------------------------------------------------
# START SAP GUI
# ------------------------------------------------------------
ensure_sap_running()
start_time = datetime.now()
export_ok = False
last_error = None
# ------------------------------------------------------------
# CONNECT TO SAP
# ------------------------------------------------------------
try:
    print("Connecting to SAP...")
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
    except Exception:
        print("❌ ยังไม่พบ SAP GUI COM Object — รอ 5 วิแล้วลองอีกครั้ง...")
        time.sleep(5)
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
    
    session = None
    if application.Children.Count > 0:
        connection = application.Children(0)
        if connection.Children.Count > 0:
            session = connection.Children(0)
            print("✅ ใช้ SAP Session เดิมที่เปิดอยู่")
    
    if session is None:
        print("🔄 ไม่พบ session ที่เปิดอยู่ — กำลังเปิด SAP Logon connection...")
        connection = application.OpenConnection(SAP_SERVER, True)
        time.sleep(3)
        session = connection.Children(0)
        print("✅ เปิด connection ใหม่สำเร็จ")
    
    # ------------------------------------------------------------
    # LOGIN
    # ------------------------------------------------------------
    try:
        session.findById("wnd[0]").maximize()
        if session.findById("wnd[0]/usr/txtRSYST-BNAME").Text == "":
            print("🔐 Logging in...")
            session.findById("wnd[0]/usr/txtRSYST-MANDT").text = SAP_CLIENT
            session.findById("wnd[0]/usr/txtRSYST-BNAME").text = SAP_USER
            session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = SAP_PASS
            session.findById("wnd[0]/usr/txtRSYST-LANGU").text = SAP_LANGUAGE
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(3)
            handle_multiple_logon_popup(session)
            print("✅ Logged in successfully.")
        else:
            print("✅ Session เดิมล็อกอินอยู่แล้ว")
    except Exception as e:
        print("⚠️ ข้ามขั้นตอน Login:", str(e))
    
    # ------------------------------------------------------------
    # ENTER T-CODE SE16N
    # ------------------------------------------------------------
    print("📘 Opening SE16N...")
    session.StartTransaction("SE16N")
    time.sleep(2)

    def dump_usr_controls(session, depth=6):
        def rec(ctrl, level):
            prefix = "  " * level
            print(f"{prefix}- {ctrl.Id} | {ctrl.Type} | text='{getattr(ctrl,'Text','')}' | name='{getattr(ctrl,'Name','')}' | tooltip='{getattr(ctrl,'Tooltip','')}'")
            if level >= depth:
                return
            try:
                ch = ctrl.Children
                for i in range(ch.Count):
                    rec(ch.Item(i), level+1)
            except Exception:
                pass

        root = session.findById("wnd[0]/usr")
        print("===== DUMP wnd[0]/usr =====")
        rec(root, 0)
        print("===== END =====")

    
    # ====================== SE16/SE16N RPA: TABLE + VARIANT + DATE/TIME (FINAL) ======================
    # ====================== SE16/SE16N RPA: TABLE + VARIANT + DATE/TIME (HARDENED) ======================
    # ใช้กับ SAP GUI Scripting (win32com) — flow:
    #  1) ใส่ Table
    #  2) F6 → Get Variant → ใส่/เลือก Variant + User (F4 ได้)
    #  3) กลับหน้าหลัก → ใส่ Running Date (วันนี้ dd.MM.yyyy) & Running Time (ตอนนี้-1ชม. HH:MM:SS)
    #     * มี 3 ชั้นป้องกัน: Grid/ALV → Row-alignment → F4 popup
    #  4) Execute (F8)

    # ---------- Utils ----------
    def wait_until(session, path, timeout=6.0, step=0.1, optional=False):
        end = time.time() + timeout
        while time.time() < end:
            try:
                return session.findById(path)
            except Exception:
                time.sleep(step)
        if optional:
            return None
        raise TimeoutError(f"Timeout waiting for {path}")

    def _set_text_safe(ctrl, value):
        try:
            ctrl.Text = value
            return True
        except Exception:
            pass
        try:
            ctrl.SetFocus()
            ctrl.Text = value
            return True
        except Exception:
            return False

    def _press_ok_popup(session):
        try:
            session.findById("wnd[1]").sendVKey(0)  # Enter
            return
        except Exception:
            pass
        for btn in ("wnd[1]/tbar[0]/btn[0]", "wnd[1]/tbar[0]/btn[2]"):
            b = wait_until(session, btn, 1.0, optional=True)
            if b:
                b.press()
                return
        raise RuntimeError("ไม่พบปุ่ม OK ใน popup")

    # ---------- Hit List (F4) ----------
    def _find_value_help_window(session):
        for w in ("wnd[2]", "wnd[1]"):
            try:
                return session.findById(w)
            except Exception:
                continue
        raise RuntimeError("ไม่พบหน้าต่าง Hit List")

    def _find_alv_like_grid(win):
        std_paths = (
            "usr/cntlGRID1/shellcont/shell",
            "usr/cntlALV_CONTAINER_1/shellcont/shell",
            "usr/cntlRESULT_LIST/shellcont/shell",
            "usr/cntlCONTAINER/shellcont/shell",
            "usr/tblSAPLALDB", "usr/tblSAPLALDB*", "usr/tbl*",
        )
        for p in std_paths:
            try:
                if p.endswith("*"):
                    ch = getattr(win, "Children", None)
                    if ch:
                        for i in range(ch.Count):
                            c = ch.Item(i)
                            if c.Id.startswith(f"{win.Id}/{p[:-1]}"):
                                return c
                else:
                    return win.findById(f"{win.Id}/{p}")
            except Exception:
                pass

        # เดินทั้งต้นไม้หา Grid/Shell/TableControl ที่อ่านค่าได้
        def it(root):
            yield root
            try:
                ch = getattr(root, "Children", None)
                if ch:
                    for i in range(ch.Count):
                        yield from it(ch.Item(i))
            except Exception:
                pass
        for c in it(win):
            t = getattr(c, "Type", "")
            if not (t.endswith("GuiShell") or t.endswith("GuiGridView") or t.endswith("GuiTableControl")):
                continue
            for probe in ("RowCount", "VisibleRowCount"):
                try:
                    getattr(c, probe); return c
                except Exception:
                    pass
            for probe in [(0,0), (0,"VARIANT"), (0,"NAME")]:
                try:
                    c.GetCellValue(*probe); return c
                except Exception:
                    continue
        return None

    def _accept_value_help_without_grid(session):
        for w in ("wnd[2]", "wnd[1]"):
            try:
                win = session.findById(w)
            except Exception:
                continue
            try:
                win.sendVKey(0); return True
            except Exception:
                pass
            for btn in ("tbar[0]/btn[0]", "tbar[0]/btn[2]"):
                try:
                    win.findById(f"{w}/{btn}").press(); return True
                except Exception:
                    pass
        return False

    def _select_in_value_help(session, variant_name, user=None):
        win  = _find_value_help_window(session)
        grid = _find_alv_like_grid(win)
        if grid is None:
            if _accept_value_help_without_grid(session): return
            raise RuntimeError("ไม่พบคอนโทรลตารางใน Hit List")

        var_cols  = ["VARIANT","VARNAME","NAME","LTNAME","VARID",0,1]
        user_cols = ["USER","UNAME","AENAM",3]

        def _get(r, cols):
            for c in cols:
                try:
                    return str(grid.GetCellValue(r, c)).strip()
                except Exception:
                    continue
            return ""

        try:
            rows = grid.RowCount
        except Exception:
            rows = 2000

        target = None
        for r in range(rows):
            try:
                v = _get(r, var_cols)
            except Exception:
                break
            if not v: continue
            if v.lower() == variant_name.lower():
                u = _get(r, user_cols) if user else ""
                if (not user) or (u.lower() == user.lower()):
                    target = r; break

        if target is None:
            if _accept_value_help_without_grid(session): return
            raise RuntimeError(f"ไม่พบ Variant='{variant_name}' User='{user or '*'}' ใน Hit List")

        grid.currentCellRow = target
        grid.selectedRows   = str(target)
        try:
            grid.doubleClickCurrentCell()
        except Exception:
            try:
                win.sendVKey(0)
            except Exception:
                for btn in ("tbar[0]/btn[0]", "tbar[0]/btn[2]"):
                    try:
                        win.findById(f"{win.Id}/{btn}").press(); break
                    except Exception:
                        pass

    # ---------- ใส่ Table บนหน้าหลัก ----------
    def _set_table_name(session, table_name: str):
        candidates = [
            "wnd[0]/usr/ctxtGD-TAB",               # (จาก dump ของคุณ)
            "wnd[0]/usr/ctxtSE16N-TAB",
            "wnd[0]/usr/ctxtDATABROWSE-TABLENAME",
            "wnd[0]/usr/ctxtRSRD1-TBMA",
            "wnd[0]/usr/ctxtSE16N-TABLE",
        ]
        for p in candidates:
            fld = wait_until(session, p, 0.3, optional=True)
            if fld:
                fld.Text = table_name
                return
        raise RuntimeError("ไม่พบช่องกรอก 'Table' บนหน้าหลัก")

    # ---------- เติม Running Date/Time (ล็อก path ตาม dump) ----------
    def _open_f4_and_fill(session, edit_ctrl, value: str):
        try:
            edit_ctrl.SetFocus()
        except Exception:
            pass
        try:
            session.findById("wnd[0]").sendVKey(4)  # F4
        except Exception:
            pass

        # หา popup
        win = None
        for w in ("wnd[2]", "wnd[1]"):
            win = wait_until(session, w, 1.0, optional=True)
            if win: break
        if not win:
            return False

        # หา input ตัวแรกแล้วกรอก
        def first_input(node):
            try:
                ch = node.Children
                for i in range(ch.Count):
                    c = ch.Item(i)
                    t = getattr(c,"Type","")
                    if t.endswith("GuiCTextField") or t.endswith("GuiTextField"):
                        return c
                    sub = first_input(c)
                    if sub: return sub
            except Exception:
                return None
            return None

        inp = first_input(win)
        if not inp: return False
        if not _set_text_safe(inp, value): return False

        try:
            win.sendVKey(0)
        except Exception:
            ok = wait_until(session, f"{win.Id}/tbar[0]/btn[0]", 0.8, optional=True) or \
                wait_until(session, f"{win.Id}/tbar[0]/btn[2]", 0.8, optional=True)
            if ok: ok.press()
        return True

    def fill_running_datetime(session, minus_hours=1):
        """
        ใช้ path จาก DUMP:
        - Date  -> /usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,1]
        - Time  -> /usr/tblSAPLSE16NSELFIELDS_TC/ctxtGS_SELFIELDS-LOW[2,2]
        ถ้าตั้งค่าโดยตรงไม่ติด → เปิด F4 แล้วกรอกใน popup
        """
        now = datetime.now()
        run_date = now.strftime("%d.%m.%Y")
        run_time = (now - timedelta(hours=minus_hours)).strftime("%H:%M:%S")

        BASE = "wnd[0]/usr/tblSAPLSE16NSELFIELDS_TC"

        # --- DATE ---
        date_edit = wait_until(session, f"{BASE}/ctxtGS_SELFIELDS-LOW[2,1]", 2.0, optional=True)
        if not date_edit:
            raise RuntimeError("ไม่พบช่อง Running Date (LOW[2,1]) ในตาราง Selection Criteria")

        if not _set_text_safe(date_edit, run_date):
            if not _open_f4_and_fill(session, date_edit, run_date):
                raise RuntimeError("ใส่ค่า Running Date ไม่สำเร็จ (ทั้งตรงและผ่าน F4)")

        # --- TIME ---
        time_edit = wait_until(session, f"{BASE}/ctxtGS_SELFIELDS-LOW[2,2]", 2.0, optional=True)
        if not time_edit:
            raise RuntimeError("ไม่พบช่อง Running Time (LOW[2,2]) ในตาราง Selection Criteria")

        if not _set_text_safe(time_edit, run_time):
            if not _open_f4_and_fill(session, time_edit, run_time):
                raise RuntimeError("ใส่ค่า Running Time ไม่สำเร็จ (ทั้งตรงและผ่าน F4)")

        print(f"📝 Set Running Date={run_date} | Running Time={run_time}")

    # ---------- MAIN ----------
    def run_with_table_and_variant(session, table_name: str, variant_name: str, user: str = "*",
                                execute_immediately: bool = True, use_f4: bool = True):
        # 1) Table
        _set_table_name(session, table_name)

        # 2) Get Variant (F6)
        session.findById("wnd[0]").sendVKey(6)
        wait_until(session, "wnd[1]", 6)

        # 3) Popup Get Variant (path ตรงตามที่คุณให้)
        vf = session.findById("wnd[1]/usr/ctxtGS_SE16N_LT-NAME")               # Variant
        uf = wait_until(session, "wnd[1]/usr/txtGS_SE16N_LT-UNAME", 1.0, True) # User
        vf.Text = variant_name
        if uf: uf.Text = user or "*"

        if use_f4:
            try:
                vf.SetFocus()
                session.findById("wnd[1]").sendVKey(4)  # F4
                _select_in_value_help(session, variant_name, user if user != "*" else None)
            except RuntimeError as e:
                if "Hit List" in str(e) or "คอนโทรลตาราง" in str(e):
                    pass
                else:
                    raise

        _press_ok_popup(session)  # กลับหน้าหลัก

        # 4) ใส่วันเวลา (ล็อก path ตาม DUMP)
        fill_running_datetime(session, minus_hours=1)

        # 5) Execute
        if execute_immediately:
            try:
                session.findById("wnd[0]/tbar[1]/btn[8]").press()
            except Exception:
                session.findById("wnd[0]").sendVKey(8)

        print(f"✅ Table={table_name} | Variant={variant_name} | User={user} : Executed with Date/Time set")

# ---------------------- EXAMPLE ----------------------
    session.findById("wnd[0]").maximize()
    run_with_table_and_variant(
        session,
        table_name="ZTMCPMM0113",
        variant_name="IS_JOB",
        user="MCP_ADMIN2",
        execute_immediately=True,
        use_f4=True
    )

    # ------------------------------------------------------------
    # EXPORT TO EXCEL (แบบเดิมรันได้)
    # ------------------------------------------------------------
    print("💾 Exporting data to Excel...")
    try:
        time.sleep(4)
    
        # --- เลือกเมนู Export (เมนูหลักของ SAP) ---
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        time.sleep(2)
        session.findById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_GUI_CUL_EXPORT_AS:0512/cmbGS_EXPORT-FORMAT").setFocus
        time.sleep(2)
    
        # --- กดปุ่ม Excel Export (ถ้ามี popup) ---
        popup_export = session.findById("wnd[1]/tbar[0]/btn[20]", False)
        if popup_export:
            popup_export.press()
            print("📄 Selected Excel export option.")
            time.sleep(2)
        
    
    # --- กำหนดเฉพาะโฟลเดอร์ปลายทาง และปล่อยให้ SAP ตั้งชื่อไฟล์เอง ---
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = EXPORT_FOLDER
        # อย่าแตะ DY_FILENAME เพื่อคงชื่อไฟล์เดิมที่ SAP เสนอ
        # session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = ...  # (ไม่ต้องใช้)
    
        # session.findById("wnd[1]/usr/ctxtDY_PATH").text = os.path.dirname(EXPORT_PATH)
        # session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = os.path.basename(EXPORT_PATH)
        session.findById("wnd[1]/tbar[0]/btn[11]").press()
        time.sleep(3)
    
        # --- ปิด popup ที่อาจขึ้น ---
        for wnd_id in range(1, 8):
            try:
                wnd = session.findById(f"wnd[{wnd_id}]", False)
                if wnd:
                    common_buttons = ["btn[0]", "btn[1]", "btn[11]", "btn[20]"]
                    for btn in common_buttons:
                        try:
                            wnd.findById(f"tbar[0]/{btn}", False).press()
                            time.sleep(1)
                        except:
                            pass
            except:
                pass
    
        print(f"✅ Exported successfully to {EXPORT_FOLDER}")
        export_ok = True
        # หลัง save/export เสร็จ
        xlsx_path = wait_for_real_xlsx(EXPORT_FOLDER, timeout=60)
        # ถ้าไม่ต้องแนบไฟล์ เราใช้แค่โชว์ในอีเมลได้
    except Exception as e:
        print("⚠️ Export ไม่สำเร็จ:", str(e))
    
    # ------------------------------------------------------------
    # CLOSE SAP & EXCEL
    # ------------------------------------------------------------
    print("🧹 Closing SAP session and Excel...")
    
    try:
        if session is not None:
            try:
                session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
                session.findById("wnd[0]").sendVKey(0)
                print("✅ Closed SAP session gracefully.")
                time.sleep(3)
            except Exception as e:
                print("⚠️ ปิด SAP ผ่านคำสั่ง /nex ไม่ได้:", str(e))
        else:
            print("ℹ️ ไม่พบ session ที่เปิดอยู่")
    
        # ปิด Excel ที่เปิดโดย SAP
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name'] and proc.info['name'].lower().startswith("excel"):
                    proc.terminate()
                    print(f"🗙 ปิด Excel (PID={proc.pid}) สำเร็จ")
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
    
        # ปิด SAP GUI
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name'] and proc.info['name'].lower().startswith("saplogon"):
                    proc.terminate()
                    print(f"🗙 ปิด SAP GUI (PID={proc.pid}) สำเร็จ")
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
    
        print("🏁 All done. SAP & Excel closed successfully.")
    except Exception as e:
        print("⚠️ ปิดโปรแกรมไม่สำเร็จ:", str(e))
    pass
except Exception as ex:
    last_error = f"{ex}\n\nTraceback:\n{traceback.format_exc()}"
    export_ok = False


try:
    # ======= สรุปผล & ส่งเมล (SUCCESS / FAILURE) =======
    elapsed = (datetime.now() - start_time).total_seconds()
    if export_ok:
        html = email_success_html(elapsed, EXPORT_FOLDER)  # หรือจะใส่ xlsx_path ในข้อความก็ได้
        send_mail(f"[SUCCESS] STO Export — {datetime.now():%Y-%m-%d}", html)
    else:
        html = email_error_html(last_error or "Unknown error", elapsed)
        send_mail(f"[FAILED] STO Export — {datetime.now():%Y-%m-%d}", html)

except Exception as e:
    # กันกรณีการส่งเมลเองล้มเหลว
    print("⚠️ ส่งเมลสรุปผลไม่สำเร็จ:", e)

