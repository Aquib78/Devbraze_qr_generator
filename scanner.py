# scanner.py
# Minimal deploy, prefers back camera, and lets user switch cameras.

import os
from datetime import datetime
from zoneinfo import ZoneInfo

from flask import Flask, request, jsonify, render_template_string
from openpyxl import load_workbook
from filelock import FileLock

# ========= CONFIG =========
APP_DIR     = os.path.dirname(__file__)
EXCEL_PATH  = os.path.join(APP_DIR, "assets", "Teams.xlsx")  # must exist; "Teams" sheet
SHEET_NAME  = "Teams"
TIMEZONE    = "Asia/Kolkata"
REQUIRED_COLS = ["Token", "Entry Confirmed", "Check-in Time", "Check-in Gate"]
# =========================

app = Flask(__name__)

SCANNER_HTML = """
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Ticket Scanner</title>
<style>
  :root { --brand:#0F766E; --accent:#0EA5E9; }
  body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; margin:0; background:#f6f7f9; }
  header { background:var(--brand); color:#fff; padding:12px 16px; font-weight:700; }
  main { max-width:980px; margin:18px auto; padding:0 12px; }
  .card { background:#fff; padding:16px; border-radius:12px; box-shadow:0 10px 30px rgba(0,0,0,.08); }
  .row { display:flex; gap:16px; flex-wrap:wrap; align-items:flex-start; }
  #reader { width: 380px; max-width:100%; }
  .status { margin-top:12px; padding:12px; border-radius:10px; font-weight:600; }
  .ok { background:#ecfdf5; color:#065f46; border:1px solid #a7f3d0; }
  .warn { background:#fef3c7; color:#92400e; border:1px solid #fde68a; }
  .err { background:#fee2e2; color:#991b1b; border:1px solid #fecaca; }
  label { display:block; font-size:14px; color:#374151; margin-top:8px; }
  input, button, select { padding:10px 12px; border-radius:10px; border:1px solid #e5e7eb; font-size:14px; }
  button { background:var(--accent); color:white; border:0; cursor:pointer; font-weight:700; }
  button:hover { filter:brightness(.95); }
  .grid { display:grid; grid-template-columns: 1fr; gap:10px; }
  .log { font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; background:#111827; color:#e5e7eb; padding:10px; border-radius:10px; height:160px; overflow:auto; font-size:12px; }
  .toolbar { display:flex; gap:8px; align-items:center; flex-wrap:wrap; margin-bottom:10px; }
</style>
</head>
<body>
<header>Code Crusade Hackathon — Scanner</header>
<main>
  <div class="card">
    <div class="row">
      <div>
        <div class="toolbar">
          <label for="cameraSelect" style="margin:0;">Camera</label>
          <select id="cameraSelect"></select>
          <button id="btnSwitch">Switch</button>
        </div>
        <div id="reader"></div>
        <div class="grid">
          <label>Gate/Desk (optional)</label>
          <input id="gate" placeholder="e.g., Main Gate A"/>
          <label>Manual token (fallback)</label>
          <input id="manual" placeholder="Paste/enter token here"/>
          <button id="btnManual">Check-in Manually</button>
        </div>
      </div>
      <div style="flex:1; min-width:280px;">
        <div id="status" class="status warn">Ready. Scan a QR or enter a token.</div>
        <div class="log" id="log"></div>
      </div>
    </div>
  </div>
</main>

<script src="https://unpkg.com/html5-qrcode"></script>
<script>
const statusBox   = document.getElementById('status');
const logBox      = document.getElementById('log');
const gateEl      = document.getElementById('gate');
const manualEl    = document.getElementById('manual');
const btnManual   = document.getElementById('btnManual');
const cameraSel   = document.getElementById('cameraSelect');
const btnSwitch   = document.getElementById('btnSwitch');

let inFlight = false;
let reader   = null;
let cameras  = [];
let currentId = null;

function setStatus(kind, msg){
  statusBox.className = 'status ' + kind;
  statusBox.textContent = msg;
  logBox.textContent = `[${new Date().toLocaleTimeString()}] ${msg}\\n` + logBox.textContent;
}

async function checkin(token){
  if (inFlight) return;
  inFlight = true;
  const gate = gateEl.value || '';
  try{
    const r = await fetch('/api/checkin', {
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body: JSON.stringify({ token, gate })
    });
    let text = await r.text();
    let j = {};
    try { j = JSON.parse(text); } catch {}
    if (!r.ok) {
      const msg = j.message || text || 'Server error';
      setStatus('err', `❌ ${msg.trim()}`);
    } else if (j.status === 'ok') {
      if (j.note) setStatus('ok', `✅ Checked in: Team ${j.team_id} at ${j.time} — ${j.note}`);
      else setStatus('ok', `✅ Checked in: Team ${j.team_id} at ${j.time}`);
    } else if (j.status === 'repeat') {
      setStatus('warn', `⚠️ Already checked in: Team ${j.team_id} at ${j.time}`);
    } else {
      setStatus('err', `❌ ${j.message || 'Invalid response'}`);
    }
  } catch(e){
    setStatus('err', 'Network or server error.');
  } finally {
    inFlight = false;
  }
}

btnManual.onclick = () => {
  const t = manualEl.value.trim();
  if (!t) return setStatus('warn','Enter a token first.');
  checkin(t);
};

function onScanSuccess(decodedText){
  const token = decodedText.trim();
  if (!token) return;
  checkin(token);
}
function onScanFailure(err){ /* ignore noisy decode errors */ }

async function stopReader(){
  if (reader) {
    try { await reader.stop(); } catch(e) {}
    try { await reader.clear(); } catch(e) {}
  }
}

async function startReader(deviceId){
  await stopReader();
  reader = new Html5Qrcode("reader");
  currentId = deviceId;
  try{
    await reader.start(
      deviceId,
      { fps: 10, qrbox: 260 },
      onScanSuccess,
      onScanFailure
    );
    setStatus('ok', 'Camera running' + (deviceId ? ` (${deviceId.slice(0,8)}…)` : ''));
  } catch(e){
    setStatus('err', 'Failed to start camera. Try another option.');
  }
}

function pickRearCamera(list){
  // Prefer labels with back/rear/environment
  const pref = ['back','rear','environment','trás','arrière','后置','背面']; // include a few intl hints
  const lc = txt => (txt || '').toLowerCase();
  // Choose by label match
  for (const w of pref){
    const hit = list.find(c => lc(c.label).includes(w));
    if (hit) return hit.id;
  }
  // Otherwise prefer the last camera (often rear on phones)
  if (list.length >= 2) return list[list.length - 1].id;
  // Fallback to first
  return list[0]?.id || null;
}

function populateCameraDropdown(list, preferredId){
  cameraSel.innerHTML = '';
  list.forEach((c, i) => {
    const opt = document.createElement('option');
    opt.value = c.id;
    opt.textContent = c.label || `Camera ${i+1}`;
    if (c.id === preferredId) opt.selected = true;
    cameraSel.appendChild(opt);
  });
}

btnSwitch.onclick = async () => {
  const chosen = cameraSel.value;
  if (!chosen) return;
  await startReader(chosen);
};

(async function init(){
  try{
    // On first call, some browsers hide labels until permission is granted.
    // Request a quick temp permission by starting default camera and stopping immediately.
    const tmp = new Html5Qrcode("reader");
    let tmpId = (await Html5Qrcode.getCameras())[0]?.id;
    if (tmpId) { try { await tmp.start(tmpId, {fps:1}, ()=>{}, ()=>{}); } catch(e) {} }
    try { await tmp.stop(); } catch(e) {}
    try { await tmp.clear(); } catch(e) {}
  } catch(e) { /* ignore */ }

  try{
    cameras = await Html5Qrcode.getCameras();
    if (!cameras || !cameras.length){
      setStatus('err','No camera available. Use manual token entry.');
      return;
    }
    const preferred = pickRearCamera(cameras);
    populateCameraDropdown(cameras, preferred);
    await startReader(preferred);
  } catch(e){
    setStatus('err','Camera enumeration failed. Use manual token entry.');
  }
})();
</script>
</body>
</html>
"""

# ---------- helpers ----------
def ensure_excel_columns(ws):
    header = {(ws.cell(row=1, column=col).value or "").strip(): col
              for col in range(1, ws.max_column + 1)}
    for name in REQUIRED_COLS:
        if name not in header:
            ws.cell(row=1, column=ws.max_column + 1, value=name)
    header = {(ws.cell(row=1, column=col).value or "").strip(): col
              for col in range(1, ws.max_column + 1)}
    return header

def now_local_string():
    try:
            tz = ZoneInfo(TIMEZONE)
            return datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S %Z")
    except Exception:
            return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def safe_save_workbook(wb, target_path: str):
    try:
        wb.save(target_path)
        return {"saved_to": target_path, "autosave": False, "error": None}
    except PermissionError as e:
        base, ext = os.path.splitext(target_path)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        alt = f"{base}.autosave_{stamp}{ext}"
        wb.save(alt)
        return {"saved_to": alt, "autosave": True, "error": str(e)}

# ---------- routes ----------
@app.get("/scanner")
def scanner():
    return render_template_string(SCANNER_HTML)

@app.get("/")
def root():
    return "<meta http-equiv='refresh' content='0; url=/scanner' />"

@app.post("/api/checkin")
def api_checkin():
    try:
        data = request.get_json(force=True, silent=True) or {}
        token = (data.get("token") or "").strip()
        gate  = (data.get("gate")  or "").strip()
        if not token:
            return jsonify(status="error", message="No token provided"), 400

        if not os.path.exists(EXCEL_PATH):
            return jsonify(status="error", message=f"Excel not found: {EXCEL_PATH}"), 500

        lock = FileLock(EXCEL_PATH + ".lock")
        with lock:
            wb = load_workbook(EXCEL_PATH)
            try:
                if SHEET_NAME not in wb.sheetnames:
                    return jsonify(status="error", message=f"Sheet '{SHEET_NAME}' not found"), 500
                ws = wb[SHEET_NAME]

                header = ensure_excel_columns(ws)
                col_token   = header["Token"]
                col_status  = header["Entry Confirmed"]
                col_time    = header["Check-in Time"]
                col_gate    = header["Check-in Gate"]

                hit_row = None
                for r in range(2, ws.max_row + 1):
                    cell_val = ws.cell(row=r, column=col_token).value
                    if (cell_val or "").strip() == token:
                        hit_row = r
                        break
                if not hit_row:
                    return jsonify(status="error", message="Invalid token"), 404

                team_id = (ws.cell(row=hit_row, column=1).value or "").strip()
                current = (ws.cell(row=hit_row, column=col_status).value or "").strip()
                time_val = ws.cell(row=hit_row, column=col_time).value

                if current and current.lower() in ("yes", "confirmed", "present"):
                    return jsonify(status="repeat", team_id=team_id, time=str(time_val) if time_val else ""), 200

                ws.cell(row=hit_row, column=col_status, value="Yes")
                ts = now_local_string()
                ws.cell(row=hit_row, column=col_time, value=ts)
                if gate:
                    ws.cell(row=hit_row, column=col_gate, value=gate)

                result = safe_save_workbook(wb, EXCEL_PATH)
                payload = {"status": "ok", "team_id": team_id, "time": ts}
                if result["autosave"]:
                    payload["note"] = f"Original file locked. Wrote to '{os.path.basename(result['saved_to'])}'. Merge later."
                return jsonify(**payload), 200
            finally:
                wb.close()

    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify(status="error", message=f"Server exception: {type(e).__name__}: {e}"), 500

# ---------- main ----------
if __name__ == "__main__":
    port = int(os.getenv("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)
