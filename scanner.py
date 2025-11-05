# scanner.py
# Minimal deploy, polished UI, DevBraze branding, rear-camera preference, manual switch,
# flash animations + optional sound feedback (assets/*), and no gate field.

import os
from datetime import datetime
from zoneinfo import ZoneInfo

from flask import Flask, request, jsonify, render_template_string, send_from_directory
from openpyxl import load_workbook
from filelock import FileLock

# ========= CONFIG =========
APP_DIR     = os.path.dirname(__file__)
ASSETS_DIR  = os.path.join(APP_DIR, "assets")
EXCEL_PATH  = os.path.join(ASSETS_DIR, "Teams.xlsx")  # must exist; sheet "Teams"
SHEET_NAME  = "Teams"
TIMEZONE    = "Asia/Kolkata"
REQUIRED_COLS = ["Token", "Entry Confirmed", "Check-in Time", "Check-in Gate"]  # Gate kept for compatibility only.
# =========================

app = Flask(__name__)

SCANNER_HTML = """
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>DevBraze — Code Crusade Scanner</title>
<style>
  :root { --brand:#0F766E; --accent:#0EA5E9; --bg:#f6f7f9; --card:#ffffff; --text:#111827; --muted:#6b7280; }
  * { box-sizing:border-box; }
  body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif; margin:0; background:var(--bg); color:var(--text); }
  header {
    background: linear-gradient(90deg, var(--brand) 0%, var(--accent) 100%);
    color:white; padding:14px 20px; font-size:18px; font-weight:700;
    display:flex; align-items:center; justify-content:space-between; border-bottom:3px solid #14b8a6;
  }
  header .title { display:flex; gap:10px; align-items:center; }
  header img { height:36px; border-radius:6px; background:rgba(255,255,255,0.1); padding:4px; }
  main { max-width:1000px; margin:18px auto; padding:0 12px; }
  .card { background:var(--card); padding:16px; border-radius:16px; box-shadow:0 12px 30px rgba(0,0,0,.08); }
  .row { display:flex; gap:18px; flex-wrap:wrap; align-items:flex-start; }
  #reader { width: 400px; max-width:100%; border:3px solid var(--accent); border-radius:12px; padding:6px; background:#f0fdfa; box-shadow: 0 0 20px rgba(14,165,233,.18); }
  .status { margin-top:12px; padding:12px; border-radius:10px; font-weight:600; }
  .ok { background:#ecfdf5; color:#065f46; border:1px solid #a7f3d0; }
  .warn { background:#fef3c7; color:#92400e; border:1px solid #fde68a; }
  .err { background:#fee2e2; color:#991b1b; border:1px solid #fecaca; }
  label { display:block; font-size:14px; color:#374151; margin-top:8px; }
  input, button, select { padding:10px 12px; border-radius:10px; border:1px solid #e5e7eb; font-size:14px; }
  button { background:var(--accent); color:white; border:0; cursor:pointer; font-weight:700; }
  button:hover { filter:brightness(.95); }
  .grid { display:grid; grid-template-columns: 1fr; gap:10px; }
  .toolbar { display:flex; gap:8px; align-items:center; flex-wrap:wrap; margin-bottom:10px; }
  .log { font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; background:#111827; color:#e5e7eb; padding:10px; border-radius:10px; height:160px; overflow:auto; font-size:12px; }
  #stats { color:var(--muted); font-size:14px; text-align:right; margin-top:6px; }

  /* Flash feedback */
  .flash-success { animation: flashGreen 280ms ease; }
  .flash-warn    { animation: flashYellow 280ms ease; }
  .flash-error   { animation: flashRed 280ms ease; }
  @keyframes flashGreen { 0%{box-shadow:0 0 24px 8px #22c55e;} 100%{box-shadow:none;} }
  @keyframes flashYellow{ 0%{box-shadow:0 0 24px 8px #facc15;} 100%{box-shadow:none;} }
  @keyframes flashRed   { 0%{box-shadow:0 0 24px 8px #ef4444;} 100%{box-shadow:none;} }

  /* Loading overlay */
  #loading {
    position:fixed; inset:0; display:flex; align-items:center; justify-content:center;
    background:rgba(0,0,0,.55); color:white; z-index:999; backdrop-filter: blur(2px);
    font-weight:700; letter-spacing:.3px;
  }
  .pill { background:rgba(255,255,255,.12); padding:10px 14px; border-radius:999px; border:1px solid rgba(255,255,255,.2); }

  @media (max-width: 680px) {
    .row { flex-direction:column; }
    #reader { width:100%!important; }
    .card { padding:12px; }
    header { padding:12px; font-size:16px; }
  }
</style>
</head>
<body>
<div id="loading"><div class="pill">🎥 Initializing camera…</div></div>

<header>
  <div class="title">
    <span>Code Crusade Hackathon — Check-In</span>
  </div>
  <img src="/assets/devbraze_logo.png" alt="DevBraze"/>
</header>

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

        <div class="grid" style="margin-top:12px;">
          <label>Manual token (fallback)</label>
          <input id="manual" placeholder="Paste/enter token here"/>
          <button id="btnManual">Check-in Manually</button>
        </div>
      </div>

      <div style="flex:1; min-width:280px;">
        <div id="status" class="status warn">Ready. Scan a QR or enter a token.</div>
        <div id="stats"></div>
        <div class="log" id="log"></div>
        <div style="margin-top:8px; display:flex; gap:8px; flex-wrap:wrap;">
          <button id="clearLog" type="button" style="background:#374151;">Clear Log</button>
        </div>
      </div>
    </div>
  </div>
</main>

<!-- Optional sounds (will be attempted if present in /assets) -->
<audio id="sound-ok"     src="/assets/success.mp3" preload="auto"></audio>
<audio id="sound-repeat" src="/assets/repeat.mp3"  preload="auto"></audio>
<audio id="sound-error"  src="/assets/error.mp3"   preload="auto"></audio>

<script src="https://unpkg.com/html5-qrcode"></script>
<script>
const statusBox   = document.getElementById('status');
const logBox      = document.getElementById('log');
const statsBox    = document.getElementById('stats');
const manualEl    = document.getElementById('manual');
const btnManual   = document.getElementById('btnManual');
const cameraSel   = document.getElementById('cameraSelect');
const btnSwitch   = document.getElementById('btnSwitch');
const clearLogBtn = document.getElementById('clearLog');
const loading     = document.getElementById('loading');

const sndOk       = document.getElementById('sound-ok');
const sndRepeat   = document.getElementById('sound-repeat');
const sndError    = document.getElementById('sound-error');

let inFlight = false;
let reader   = null;
let cameras  = [];
let currentId = null;

let countChecked = 0; // lightweight local counter (optional)
let totalTeams   = null; // if you want to feed this from server later

function safePlay(audioEl){
  try { audioEl && audioEl.play().catch(()=>{}); } catch(e) {}
}

function flash(kind){
  const cls = kind === 'ok' ? 'flash-success' : kind === 'warn' ? 'flash-warn' : 'flash-error';
  document.body.classList.add(cls);
  setTimeout(()=>document.body.classList.remove(cls), 320);
}

function setStatus(kind, msg){
  statusBox.className = 'status ' + kind;
  statusBox.textContent = msg;
  logBox.textContent = `[${new Date().toLocaleTimeString()}] ${msg}\\n` + logBox.textContent;
}

function updateStats(){
  if (totalTeams != null) {
    statsBox.textContent = `✅ ${countChecked} / ${totalTeams} Teams Checked In`;
  } else {
    statsBox.textContent = `✅ ${countChecked} checked-in (session)`;
  }
}

async function checkin(token){
  if (inFlight) return;
  inFlight = true;
  try{
    const r = await fetch('/api/checkin', {
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body: JSON.stringify({ token }) // no gate
    });
    let text = await r.text();
    let j = {};
    try { j = JSON.parse(text); } catch {}

    if (!r.ok) {
      const msg = j.message || text || 'Server error';
      setStatus('err', `❌ ${msg.trim()}`);
      flash('err'); safePlay(sndError);
    } else if (j.status === 'ok') {
      setStatus('ok', `✅ Checked in at ${j.time}`);
      flash('ok'); safePlay(sndOk);
      countChecked++; updateStats();
    } else if (j.status === 'repeat') {
      setStatus('warn', `⚠️ Already checked in at ${j.time}`);
      flash('warn'); safePlay(sndRepeat);
    } else {
      setStatus('err', `❌ ${j.message || 'Invalid response'}`);
      flash('err'); safePlay(sndError);
    }
  } catch(e){
    setStatus('err', 'Network or server error.');
    flash('err'); safePlay(sndError);
  } finally {
    inFlight = false;
  }
}

btnManual.onclick = () => {
  const t = manualEl.value.trim();
  if (!t) return setStatus('warn','Enter a token first.');
  checkin(t);
};
clearLogBtn.onclick = () => { logBox.textContent = ''; };

function onScanSuccess(decodedText){
  const token = decodedText.trim();
  if (!token) return;
  checkin(token);
}
function onScanFailure(err){ /* ignore continuous decode errors */ }

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
      { fps: 10, qrbox: 280 },
      onScanSuccess,
      onScanFailure
    );
    setStatus('ok', 'Camera running');
  } catch(e){
    setStatus('err', 'Failed to start camera. Try another option.');
  } finally {
    loading.style.display = 'none';
  }
}

function pickRearCamera(list){
  const pref = ['back','rear','environment','trás','arrière','后置','背面'];
  const lc = txt => (txt || '').toLowerCase();
  for (const w of pref){
    const hit = list.find(c => lc(c.label).includes(w));
    if (hit) return hit.id;
  }
  if (list.length >= 2) return list[list.length - 1].id; // often rear on phones
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
  loading.style.display = 'flex';
  await startReader(chosen);
};

(async function init(){
  // Try to nudge permission so labels become visible on some browsers
  try{
    const tmp = new Html5Qrcode("reader");
    let tmpCams = await Html5Qrcode.getCameras();
    let tmpId = tmpCams[0]?.id;
    if (tmpId) { try { await tmp.start(tmpId, {fps:1}, ()=>{}, ()=>{}); } catch(e) {} }
    try { await tmp.stop(); } catch(e) {}
    try { await tmp.clear(); } catch(e) {}
  } catch(e) {}

  try{
    cameras = await Html5Qrcode.getCameras();
    if (!cameras || !cameras.length){
      setStatus('err','No camera available. Use manual token entry.');
      loading.style.display = 'none';
      return;
    }
    const preferred = pickRearCamera(cameras);
    populateCameraDropdown(cameras, preferred);
    await startReader(preferred);
  } catch(e){
    setStatus('err','Camera enumeration failed. Use manual token entry.');
    loading.style.display = 'none';
  }
  updateStats();
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

# Serve files from ./assets (logo + optional sounds + Teams.xlsx if you want to download)
@app.get("/assets/<path:filename>")
def serve_assets(filename):
    return send_from_directory(ASSETS_DIR, filename)

@app.post("/api/checkin")
def api_checkin():
    try:
        data = request.get_json(force=True, silent=True) or {}
        token = (data.get("token") or "").strip()
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

                # find row by token
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
