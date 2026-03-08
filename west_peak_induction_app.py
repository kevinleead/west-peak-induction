import os
import json
import uuid
import base64
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

from flask import Flask, request, render_template_string, send_file, redirect, url_for
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas

app = Flask(__name__)

BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = BASE_DIR / "data.json"
PDF_DIR = BASE_DIR / "pdf"
UPLOAD_DIR = BASE_DIR / "uploads"
DOCS_DIR = BASE_DIR / "docs"
SIG_FILE = BASE_DIR / "sig_temp.png"
CARD_FILE = UPLOAD_DIR / "card_temp.png"

PDF_DIR.mkdir(exist_ok=True)
UPLOAD_DIR.mkdir(exist_ok=True)
DOCS_DIR.mkdir(exist_ok=True)

TITLE = "West Peak Construction Site Induction & SWMS Form V3"
BASE_URL = os.environ.get("BASE_URL", "http://127.0.0.1:5050")
SITE_INDUCTION_FILENAME = "SITE INDUCTION & OHS Coordination Plan.pdf"
WORKING_AT_HEIGHT_FILENAME = "SWMS - Working at Height.doc"

SITES = {
    "L1": {"code": "L1", "address": "1 Viscount Drive Doncaster"},
    "A25": {"code": "A25", "address": "25 Allanfield Cres, Wantirna South"},
}

TRADES = [
    "Demolisher",
    "Excavator",
    "Concreter",
    "Steel fabricator",
    "Plumber",
    "Electrician",
    "Carpenter",
    "Roofer",
    "Plasterer",
    "Scaffolder",
    "Cladding installer",
    "Bricklayer",
    "Air-conditioner",
    "Tiler",
    "Joiner",
    "Glazier",
    "Other Sub-contractors",
]

TRADE_SWMS_MAP = {
    "Demolisher": ["SWMS - Demolition.docx"],
    "Excavator": ["SWMS - Excavation.doc"],
    "Concreter": [WORKING_AT_HEIGHT_FILENAME],
    "Steel fabricator": [WORKING_AT_HEIGHT_FILENAME],
    "Plumber": ["SWMS - Plumbing.doc", WORKING_AT_HEIGHT_FILENAME],
    "Electrician": ["SWMS - Electrical.doc", WORKING_AT_HEIGHT_FILENAME],
    "Carpenter": ["SWMS - General Carpentry (multi storey house frame erection).doc", WORKING_AT_HEIGHT_FILENAME],
    "Roofer": ["SWMS - Roof Tiling.docx", WORKING_AT_HEIGHT_FILENAME],
    "Plasterer": [WORKING_AT_HEIGHT_FILENAME],
    "Scaffolder": [WORKING_AT_HEIGHT_FILENAME],
    "Cladding installer": [WORKING_AT_HEIGHT_FILENAME],
    "Bricklayer": ["SWMS - Bricklaying.doc", WORKING_AT_HEIGHT_FILENAME],
    "Air-conditioner": [WORKING_AT_HEIGHT_FILENAME],
    "Tiler": [WORKING_AT_HEIGHT_FILENAME],
    "Joiner": [WORKING_AT_HEIGHT_FILENAME],
    "Glazier": [WORKING_AT_HEIGHT_FILENAME],
    "Other Sub-contractors": [WORKING_AT_HEIGHT_FILENAME],
}

SITE_INDUCTION_SUMMARY = """
Site Induction / OHS Coordination Plan summary:
- Principal contractor is West Peak Constructions Pty Ltd.
- All workers must receive site safety information before starting work.
- Workers must follow site rules, PPE requirements, electrical safety, scaffold rules, housekeeping, incident reporting and working at height controls.
- Subcontractors must provide SWMS and ensure work is performed in accordance with SWMS.
""".strip()

SWMS_SUMMARIES = {
    "SWMS - Demolition.docx": "Demolition SWMS covers service isolation, exclusion zones, demolition sequencing, roof demolition controls, dust suppression and housekeeping.",
    "SWMS - Excavation.doc": "Excavation SWMS covers locating underground services, plant movement, trench collapse risk, traffic control and rescue from excavation.",
    "SWMS - Plumbing.doc": "Plumbing SWMS covers service isolation, site PPE, housekeeping, excavation awareness, equipment planning and working at heights.",
    "SWMS - Electrical.doc": "Electrical SWMS covers isolation / lockout, testing for de-energisation, safety observer, PPE, energised work controls and fault-finding precautions.",
    "SWMS - General Carpentry (multi storey house frame erection).doc": "Carpentry SWMS covers frame and truss erection, powered mobile plant, ladders, fall protection, material handling and site clean up.",
    "SWMS - Roof Tiling.docx": "Roof tiling SWMS covers roof access, ladder setup, tile elevator use, fall protection, dust control and safe roof clean-up.",
    "SWMS - Bricklaying.doc": "Bricklaying SWMS covers manual handling, mixers and saws, scaffold use, laying bricks and blocks, brick elevators and housekeeping.",
    WORKING_AT_HEIGHT_FILENAME: "Working at Height SWMS covers ladder safety, scaffolds, EWP use, roof edge protection, fall arrest, falling objects and work area access controls.",
}


def ensure_data_file() -> None:
    if not DATA_FILE.exists():
        DATA_FILE.write_text("[]", encoding="utf-8")


def load_data() -> list:
    ensure_data_file()
    return json.loads(DATA_FILE.read_text(encoding="utf-8"))


def save_data(data: list) -> None:
    DATA_FILE.write_text(json.dumps(data, indent=2), encoding="utf-8")


def doc_url(file_name: str) -> str:
    return f"{BASE_URL}/docs/{quote(file_name)}"


def swms_docs_for_trade(trade: str) -> list:
    files = TRADE_SWMS_MAP.get(trade, [WORKING_AT_HEIGHT_FILENAME])
    docs = []
    for file_name in files:
        docs.append(
            {
                "file_name": file_name,
                "display_name": file_name.replace("SWMS - ", "").replace(".docx", "").replace(".doc", ""),
                "summary": SWMS_SUMMARIES.get(file_name, "Please review the attached SWMS document."),
                "link": doc_url(file_name),
            }
        )
    return docs


def write_temp_image(base64_data: str, path: Path) -> Path:
    raw = base64.b64decode(base64_data.split(",", 1)[1])
    path.write_bytes(raw)
    return path

def upload_to_drive(file_path, file_name):
    import json
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload

    creds_dict = json.loads(os.environ["GOOGLE_CREDENTIALS"])

    credentials = service_account.Credentials.from_service_account_info(
        creds_dict,
        scopes=["https://www.googleapis.com/auth/drive"]
    )

    drive_service = build("drive", "v3", credentials=credentials)

    folder_id = "1RymqcvXVYAZc90R-k4JqQ13Lx0ZqtDhJ"

    file_metadata = {
        "name": file_name,
        "parents": [folder_id]
    }

    media = MediaFileUpload(file_path, mimetype="application/pdf")

    drive_service.files().create(
        body=file_metadata,
        media_body=media,
        fields="id"
    ).execute()


def create_pdf(record: dict, payload: dict) -> str:
    safe_name = payload["full_legal_name"].replace(" ", "_")
    pdf_name = f"{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}_{record['site_code']}_{record['trade']}_{safe_name}.pdf"
    pdf_path = PDF_DIR / pdf_name

    sig_path = write_temp_image(payload["signature"], SIG_FILE)
    card_path = write_temp_image(payload["card_photo"], CARD_FILE)

    c = canvas.Canvas(str(pdf_path), pagesize=A4)
    width, height = A4
    y = height - 30

    def line(text: str, size: int = 9, step: int = 13):
        nonlocal y
        c.setFont("Helvetica", size)
        c.drawString(30, y, text[:140])
        y -= step

    c.setFont("Helvetica-Bold", 12)
    c.drawString(30, y, TITLE)
    y -= 20

    line(f"Site Code: {record['site_code']}")
    line(f"Site Address: {record['site_address']}")
    line(f"Trade: {record['trade']}")
    line(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    y -= 4

    c.setFont("Helvetica-Bold", 10)
    c.drawString(30, y, "Worker Details")
    y -= 16
    line(f"Full Legal Name: {payload['full_legal_name']}")
    line(f"Mobile Number: {payload['mobile_number']}")

    c.setFont("Helvetica-Bold", 10)
    c.drawString(30, y, "Safety Card Photo")
    y -= 12
    c.drawImage(ImageReader(str(card_path)), 30, y - 70, width=110, height=70, preserveAspectRatio=True, mask='auto')
    y -= 82

    c.setFont("Helvetica-Bold", 10)
    c.drawString(30, y, "Acknowledgements")
    y -= 16
    line("[X] I have read and understood all Site Induction / OHS Coordination Plan requirements")
    for item in payload["swms_acks"]:
        line(f"[X] I have read and understood all requirements in SWMS: {item['name']}")
    line("[X] I agree to comply with all site safety requirements")
    line("[X] I agree to follow the applicable SWMS")
    line("[X] I confirm the information provided is correct")
    y -= 6

    c.setFont("Helvetica-Bold", 10)
    c.drawString(30, y, "Document Links")
    y -= 16
    induction_url = payload.get("induction_link", "")
    if induction_url:
        c.setFont("Helvetica", 8)
        c.drawString(30, y, "Site Induction / OHS Coordination Plan")
        c.linkURL(induction_url, (30, y - 2, 230, y + 8), relative=0)
        y -= 12
    for item in payload["swms_acks"]:
        if item.get("link"):
            c.setFont("Helvetica", 8)
            c.drawString(30, y, f"SWMS: {item['name']}")
            c.linkURL(item["link"], (30, y - 2, 200, y + 8), relative=0)
            y -= 12
    y -= 4

    c.setFont("Helvetica-Bold", 10)
    c.drawString(30, y, "Signature")
    y -= 12
    c.drawImage(ImageReader(str(sig_path)), 30, y - 60, width=170, height=60, preserveAspectRatio=True, mask='auto')

c.save()

if os.environ.get("GOOGLE_CREDENTIALS"):
    try:
        upload_to_drive(pdf_path, pdf_name)
    except Exception as e:
        print(f"Google Drive upload failed: {e}")

return pdf_name


HOME_HTML = """
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{{ title }}</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 16px; background: #f5f5f5; color: #111; }
    .wrap { max-width: 980px; margin: 0 auto; }
    .card { background: #fff; border-radius: 10px; padding: 16px; margin-bottom: 16px; box-shadow: 0 1px 3px rgba(0,0,0,.08); }
    label { display: block; font-size: 14px; margin-bottom: 6px; }
    input, select, button { width: 100%; box-sizing: border-box; padding: 10px; margin-bottom: 12px; font-size: 16px; }
    .row { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
    table { width: 100%; border-collapse: collapse; font-size: 14px; }
    th, td { text-align: left; border-bottom: 1px solid #ddd; padding: 8px; vertical-align: top; }
    .pill { display: inline-block; padding: 4px 8px; border-radius: 999px; font-size: 12px; }
    .pending { background: #fff3cd; }
    .signed { background: #d1e7dd; }
    .actions a { display: inline-block; margin-right: 8px; margin-bottom: 6px; }
    @media (max-width: 680px) { .row { grid-template-columns: 1fr; } }
  </style>
</head>
<body>
<div class="wrap">
  <div class="card">
    <h2>{{ title }}</h2>
    <p>Site Supervisor Page</p>
    <form action="/create" method="post">
      <div class="row">
        <div>
          <label>Site Code</label>
          <select id="site_code" name="site_code" onchange="syncAddress()" required>
            {% for code, site in sites.items() %}
              <option value="{{ code }}">{{ code }}</option>
            {% endfor %}
          </select>
        </div>
        <div>
          <label>Site Address</label>
          <input id="site_address" name="site_address" readonly required>
        </div>
      </div>
      <div class="row">
        <div>
          <label>Trade</label>
          <select name="trade" required>
            {% for trade in trades %}
              <option value="{{ trade }}">{{ trade }}</option>
            {% endfor %}
          </select>
        </div>
        <div>
          <label>Worker Mobile</label>
          <input name="worker_mobile" placeholder="Optional for record only">
        </div>
      </div>
      <button type="submit">Generate Link / QR</button>
    </form>
  </div>

  <div class="card">
    <h3>Records</h3>
    <table>
      <tr>
        <th>#</th>
        <th>Site</th>
        <th>Trade</th>
        <th>Worker</th>
        <th>Status</th>
        <th>Actions</th>
      </tr>
      {% for r in data %}
      <tr>
        <td>{{ loop.index }}</td>
        <td>{{ r.site_code }}<br>{{ r.site_address }}</td>
        <td>{{ r.trade }}</td>
        <td>{{ r.full_legal_name or '-' }}</td>
        <td>{% if r.status == 'signed' %}<span class="pill signed">Signed</span>{% else %}<span class="pill pending">Pending</span>{% endif %}</td>
        <td class="actions">
          <a href="/sign/{{ r.token }}" target="_blank">Open link</a>
          <a href="/qr/{{ r.token }}" target="_blank">QR code</a>
          {% if r.pdf %}<a href="/pdf/{{ r.pdf }}">Download PDF</a>{% endif %}
        </td>
      </tr>
      {% endfor %}
    </table>
  </div>
</div>
<script>
const sites = {{ sites_json | safe }};
function syncAddress(){
  const code = document.getElementById('site_code').value;
  document.getElementById('site_address').value = sites[code].address;
}
syncAddress();
</script>
</body>
</html>
"""

SIGN_HTML = """
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{{ title }}</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 14px; line-height: 1.45; color: #111; background: #fafafa; }
    .wrap { max-width: 760px; margin: 0 auto; }
    .section { border: 1px solid #ddd; border-radius: 10px; padding: 14px; margin-bottom: 14px; background: #fff; }
    input[type=text], input[type=tel], input[type=file], button { width: 100%; box-sizing: border-box; padding: 10px; margin-top: 6px; margin-bottom: 10px; font-size: 16px; }
    .doc-link { display: inline-block; padding: 8px 12px; background: #f2f2f2; border-radius: 8px; text-decoration: none; color: #111; margin-top: 8px; }
    .checkbox { margin-bottom: 10px; }
    .checkbox label { display: flex; gap: 8px; align-items: flex-start; }
    canvas { border: 1px solid #000; background: #fff; width: 100%; max-width: 360px; height: 180px; touch-action: none; display: block; }
    .muted { color: #555; font-size: 14px; white-space: pre-wrap; }
    .swms-item { border: 1px dashed #ccc; border-radius: 8px; padding: 10px; margin-top: 10px; }
  </style>
</head>
<body>
<div class="wrap">
  <div class="section">
    <h2>{{ title }}</h2>
    <p><b>Site Code:</b> {{ record.site_code }}</p>
    <p><b>Site Address:</b> {{ record.site_address }}</p>
    <p><b>Trade:</b> {{ record.trade }}</p>
    <p><b>Worker Mobile:</b> {{ record.worker_mobile or '-' }}</p>
  </div>

  <div class="section">
    <h3>Worker Details</h3>
    <label>Full Legal Name</label>
    <input type="text" id="full_legal_name" required>

    <label>Mobile Number</label>
    <input type="tel" id="mobile_number" required>

    <label>Safety Card Photo (White Card / Red Card)</label>
    <input type="file" id="card_photo" accept="image/*" capture="environment" required>
    <div class="muted">Tap here to take a photo and upload it directly. On some phones the browser may still show camera or photo library choices.</div>
  </div>

  <div class="section">
    <h3>Site Induction / OHS Coordination Plan</h3>
    <div class="muted">{{ site_induction_summary }}</div>
    <a class="doc-link" href="{{ induction_link }}" target="_blank">View Site Induction / OHS Coordination Plan</a>
    <div class="checkbox">
      <label><input type="checkbox" id="induction_ack"> <span>I have read and understood all Site Induction / OHS Coordination Plan requirements.</span></label>
    </div>
  </div>

  <div class="section">
    <h3>Applicable SWMS</h3>
    {% for swms in swms_docs %}
      <div class="swms-item">
        <b>{{ swms.display_name }}</b>
        <div class="muted">{{ swms.summary }}</div>
        <a class="doc-link" href="{{ swms.link }}" target="_blank">View SWMS</a>
        <div class="checkbox">
          <label><input type="checkbox" class="swms_ack" data-swms="{{ swms.display_name }}" data-link="{{ swms.link }}"> <span>I have read and understood all requirements in this SWMS.</span></label>
        </div>
      </div>
    {% endfor %}
  </div>

  <div class="section">
    <h3>Declaration</h3>
    <div class="checkbox">
      <label><input type="checkbox" id="decl_1"> <span>I agree to comply with all site safety requirements.</span></label>
    </div>
    <div class="checkbox">
      <label><input type="checkbox" id="decl_2"> <span>I agree to follow the applicable SWMS.</span></label>
    </div>
    <div class="checkbox">
      <label><input type="checkbox" id="decl_3"> <span>I confirm the information provided is correct.</span></label>
    </div>
  </div>

  <div class="section">
    <h3>Signature</h3>
    <canvas id="sig" width="360" height="180"></canvas>
    <button type="button" onclick="clearSig()">Clear</button>
    <button type="button" onclick="submitForm()">Submit</button>
  </div>
</div>

<script>
const canvas = document.getElementById('sig');
const ctx = canvas.getContext('2d');
ctx.lineWidth = 2;
ctx.lineCap = 'round';
let drawing = false;

function getPos(e) {
  const rect = canvas.getBoundingClientRect();
  if (e.touches && e.touches.length) {
    return { x: e.touches[0].clientX - rect.left, y: e.touches[0].clientY - rect.top };
  }
  return { x: e.clientX - rect.left, y: e.clientY - rect.top };
}

function startDraw(e) {
  drawing = true;
  const p = getPos(e);
  ctx.beginPath();
  ctx.moveTo(p.x, p.y);
}

function moveDraw(e) {
  if (!drawing) return;
  e.preventDefault();
  const p = getPos(e);
  ctx.lineTo(p.x, p.y);
  ctx.stroke();
}

function stopDraw() {
  drawing = false;
}

canvas.addEventListener('mousedown', startDraw);
canvas.addEventListener('mousemove', moveDraw);
canvas.addEventListener('mouseup', stopDraw);
canvas.addEventListener('mouseleave', stopDraw);
canvas.addEventListener('touchstart', startDraw, {passive:false});
canvas.addEventListener('touchmove', moveDraw, {passive:false});
canvas.addEventListener('touchend', stopDraw, {passive:false});

function clearSig() {
  ctx.clearRect(0, 0, canvas.width, canvas.height);
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

async function submitForm() {
  const fullLegalName = document.getElementById('full_legal_name').value.trim();
  const mobileNumber = document.getElementById('mobile_number').value.trim();
  const cardFile = document.getElementById('card_photo').files[0];
  const inductionAck = document.getElementById('induction_ack').checked;
  const decl1 = document.getElementById('decl_1').checked;
  const decl2 = document.getElementById('decl_2').checked;
  const decl3 = document.getElementById('decl_3').checked;
  const swmsAcks = Array.from(document.querySelectorAll('.swms_ack')).map(cb => ({ name: cb.dataset.swms, checked: cb.checked, link: cb.dataset.link }));

  if (!fullLegalName || !mobileNumber || !cardFile) {
    alert('Please complete Full Legal Name, Mobile Number and Safety Card Photo.');
    return;
  }
  if (!inductionAck || !decl1 || !decl2 || !decl3 || swmsAcks.some(x => !x.checked)) {
    alert('Please tick all required acknowledgements before submitting.');
    return;
  }

  const payload = {
    full_legal_name: fullLegalName,
    mobile_number: mobileNumber,
    induction_ack: inductionAck,
    declarations: [decl1, decl2, decl3],
    swms_acks: swmsAcks,
    signature: canvas.toDataURL('image/png'),
    card_photo: await fileToBase64(cardFile),
    induction_link: '{{ induction_link }}'
  };

  const res = await fetch('/submit/{{ record.token }}', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(payload),
  });

  if (res.ok) {
    alert('PDF saved');
    window.location.href = '/';
  } else {
    const txt = await res.text();
    alert('Submit failed: ' + txt);
  }
}
</script>
</body>
</html>
"""

QR_HTML = """
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>QR Code</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 20px; text-align: center; }
    .box { max-width: 420px; margin: 0 auto; border: 1px solid #ddd; padding: 16px; border-radius: 10px; }
    img { max-width: 100%; }
    .link { margin-top: 12px; word-break: break-all; font-size: 14px; }
  </style>
</head>
<body>
  <div class="box">
    <h3>Worker QR Code</h3>
    <img src="{{ qr_src }}" alt="QR">
    <div class="link">{{ full_link }}</div>
  </div>
</body>
</html>
"""


@app.route("/")
def home():
    data = load_data()
    return render_template_string(HOME_HTML, title=TITLE, data=data, sites=SITES, trades=TRADES, sites_json=json.dumps(SITES))


@app.route("/create", methods=["POST"])
def create():
    data = load_data()
    token = str(uuid.uuid4())
    site_code = request.form["site_code"]
    record = {
        "token": token,
        "site_code": site_code,
        "site_address": request.form["site_address"],
        "trade": request.form["trade"],
        "worker_mobile": request.form.get("worker_mobile", ""),
        "status": "pending",
        "full_legal_name": "",
        "pdf": "",
    }
    data.append(record)
    save_data(data)
    return redirect(url_for("home"))


@app.route("/sign/<token>")
def sign(token):
    data = load_data()
    record = next((r for r in data if r.get("token") == token), None)
    if not record:
        return "Record not found", 404
    return render_template_string(
        SIGN_HTML,
        title=TITLE,
        record=record,
        site_induction_summary=SITE_INDUCTION_SUMMARY,
        induction_link=doc_url(SITE_INDUCTION_FILENAME),
        swms_docs=swms_docs_for_trade(record["trade"]),
    )


@app.route("/submit/<token>", methods=["POST"])
def submit(token):
    data = load_data()
    record = next((r for r in data if r.get("token") == token), None)
    if not record:
        return "Record not found", 404

    payload = request.get_json()
    record["full_legal_name"] = payload["full_legal_name"]
    record["status"] = "signed"
    record["pdf"] = create_pdf(record, payload)
    save_data(data)
    return "OK"


@app.route("/pdf/<name>")
def pdf(name):
    return send_file(PDF_DIR / name, as_attachment=True)


@app.route("/docs/<path:name>")
def docs(name):
    path = DOCS_DIR / name
    if not path.exists():
        return f"Document not found: {name}", 404
    return send_file(path, as_attachment=False)


@app.route("/qr/<token>")
def qr(token):
    data = load_data()
    record = next((r for r in data if r.get("token") == token), None)
    if not record:
        return "Record not found", 404
    full_link = request.url_root.rstrip("/") + url_for("sign", token=token)
    qr_src = f"https://api.qrserver.com/v1/create-qr-code/?size=260x260&data={quote(full_link)}"
    return render_template_string(QR_HTML, qr_src=qr_src, full_link=full_link)


if __name__ == "__main__":
    ensure_data_file()
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5050)), debug=True)
