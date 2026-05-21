#!/usr/bin/env python3
"""Highstreet Leasing Bot - Pure HTTP, no library dependencies"""
import os, json, time, logging, openpyxl
import urllib.request, urllib.parse

BOT_TOKEN = "8818096411:AAGF1Q4vJySHKtLNoHDg_5E4cktG8M-gD9A"
ADMIN_ID  = 6944516887
EXCEL_PATH = "prices.xlsx"
API = f"https://api.telegram.org/bot{BOT_TOKEN}"

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(message)s')
log = logging.getLogger(__name__)

AGENTS = {
    "tamer":   {"name": "Tamer Shalaby",  "phone": "+201005394515",  "email": "tamer.origin@gmail.com"},
    "ahmed":   {"name": "Ahmed Elhakim",  "phone": "+201221414145",  "email": "aelhakim@virgo-eg.com"},
    "azmy":    {"name": "Azmy Mohamed",   "phone": "01050405780",    "email": "azmy932@gmail.com"},
    "ehab":    {"name": "Ehab Mustafa",   "phone": "01116064830",    "email": "Elbakhshwangyehab@gmail.com"},
    "mahmoud": {"name": "Mahmoud Shaker", "phone": "+201007000933",  "email": "mshaker@virgo-eg.com"},
    "noor":    {"name": "Noor Ilani",     "phone": "+32489647992",   "email": "Noor.ilani@outlook.com"},
}

# In-memory session store
sessions = {}
pending  = {}

def api(method, data=None, files=None):
    url = f"{API}/{method}"
    if files:
        import multipart_encoder
    data = data or {}
    body = json.dumps(data).encode()
    req = urllib.request.Request(url, body, {'Content-Type': 'application/json'})
    try:
        with urllib.request.urlopen(req, timeout=30) as r:
            return json.loads(r.read())
    except Exception as e:
        log.error(f"API error {method}: {e}")
        return {}

def send(chat_id, text, reply_markup=None, parse_mode="Markdown"):
    data = {"chat_id": chat_id, "text": text, "parse_mode": parse_mode}
    if reply_markup:
        data["reply_markup"] = json.dumps(reply_markup)
    return api("sendMessage", data)

def edit(chat_id, msg_id, text):
    api("editMessageText", {"chat_id": chat_id, "message_id": msg_id, "text": text})

def answer_cb(cb_id):
    api("answerCallbackQuery", {"callback_query_id": cb_id})

def send_doc(chat_id, path, filename, caption):
    url = f"{API}/sendDocument"
    with open(path, 'rb') as f:
        file_data = f.read()
    import http.client, mimetypes
    boundary = "----FormBoundary7MA4YWxkTrZu0gW"
    body = (
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="chat_id"\r\n\r\n{chat_id}\r\n'
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="caption"\r\n\r\n{caption}\r\n'
        f"--{boundary}\r\n"
        f'Content-Disposition: form-data; name="document"; filename="{filename}"\r\n'
        f"Content-Type: application/pdf\r\n\r\n"
    ).encode() + file_data + f"\r\n--{boundary}--\r\n".encode()
    req = urllib.request.Request(url, body,
        {'Content-Type': f'multipart/form-data; boundary={boundary}'})
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            return json.loads(r.read())
    except Exception as e:
        log.error(f"Send doc error: {e}")
        return {}

def get_unit(unit_id):
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
        ws = wb['Units_Data']
        uid = unit_id.upper().replace(' ','')
        for row in ws.iter_rows(min_row=3, values_only=True):
            if row[0] and str(row[0]).upper().replace(' ','') == uid:
                return {'unit': row[0], 'type': str(row[1]).strip(),
                        'indoor': round(row[3],2),
                        'outdoor': round(row[6],2) if row[6] else 0,
                        'rate': round(row[9])}
    except Exception as e:
        log.error(e)
    return None

def fmt(n): return f"EGP {n:,.0f}"

def generate_pdf(unit, tenant, agent_key):
    agent = AGENTS.get(agent_key, {"name": agent_key, "phone": "", "email": ""})
    rate = unit['rate']
    sc = round(rate * 0.08)
    mf = round(rate * 0.05)
    cheq = rate + sc + mf
    dep = rate * 2
    parts = unit['unit'].split('-')
    bld = parts[0] if len(parts)>0 else ''
    flr = parts[1] if len(parts)>1 else ''
    unt = parts[2] if len(parts)>2 else ''
    html = f"""<!DOCTYPE html><html><head><meta charset="UTF-8">
<style>*{{box-sizing:border-box;margin:0;padding:0;}}body{{font-family:Georgia,serif;font-size:11px;}}
.hdr{{background:#2B1810;padding:22px 36px 18px;}}.brand{{font-size:22px;font-weight:700;letter-spacing:6px;color:#C9A84C;}}
.brand-sub{{font-size:7px;letter-spacing:3px;color:#B8A88A;}}.hdr-top{{display:flex;justify-content:space-between;align-items:flex-start;}}
.meta{{text-align:right;}}.meta-lbl{{font-size:7px;color:#B8A88A;}}.meta-date{{font-size:12px;font-weight:600;color:#F5ECD7;}}
.div{{width:100%;height:1px;background:#C9A84C44;margin:10px 0 8px;}}.prep{{font-size:7px;letter-spacing:3px;color:#C9A84C;margin-bottom:3px;}}
.tenant{{font-size:16px;font-weight:700;color:#F5ECD7;}}.body{{padding:18px 36px;}}.sec{{margin-bottom:13px;}}
.sec-lbl{{font-size:7px;letter-spacing:3px;text-transform:uppercase;color:#8B6040;font-weight:700;padding:9px 0 5px;border-top:1px solid #F0E6D6;margin-bottom:7px;}}
.utbl{{width:100%;border-collapse:collapse;border:1px solid #E8D8C0;}}.utbl th{{background:#2B1810;color:#C9A84C;font-size:6px;padding:6px 8px;text-align:center;}}
.utbl td{{background:#FBF6EE;color:#2B1810;font-size:9px;font-weight:600;padding:8px;text-align:center;border-right:1px solid #E8D8C0;}}
.ttbl{{width:100%;border-collapse:collapse;}}.ttbl tr{{border-bottom:1px solid #F0E6D6;}}.ttbl td{{padding:4px 0;font-size:10px;}}
.ttbl td:first-child{{color:#6B4A30;width:45%;font-size:9px;}}.ttbl td:last-child{{color:#2B1810;font-weight:600;}}
.pbox{{background:#FBF6EE;border:1px solid #E8D8C0;border-radius:5px;padding:12px;}}
.pgrid{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-bottom:7px;}}
.pitem{{background:white;border:1px solid #E8D8C0;border-radius:4px;padding:7px;text-align:center;}}
.pk{{font-size:6px;color:#8B6040;margin-bottom:2px;text-transform:uppercase;}}.pv{{font-size:11px;font-weight:700;color:#2B1810;}}
.pnote{{font-size:8px;color:#6B4A30;border-top:1px solid #E8D8C0;padding-top:6px;font-style:italic;}}
.p2{{page-break-before:always;padding:18px 36px;}}.gt{{font-size:8px;color:#6B4A30;line-height:1.7;}}.gt p{{margin-bottom:2px;}}
.agrid{{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-top:7px;}}
.slbl{{font-size:7px;color:#8B6040;margin-bottom:10px;text-transform:uppercase;}}
.sline{{border-bottom:1px solid #2B1810;margin-bottom:3px;height:12px;}}.shint{{font-size:7px;color:#6B4A30;margin-bottom:8px;}}
.ftr{{border-top:1px solid #E8D8C0;padding-top:8px;margin-top:12px;display:flex;justify-content:space-between;}}
.ftr-brand{{font-size:11px;font-weight:700;letter-spacing:3px;color:#2B1810;}}.ftr-agent{{text-align:right;}}
.ftr-albl{{font-size:6px;color:#8B6040;text-transform:uppercase;margin-bottom:2px;}}
.ftr-aname{{font-size:10px;font-weight:700;color:#2B1810;}}.ftr-asub{{font-size:7px;color:#6B4A30;line-height:1.5;}}</style></head><body>
<div class="hdr"><div class="hdr-top"><div><div class="brand">HIGHSTREET</div><div class="brand-sub">BY WEAVLONE · NEW CAIRO, EGYPT</div></div>
<div class="meta"><div class="meta-lbl">LEASING OFFER</div><div class="meta-date">21 May 2026</div><div style="font-size:9px;color:#B8A88A;">Valid until 28 May 2026</div></div></div>
<div class="div"></div><div class="prep">PREPARED FOR</div><div class="tenant">{tenant}</div></div>
<div class="body">
<div class="sec"><div class="sec-lbl">Unit Details</div>
<table class="utbl"><thead><tr><th>BUILDING</th><th>FLOOR</th><th>UNIT NO.</th><th>CONDITION</th><th>INDOOR</th><th>OUTDOOR</th></tr></thead>
<tbody><tr><td>{bld}</td><td>{flr}</td><td>{unt}</td><td>Core &amp; Shell</td><td>{unit['indoor']} sqm</td><td>{"—" if unit['outdoor']==0 else str(unit['outdoor'])+" sqm"}</td></tr></tbody></table></div>
<div class="sec"><div class="sec-lbl">Commercial Terms</div>
<table class="ttbl"><tr><td>Lease Term</td><td>5 Years</td></tr><tr><td>Base Rate</td><td>{fmt(rate)} / month</td></tr>
<tr><td>Maintenance Charge / Month</td><td>{fmt(sc)}</td></tr><tr><td>Marketing Charge / Month</td><td>{fmt(mf)}</td></tr>
<tr><td>Annual Escalation</td><td>10% starting Year 2</td></tr><tr><td>Rent-Free Period</td><td>3 Months (Fit-Out Grace Period)</td></tr></table></div>
<div class="sec"><div class="sec-lbl">Payment Structure</div><div class="pbox"><div class="pgrid">
<div class="pitem"><div class="pk">Payment Frequency</div><div class="pv">Monthly (12 Cheques)</div></div>
<div class="pitem"><div class="pk">Value Per Cheque</div><div class="pv">{fmt(cheq)}</div></div>
<div class="pitem"><div class="pk">Security Deposit</div><div class="pv">{fmt(dep)}</div></div></div>
<div class="pnote">* 3 months advance payment required upon signing.</div></div></div></div>
<div class="p2"><div class="sec"><div class="sec-lbl">General Terms</div><div class="gt">
<p>1. This offer is subject to management approval and formal lease agreement execution.</p>
<p>2. All figures are exclusive of VAT.</p><p>3. This offer is valid until 28 May 2026.</p>
<p>4. The security deposit is refundable at the end of the lease term, subject to unit condition.</p>
<p>5. Tenant is responsible for fit-out works in accordance with Highstreet fit-out guidelines.</p>
<p>6. Subletting or assignment of the lease requires prior written approval from management.</p></div></div>
<div class="sec"><div class="sec-lbl">Acceptance</div><div class="agrid">
<div><div class="slbl">Tenant Acceptance</div><div class="sline"></div><div class="shint">Signature</div>
<div class="sline"></div><div class="shint">Date</div><div class="sline"></div><div class="shint">Company Stamp</div></div>
<div><div class="slbl">Highstreet Leasing</div><div class="sline"></div><div class="shint">Authorized Signature</div>
<div class="sline"></div><div class="shint">Date</div></div></div></div>
<div class="ftr"><div><div class="ftr-brand">HIGHSTREET</div></div>
<div class="ftr-agent"><div class="ftr-albl">Prepared By</div><div class="ftr-aname">{agent['name']}</div>
<div class="ftr-asub">{agent['phone']}<br>{agent['email']}</div></div></div></div>
</body></html>"""
    html_path = f"/tmp/offer_{unit['unit'].replace('-','_')}.html"
    pdf_path = f"/tmp/offer_{unit['unit'].replace('-','_')}_{tenant.replace(' ','_')}.pdf"
    with open(html_path, 'w') as f:
        f.write(html)
    os.system(f'wkhtmltopdf --page-size A4 --margin-top 0 --margin-bottom 0 --margin-left 0 --margin-right 0 "{html_path}" "{pdf_path}" 2>/dev/null')
    return pdf_path if os.path.exists(pdf_path) else None

import re as _re
def parse_unit(text):
    m = _re.match(r'(B\d)[\s\-]?(GF|FF|SF|BA|THF)[\s\-]?(\d+)', text.upper())
    if m:
        return f"{m.group(1)}-{m.group(2)}-{m.group(3).zfill(2)}"
    return None

def handle_message(msg):
    chat_id = msg['chat']['id']
    user_id = msg['from']['id']
    user_name = msg['from'].get('first_name','') + ' ' + msg['from'].get('last_name','')
    text = msg.get('text','').strip()

    if text == '/start':
        send(chat_id, "🏢 *HIGHSTREET Leasing Bot*\n\nSend me a unit number (e.g. B1-GF-05)", parse_mode='Markdown')
        return

    session = sessions.get(user_id, {})

    if session.get('state') == 'ASK_TENANT':
        session['tenant'] = text
        session['state'] = 'ASK_AGENT'
        sessions[user_id] = session
        buttons = [[{"text": a['name'], "callback_data": f"agent_{k}"}] for k, a in AGENTS.items()]
        send(chat_id, "👤 Which agent?", reply_markup={"inline_keyboard": buttons})
        return

    unit_id = parse_unit(text)
    if unit_id:
        unit = get_unit(unit_id)
        if not unit:
            send(chat_id, f"❌ Unit *{unit_id}* not found.", parse_mode='Markdown')
            return
        sc = round(unit['rate'] * 0.08)
        mf = round(unit['rate'] * 0.05)
        sessions[user_id] = {'state': 'ASK_TENANT', 'unit': unit,
                             'requester_id': user_id, 'requester_name': user_name.strip(),
                             'chat_id': chat_id}
        send(chat_id,
             f"✅ *{unit['unit']}* — {unit['type']}\n📐 {unit['indoor']} sqm indoor\n💰 {fmt(unit['rate'])}/month\n🧾 Cheque: {fmt(unit['rate']+sc+mf)}\n\nPrepared for who?",
             parse_mode='Markdown')
    else:
        send(chat_id, "Send a unit number like B1-GF-05")

def handle_callback(cb):
    cb_id = cb['id']
    data = cb['data']
    msg = cb['message']
    chat_id = msg['chat']['id']
    msg_id = msg['message_id']
    user_id = cb['from']['id']
    answer_cb(cb_id)

    if data.startswith('agent_'):
        agent_key = data.replace('agent_', '')
        session = sessions.get(user_id, {})
        if not session:
            edit(chat_id, msg_id, "Session expired. Send unit number again.")
            return
        unit = session['unit']
        tenant = session['tenant']
        agent = AGENTS[agent_key]
        sc = round(unit['rate'] * 0.08)
        mf = round(unit['rate'] * 0.05)
        cheq = unit['rate'] + sc + mf
        key = f"{user_id}_{int(time.time())}"
        pending[key] = {'unit': unit, 'tenant': tenant, 'agent_key': agent_key,
                        'requester_id': user_id, 'chat_id': session['chat_id']}
        sessions.pop(user_id, None)
        approval_msg = (f"📋 *NEW OFFER REQUEST*\n\n👤 From: {session['requester_name']}\n"
                        f"🏢 Unit: *{unit['unit']}* — {unit['type']}\n📐 {unit['indoor']} sqm\n\n"
                        f"🎯 Tenant: *{tenant}*\n🤝 Agent: *{agent['name']}*\n\n"
                        f"💰 {fmt(unit['rate'])}/month\n🧾 Cheque: {fmt(cheq)}\n\n*Approve?*")
        kb = {"inline_keyboard": [[
            {"text": "✅ Approve", "callback_data": f"approve_{key}"},
            {"text": "❌ Reject",  "callback_data": f"reject_{key}"}]]}
        send(ADMIN_ID, approval_msg, reply_markup=kb, parse_mode='Markdown')
        edit(chat_id, msg_id, "⏳ Sent to management for approval.")

    elif data.startswith('approve_') or data.startswith('reject_'):
        if user_id != ADMIN_ID:
            return
        action, key = data.split('_', 1)
        p = pending.get(key)
        if not p:
            edit(chat_id, msg_id, "⚠️ Request expired.")
            return
        if action == 'reject':
            edit(chat_id, msg_id, "❌ Rejected.")
            send(p['chat_id'], "❌ Offer not approved by management.")
            pending.pop(key, None)
            return
        edit(chat_id, msg_id, "⏳ Generating PDF...")
        pdf_path = generate_pdf(p['unit'], p['tenant'], p['agent_key'])
        if pdf_path:
            filename = f"HS-{p['unit']['unit']}-{p['tenant'].replace(' ','_')}.pdf"
            send_doc(p['chat_id'], pdf_path, filename, f"✅ Offer Approved: {p['unit']['unit']} — {p['tenant']}")
            edit(chat_id, msg_id, f"✅ PDF sent for {p['unit']['unit']} — {p['tenant']}")
        else:
            edit(chat_id, msg_id, "⚠️ PDF generation failed.")
        pending.pop(key, None)

def main():
    log.info("🤖 Highstreet Bot starting (pure HTTP mode)...")
    offset = 0
    while True:
        try:
            res = api("getUpdates", {"offset": offset, "timeout": 30})
            updates = res.get('result', [])
            for u in updates:
                offset = u['update_id'] + 1
                if 'message' in u:
                    handle_message(u['message'])
                elif 'callback_query' in u:
                    handle_callback(u['callback_query'])
        except Exception as e:
            log.error(f"Main loop error: {e}")
            time.sleep(5)

if __name__ == '__main__':
    main()
