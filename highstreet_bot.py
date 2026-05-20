#!/usr/bin/env python3
"""
HIGHSTREET Leasing Bot
- Anyone sends a unit number → bot asks for tenant name
- Bot pulls pricing from Excel and sends details to ADMIN for approval
- Admin approves → bot generates PDF and sends it back
"""

import os, logging, base64, openpyxl
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ConversationHandler, filters, ContextTypes
)

# ── CONFIG ──────────────────────────────────────────────────────────────────
BOT_TOKEN   = "8818096411:AAHzieYk9WavFFKh2ikjkKl2Ajsf6dpGhx4"
ADMIN_ID    = 6944516887
EXCEL_PATH  = "FINAL_SALES_PRICE_FF__2_.xlsx"

# Agents directory
AGENTS = {
    "tamer":   {"name": "Tamer Shalaby",  "phone": "+201005394515",  "email": "tamer.origin@gmail.com"},
    "ahmed":   {"name": "Ahmed Elhakim",  "phone": "+201221414145",  "email": "aelhakim@virgo-eg.com"},
    "azmy":    {"name": "Azmy Mohamed",   "phone": "01050405780",    "email": "azmy932@gmail.com"},
    "ehab":    {"name": "Ehab Mustafa",   "phone": "01116064830",    "email": "Elbakhshwangyehab@gmail.com"},
    "mahmoud": {"name": "Mahmoud Shaker", "phone": "+201007000933",  "email": "mshaker@virgo-eg.com"},
    "noor":    {"name": "Noor Ilani",     "phone": "+32489647992",   "email": "Noor.ilani@outlook.com"},
}

# Conversation states
ASK_TENANT, ASK_AGENT, WAIT_APPROVAL = range(3)

logging.basicConfig(level=logging.INFO)
log = logging.getLogger(__name__)

# ── EXCEL LOOKUP ─────────────────────────────────────────────────────────────
def get_unit(unit_id):
    try:
        wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
        ws = wb['Units_Data']
        for row in ws.iter_rows(min_row=3, values_only=True):
            if row[0] and str(row[0]).upper().replace(' ', '') == unit_id.upper().replace(' ', '').replace('-', '').replace('GF', 'GF'):
                return {
                    'unit':    row[0],
                    'type':    str(row[1]).strip(),
                    'floor':   row[2],
                    'indoor':  round(row[3], 2),
                    'outdoor': round(row[6], 2) if row[6] else 0,
                    'rate':    round(row[9]),
                }
        # Try flexible match
        wb2 = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
        ws2 = wb2['Units_Data']
        for row in ws2.iter_rows(min_row=3, values_only=True):
            if row[0] and unit_id.upper().replace('-','').replace(' ','') in str(row[0]).upper().replace('-',''):
                return {
                    'unit':    row[0],
                    'type':    str(row[1]).strip(),
                    'floor':   row[2],
                    'indoor':  round(row[3], 2),
                    'outdoor': round(row[6], 2) if row[6] else 0,
                    'rate':    round(row[9]),
                }
    except Exception as e:
        log.error(f"Excel error: {e}")
    return None

def fmt(n): return f"EGP {n:,.0f}"

# ── PDF GENERATOR ─────────────────────────────────────────────────────────────
def generate_pdf(unit, tenant, agent_key, offer_date="21 May 2026", valid_until="28 May 2026"):
    agent = AGENTS.get(agent_key.lower(), {"name": agent_key, "phone": "", "email": ""})
    rate = unit['rate']
    sc   = round(rate * 0.08)
    mf   = round(rate * 0.05)
    cheq = rate + sc + mf
    dep  = rate * 2

    css = """
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'Palatino Linotype',Palatino,Georgia,serif;background:#fff;font-size:11px;}
.doc{max-width:794px;margin:0 auto;}
.hdr{background:#2B1810;padding:22px 36px 18px;}
.hdr-top{display:flex;justify-content:space-between;align-items:flex-start;}
.brand{font-size:24px;font-weight:700;letter-spacing:8px;color:#C9A84C;margin-bottom:3px;}
.brand-sub{font-size:7px;letter-spacing:4px;color:#B8A88A;}
.meta{text-align:right;}
.meta-lbl{font-size:7px;letter-spacing:3px;color:#B8A88A;}
.meta-date{font-size:13px;font-weight:600;color:#F5ECD7;}
.meta-valid{font-size:9px;color:#B8A88A;margin-top:2px;}
.divider{width:100%;height:1px;background:#C9A84C44;margin:10px 0 8px;}
.prep-lbl{font-size:7px;letter-spacing:3px;color:#C9A84C;margin-bottom:3px;}
.tenant{font-size:17px;font-weight:700;color:#F5ECD7;}
.body{padding:18px 36px;}
.sec{margin-bottom:13px;}
.sec-lbl{font-size:7px;letter-spacing:3px;text-transform:uppercase;color:#8B6040;font-weight:700;padding:9px 0 5px;border-top:1px solid #F0E6D6;margin-bottom:7px;}
.utbl{width:100%;border-collapse:collapse;border:1px solid #E8D8C0;}
.utbl th{background:#2B1810;color:#C9A84C;font-size:6px;letter-spacing:1px;padding:6px 8px;text-align:center;font-weight:700;}
.utbl td{background:#FBF6EE;color:#2B1810;font-size:9px;font-weight:600;padding:8px;text-align:center;border-right:1px solid #E8D8C0;}
.ttbl{width:100%;border-collapse:collapse;}
.ttbl tr{border-bottom:1px solid #F0E6D6;}
.ttbl td{padding:4px 0;font-size:10px;}
.ttbl td:first-child{color:#6B4A30;width:45%;font-size:9px;}
.ttbl td:last-child{color:#2B1810;font-weight:600;}
.pbox{background:#FBF6EE;border:1px solid #E8D8C0;border-radius:5px;padding:12px;}
.pgrid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-bottom:7px;}
.pitem{background:white;border:1px solid #E8D8C0;border-radius:4px;padding:7px;text-align:center;}
.pk{font-size:6px;letter-spacing:1px;color:#8B6040;margin-bottom:2px;text-transform:uppercase;}
.pv{font-size:11px;font-weight:700;color:#2B1810;}
.pnote{font-size:8px;color:#6B4A30;border-top:1px solid #E8D8C0;padding-top:6px;font-style:italic;}
.gt{font-size:8px;color:#6B4A30;line-height:1.7;}
.gt p{margin-bottom:2px;}
.agrid{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-top:7px;}
.slbl{font-size:7px;letter-spacing:2px;color:#8B6040;margin-bottom:10px;text-transform:uppercase;}
.sline{border-bottom:1px solid #2B1810;margin-bottom:3px;height:12px;}
.shint{font-size:7px;color:#6B4A30;margin-bottom:8px;}
.ftr{border-top:1px solid #E8D8C0;padding-top:8px;margin-top:12px;display:flex;justify-content:space-between;}
.ftr-brand{font-size:11px;font-weight:700;letter-spacing:3px;color:#2B1810;}
.ftr-agent{text-align:right;}
.ftr-albl{font-size:6px;letter-spacing:1px;color:#8B6040;text-transform:uppercase;margin-bottom:2px;}
.ftr-aname{font-size:10px;font-weight:700;color:#2B1810;}
.ftr-asub{font-size:7px;color:#6B4A30;line-height:1.5;}
.p2{page-break-before:always;padding:18px 36px;}
"""

    html = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"><style>{css}</style></head><body>
<div class="doc">
<div class="hdr">
  <div class="hdr-top">
    <div><div class="brand">HIGHSTREET</div><div class="brand-sub">BY WEAVLONE &nbsp;·&nbsp; NEW CAIRO, EGYPT</div></div>
    <div class="meta">
      <div class="meta-lbl">LEASING OFFER</div>
      <div class="meta-date">{offer_date}</div>
      <div class="meta-valid">Valid until {valid_until}</div>
    </div>
  </div>
  <div class="divider"></div>
  <div class="prep-lbl">PREPARED FOR</div>
  <div class="tenant">{tenant}</div>
</div>
<div class="body">
  <div class="sec">
    <div class="sec-lbl">Unit Details</div>
    <table class="utbl">
      <thead><tr><th>BUILDING</th><th>FLOOR</th><th>UNIT NO.</th><th>CONDITION</th><th>INDOOR AREA</th><th>OUTDOOR AREA</th></tr></thead>
      <tbody><tr>
        <td>{unit['unit'].split('-')[0]}</td>
        <td>{unit['unit'].split('-')[1]}</td>
        <td>{unit['unit'].split('-')[2]}</td>
        <td>Core & Shell</td>
        <td>{unit['indoor']} sqm</td>
        <td>{"— " if unit['outdoor']==0 else str(unit['outdoor'])+" sqm"}</td>
      </tr></tbody>
    </table>
  </div>
  <div class="sec">
    <div class="sec-lbl">Commercial Terms</div>
    <table class="ttbl">
      <tr><td>Lease Term</td><td>5 Years</td></tr>
      <tr><td>Base Rate</td><td>{fmt(rate)} / month</td></tr>
      <tr><td>Maintenance Charge / Month</td><td>{fmt(sc)}</td></tr>
      <tr><td>Marketing Charge / Month</td><td>{fmt(mf)}</td></tr>
      <tr><td>Annual Escalation</td><td>10% starting Year 2</td></tr>
      <tr><td>Rent-Free Period</td><td>3 Months (Fit-Out Grace Period)</td></tr>
    </table>
  </div>
  <div class="sec">
    <div class="sec-lbl">Payment Structure</div>
    <div class="pbox">
      <div class="pgrid">
        <div class="pitem"><div class="pk">Payment Frequency</div><div class="pv">Monthly (12 Cheques)</div></div>
        <div class="pitem"><div class="pk">Value Per Cheque</div><div class="pv">{fmt(cheq)}</div></div>
        <div class="pitem"><div class="pk">Security Deposit</div><div class="pv">{fmt(dep)}</div></div>
      </div>
      <div class="pnote">* 3 months advance payment required upon signing.</div>
    </div>
  </div>
</div>
<div class="p2">
  <div class="sec">
    <div class="sec-lbl">General Terms</div>
    <div class="gt">
      <p>1. This offer is subject to management approval and formal lease agreement execution.</p>
      <p>2. All figures are exclusive of VAT.</p>
      <p>3. This offer is valid until {valid_until} and may be withdrawn or amended thereafter.</p>
      <p>4. The security deposit is refundable at the end of the lease term, subject to unit condition.</p>
      <p>5. Tenant is responsible for fit-out works in accordance with Highstreet fit-out guidelines.</p>
      <p>6. Subletting or assignment of the lease requires prior written approval from management.</p>
    </div>
  </div>
  <div class="sec">
    <div class="sec-lbl">Acceptance</div>
    <div class="agrid">
      <div>
        <div class="slbl">Tenant Acceptance</div>
        <div class="sline"></div><div class="shint">Signature</div>
        <div class="sline"></div><div class="shint">Date</div>
        <div class="sline"></div><div class="shint">Company Stamp</div>
      </div>
      <div>
        <div class="slbl">Highstreet Leasing</div>
        <div class="sline"></div><div class="shint">Authorized Signature</div>
        <div class="sline"></div><div class="shint">Date</div>
      </div>
    </div>
  </div>
  <div class="ftr">
    <div><div class="ftr-brand">HIGHSTREET</div></div>
    <div class="ftr-agent">
      <div class="ftr-albl">Prepared By</div>
      <div class="ftr-aname">{agent['name']}</div>
      <div class="ftr-asub">{agent['phone']}<br>{agent['email']}</div>
    </div>
  </div>
</div>
</div>
</body></html>"""

    html_path = f"/tmp/offer_{unit['unit'].replace('-','_')}.html"
    pdf_path  = f"/tmp/offer_{unit['unit'].replace('-','_')}_{tenant.replace(' ','_')}.pdf"

    with open(html_path, 'w') as f:
        f.write(html)

    os.system(f'wkhtmltopdf --page-size A4 --margin-top 0 --margin-bottom 0 --margin-left 0 --margin-right 0 --enable-local-file-access "{html_path}" "{pdf_path}" 2>/dev/null')
    return pdf_path if os.path.exists(pdf_path) else None

# ── BOT HANDLERS ─────────────────────────────────────────────────────────────

async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "🏢 *HIGHSTREET Leasing Bot*\n\nSend me a unit number (e.g. B1-GF-05) and I'll prepare a leasing offer.",
        parse_mode='Markdown'
    )

async def handle_unit(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip().upper()

    # Normalize: B1GF05 → B1-GF-05
    import re
    match = re.match(r'(B\d)[\s\-]?(GF|FF|SF|BA|THF)[\s\-]?(\d+)', text)
    if not match:
        await update.message.reply_text("❌ Format not recognized. Try: B1-GF-05 or B2 GF 14")
        return ConversationHandler.END

    unit_id = f"{match.group(1)}-{match.group(2)}-{match.group(3).zfill(2)}"
    unit = get_unit(unit_id)

    if not unit:
        await update.message.reply_text(f"❌ Unit *{unit_id}* not found in the database.", parse_mode='Markdown')
        return ConversationHandler.END

    ctx.user_data['unit'] = unit
    ctx.user_data['requester_id'] = update.effective_user.id
    ctx.user_data['requester_name'] = update.effective_user.full_name

    sc = round(unit['rate'] * 0.08)
    mf = round(unit['rate'] * 0.05)
    cheq = unit['rate'] + sc + mf

    await update.message.reply_text(
        f"✅ *{unit['unit']}* — {unit['type']}\n"
        f"📐 Indoor: {unit['indoor']} sqm | Outdoor: {unit['outdoor'] or '—'} sqm\n"
        f"💰 Base Rate: {fmt(unit['rate'])}/month\n"
        f"🧾 Cheque Value: {fmt(cheq)}\n\n"
        f"Who is this offer prepared for?\n_(Type the tenant/company name)_",
        parse_mode='Markdown'
    )
    return ASK_TENANT

async def ask_tenant(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['tenant'] = update.message.text.strip()

    # Build agent keyboard
    buttons = [[InlineKeyboardButton(a['name'], callback_data=f"agent_{k}")] for k, a in AGENTS.items()]
    await update.message.reply_text(
        "👤 Which agent is handling this offer?",
        reply_markup=InlineKeyboardMarkup(buttons)
    )
    return ASK_AGENT

async def ask_agent(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    agent_key = query.data.replace('agent_', '')
    agent = AGENTS[agent_key]
    ctx.user_data['agent_key'] = agent_key

    unit   = ctx.user_data['unit']
    tenant = ctx.user_data['tenant']
    sc = round(unit['rate'] * 0.08)
    mf = round(unit['rate'] * 0.05)
    cheq = unit['rate'] + sc + mf
    dep  = unit['rate'] * 2

    # Send approval request to admin
    msg = (
        f"📋 *NEW OFFER REQUEST*\n\n"
        f"👤 Requested by: {ctx.user_data['requester_name']}\n"
        f"🏢 Unit: *{unit['unit']}* — {unit['type']}\n"
        f"📐 Indoor: {unit['indoor']} sqm | Outdoor: {unit['outdoor'] or '—'} sqm\n\n"
        f"🎯 Tenant: *{tenant}*\n"
        f"🤝 Agent: *{agent['name']}*\n\n"
        f"💰 Base Rate: {fmt(unit['rate'])}/month\n"
        f"🔧 Maintenance: {fmt(sc)}/month\n"
        f"📣 Marketing: {fmt(mf)}/month\n"
        f"🧾 Cheque Value: {fmt(cheq)}\n"
        f"🔒 Security Deposit: {fmt(dep)}\n\n"
        f"*Approve to generate and send PDF?*"
    )

    approve_btn = InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ Approve", callback_data=f"approve_{query.from_user.id}"),
        InlineKeyboardButton("❌ Reject",  callback_data=f"reject_{query.from_user.id}"),
    ]])

    # Store pending request in bot data
    ctx.application.bot_data[f"pending_{query.from_user.id}"] = {
        'unit': unit, 'tenant': tenant, 'agent_key': agent_key,
        'requester_id': ctx.user_data['requester_id'],
        'chat_id': query.message.chat_id
    }

    await ctx.application.bot.send_message(ADMIN_ID, msg, parse_mode='Markdown', reply_markup=approve_btn)
    await query.edit_message_text(f"⏳ Offer for *{unit['unit']} — {tenant}* sent to management for approval. You'll receive the PDF once approved.", parse_mode='Markdown')
    return ConversationHandler.END

async def handle_approval(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    if update.effective_user.id != ADMIN_ID:
        await query.answer("Not authorized.", show_alert=True)
        return

    action, requester_id = query.data.split('_', 1)
    pending = ctx.application.bot_data.get(f"pending_{requester_id}")

    if not pending:
        await query.edit_message_text("⚠️ Request expired or not found.")
        return

    if action == 'reject':
        await query.edit_message_text("❌ Offer rejected.")
        await ctx.application.bot.send_message(pending['chat_id'], "❌ Your offer request was not approved by management.")
        return

    # Generate PDF
    await query.edit_message_text("⏳ Generating PDF...")
    pdf_path = generate_pdf(pending['unit'], pending['tenant'], pending['agent_key'])

    if pdf_path:
        await ctx.application.bot.send_document(
            chat_id=pending['chat_id'],
            document=open(pdf_path, 'rb'),
            filename=f"HS-{pending['unit']['unit']}-{pending['tenant'].replace(' ', '_')}.pdf",
            caption=f"✅ *Leasing Offer Approved*\n{pending['unit']['unit']} — {pending['tenant']}",
            parse_mode='Markdown'
        )
        await query.edit_message_text(f"✅ PDF sent for {pending['unit']['unit']} — {pending['tenant']}")
    else:
        await query.edit_message_text("⚠️ PDF generation failed.")

async def cancel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Cancelled.")
    return ConversationHandler.END

# ── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    app = Application.builder().token(BOT_TOKEN).build()

    conv = ConversationHandler(
        entry_points=[MessageHandler(filters.TEXT & ~filters.COMMAND, handle_unit)],
        states={
            ASK_TENANT: [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_tenant)],
            ASK_AGENT:  [CallbackQueryHandler(ask_agent, pattern='^agent_')],
        },
        fallbacks=[CommandHandler('cancel', cancel)],
    )

    app.add_handler(CommandHandler('start', start))
    app.add_handler(conv)
    app.add_handler(CallbackQueryHandler(handle_approval, pattern='^(approve|reject)_'))

    print("🤖 Highstreet Leasing Bot is running...")
    app.run_polling(drop_pending_updates=True)

if __name__ == '__main__':
    main()
