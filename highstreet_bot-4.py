#!/usr/bin/env python3
import os, logging, openpyxl, asyncio
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, ConversationHandler, filters, ContextTypes

BOT_TOKEN = "8818096411:AAHzieYk9WavFFKh2ikjkKl2Ajsf6dpGhx4"
ADMIN_ID  = 6944516887
EXCEL_PATH = "FINAL_SALES_PRICE_FF__2_.xlsx"

AGENTS = {
    "tamer":   {"name": "Tamer Shalaby",  "phone": "+201005394515",  "email": "tamer.origin@gmail.com"},
    "ahmed":   {"name": "Ahmed Elhakim",  "phone": "+201221414145",  "email": "aelhakim@virgo-eg.com"},
    "azmy":    {"name": "Azmy Mohamed",   "phone": "01050405780",    "email": "azmy932@gmail.com"},
    "ehab":    {"name": "Ehab Mustafa",   "phone": "01116064830",    "email": "Elbakhshwangyehab@gmail.com"},
    "mahmoud": {"name": "Mahmoud Shaker", "phone": "+201007000933",  "email": "mshaker@virgo-eg.com"},
    "noor":    {"name": "Noor Ilani",     "phone": "+32489647992",   "email": "Noor.ilani@outlook.com"},
}

ASK_TENANT, ASK_AGENT = range(2)
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

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
        logger.error(e)
    return None

def fmt(n): return f"EGP {n:,.0f}"

def generate_pdf(unit, tenant, agent_key):
    agent = AGENTS.get(agent_key.lower(), {"name": agent_key, "phone": "", "email": ""})
    rate = unit['rate']
    sc = round(rate * 0.08)
    mf = round(rate * 0.05)
    cheq = rate + sc + mf
    dep = rate * 2
    parts = unit['unit'].split('-')
    bld = parts[0] if len(parts) > 0 else ''
    flr = parts[1] if len(parts) > 1 else ''
    unt = parts[2] if len(parts) > 2 else ''
    html = f"""<!DOCTYPE html><html><head><meta charset="UTF-8">
<style>*{{box-sizing:border-box;margin:0;padding:0;}}body{{font-family:Georgia,serif;font-size:11px;}}
.hdr{{background:#2B1810;padding:22px 36px 18px;}}.hdr-top{{display:flex;justify-content:space-between;}}
.brand{{font-size:24px;font-weight:700;letter-spacing:8px;color:#C9A84C;}}.brand-sub{{font-size:7px;letter-spacing:4px;color:#B8A88A;}}
.meta{{text-align:right;}}.meta-lbl{{font-size:7px;letter-spacing:3px;color:#B8A88A;}}.meta-date{{font-size:13px;font-weight:600;color:#F5ECD7;}}
.meta-valid{{font-size:9px;color:#B8A88A;}}.div{{width:100%;height:1px;background:#C9A84C44;margin:10px 0 8px;}}
.prep{{font-size:7px;letter-spacing:3px;color:#C9A84C;margin-bottom:3px;}}.tenant{{font-size:17px;font-weight:700;color:#F5ECD7;}}
.body{{padding:18px 36px;}}.sec{{margin-bottom:13px;}}.sec-lbl{{font-size:7px;letter-spacing:3px;text-transform:uppercase;color:#8B6040;font-weight:700;padding:9px 0 5px;border-top:1px solid #F0E6D6;margin-bottom:7px;}}
.utbl{{width:100%;border-collapse:collapse;border:1px solid #E8D8C0;}}.utbl th{{background:#2B1810;color:#C9A84C;font-size:6px;padding:6px 8px;text-align:center;}}
.utbl td{{background:#FBF6EE;color:#2B1810;font-size:9px;font-weight:600;padding:8px;text-align:center;border-right:1px solid #E8D8C0;}}
.ttbl{{width:100%;border-collapse:collapse;}}.ttbl tr{{border-bottom:1px solid #F0E6D6;}}.ttbl td{{padding:4px 0;font-size:10px;}}
.ttbl td:first-child{{color:#6B4A30;width:45%;font-size:9px;}}.ttbl td:last-child{{color:#2B1810;font-weight:600;}}
.pbox{{background:#FBF6EE;border:1px solid #E8D8C0;border-radius:5px;padding:12px;}}.pgrid{{display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-bottom:7px;}}
.pitem{{background:white;border:1px solid #E8D8C0;border-radius:4px;padding:7px;text-align:center;}}.pk{{font-size:6px;color:#8B6040;margin-bottom:2px;text-transform:uppercase;}}
.pv{{font-size:11px;font-weight:700;color:#2B1810;}}.pnote{{font-size:8px;color:#6B4A30;border-top:1px solid #E8D8C0;padding-top:6px;font-style:italic;}}
.p2{{page-break-before:always;padding:18px 36px;}}.gt{{font-size:8px;color:#6B4A30;line-height:1.7;}}.gt p{{margin-bottom:2px;}}
.agrid{{display:grid;grid-template-columns:1fr 1fr;gap:20px;margin-top:7px;}}.slbl{{font-size:7px;color:#8B6040;margin-bottom:10px;text-transform:uppercase;}}
.sline{{border-bottom:1px solid #2B1810;margin-bottom:3px;height:12px;}}.shint{{font-size:7px;color:#6B4A30;margin-bottom:8px;}}
.ftr{{border-top:1px solid #E8D8C0;padding-top:8px;margin-top:12px;display:flex;justify-content:space-between;}}
.ftr-brand{{font-size:11px;font-weight:700;letter-spacing:3px;color:#2B1810;}}.ftr-agent{{text-align:right;}}
.ftr-albl{{font-size:6px;color:#8B6040;text-transform:uppercase;margin-bottom:2px;}}.ftr-aname{{font-size:10px;font-weight:700;color:#2B1810;}}
.ftr-asub{{font-size:7px;color:#6B4A30;line-height:1.5;}}</style></head><body>
<div class="hdr"><div class="hdr-top"><div><div class="brand">HIGHSTREET</div><div class="brand-sub">BY WEAVLONE · NEW CAIRO, EGYPT</div></div>
<div class="meta"><div class="meta-lbl">LEASING OFFER</div><div class="meta-date">21 May 2026</div><div class="meta-valid">Valid until 28 May 2026</div></div></div>
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

async def start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("🏢 *HIGHSTREET Leasing Bot*\n\nSend me a unit number (e.g. B1-GF-05)", parse_mode='Markdown')

async def handle_unit(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    import re
    text = update.message.text.strip().upper()
    match = re.match(r'(B\d)[\s\-]?(GF|FF|SF|BA|THF)[\s\-]?(\d+)', text)
    if not match:
        await update.message.reply_text("❌ Format: B1-GF-05")
        return ConversationHandler.END
    unit_id = f"{match.group(1)}-{match.group(2)}-{match.group(3).zfill(2)}"
    unit = get_unit(unit_id)
    if not unit:
        await update.message.reply_text(f"❌ Unit *{unit_id}* not found.", parse_mode='Markdown')
        return ConversationHandler.END
    ctx.user_data['unit'] = unit
    ctx.user_data['requester_id'] = update.effective_user.id
    ctx.user_data['requester_name'] = update.effective_user.full_name
    sc = round(unit['rate'] * 0.08)
    mf = round(unit['rate'] * 0.05)
    await update.message.reply_text(
        f"✅ *{unit['unit']}* — {unit['type']}\n📐 {unit['indoor']} sqm indoor\n💰 {fmt(unit['rate'])}/month\n🧾 Cheque: {fmt(unit['rate']+sc+mf)}\n\nPrepared for who?",
        parse_mode='Markdown')
    return ASK_TENANT

async def ask_tenant(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data['tenant'] = update.message.text.strip()
    buttons = [[InlineKeyboardButton(a['name'], callback_data=f"agent_{k}")] for k, a in AGENTS.items()]
    await update.message.reply_text("👤 Which agent?", reply_markup=InlineKeyboardMarkup(buttons))
    return ASK_AGENT

async def ask_agent(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    agent_key = query.data.replace('agent_', '')
    agent = AGENTS[agent_key]
    ctx.user_data['agent_key'] = agent_key
    unit = ctx.user_data['unit']
    tenant = ctx.user_data['tenant']
    sc = round(unit['rate'] * 0.08)
    mf = round(unit['rate'] * 0.05)
    cheq = unit['rate'] + sc + mf
    msg = (f"📋 *NEW OFFER REQUEST*\n\n👤 From: {ctx.user_data['requester_name']}\n"
           f"🏢 Unit: *{unit['unit']}* — {unit['type']}\n📐 {unit['indoor']} sqm\n\n"
           f"🎯 Tenant: *{tenant}*\n🤝 Agent: *{agent['name']}*\n\n"
           f"💰 {fmt(unit['rate'])}/month\n🧾 Cheque: {fmt(cheq)}\n\n*Approve?*")
    ctx.application.bot_data[f"pending_{query.from_user.id}"] = {
        'unit': unit, 'tenant': tenant, 'agent_key': agent_key,
        'requester_id': ctx.user_data['requester_id'], 'chat_id': query.message.chat_id}
    kb = InlineKeyboardMarkup([[
        InlineKeyboardButton("✅ Approve", callback_data=f"approve_{query.from_user.id}"),
        InlineKeyboardButton("❌ Reject", callback_data=f"reject_{query.from_user.id}")]])
    await ctx.application.bot.send_message(ADMIN_ID, msg, parse_mode='Markdown', reply_markup=kb)
    await query.edit_message_text("⏳ Sent to management for approval.")
    return ConversationHandler.END

async def handle_approval(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if update.effective_user.id != ADMIN_ID:
        return
    action, requester_id = query.data.split('_', 1)
    pending = ctx.application.bot_data.get(f"pending_{requester_id}")
    if not pending:
        await query.edit_message_text("⚠️ Request expired.")
        return
    if action == 'reject':
        await query.edit_message_text("❌ Rejected.")
        await ctx.application.bot.send_message(pending['chat_id'], "❌ Offer not approved.")
        return
    await query.edit_message_text("⏳ Generating PDF...")
    pdf_path = generate_pdf(pending['unit'], pending['tenant'], pending['agent_key'])
    if pdf_path:
        with open(pdf_path, 'rb') as f:
            await ctx.application.bot.send_document(
                chat_id=pending['chat_id'], document=f,
                filename=f"HS-{pending['unit']['unit']}-{pending['tenant'].replace(' ','_')}.pdf",
                caption=f"✅ *Offer Approved*\n{pending['unit']['unit']} — {pending['tenant']}",
                parse_mode='Markdown')
        await query.edit_message_text(f"✅ PDF sent.")
    else:
        await query.edit_message_text("⚠️ PDF generation failed.")

async def cancel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Cancelled.")
    return ConversationHandler.END

def main():
    app = (Application.builder()
           .token(BOT_TOKEN)
           .build())
    conv = ConversationHandler(
        entry_points=[MessageHandler(filters.TEXT & ~filters.COMMAND, handle_unit)],
        states={
            ASK_TENANT: [MessageHandler(filters.TEXT & ~filters.COMMAND, ask_tenant)],
            ASK_AGENT: [CallbackQueryHandler(ask_agent, pattern='^agent_')],
        },
        fallbacks=[CommandHandler('cancel', cancel)],
    )
    app.add_handler(CommandHandler('start', start))
    app.add_handler(conv)
    app.add_handler(CallbackQueryHandler(handle_approval, pattern='^(approve|reject)_'))
    logger.info("Bot starting...")
    app.run_polling(drop_pending_updates=True)

if __name__ == '__main__':
    main()
