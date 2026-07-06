from datetime import date
from io import BytesIO
from pathlib import Path
import html
import math
import re

colors = None
TA_CENTER = None
A4 = None
ParagraphStyle = None
getSampleStyleSheet = None
Image = None
KeepTogether = None
ListFlowable = None
ListItem = None
Paragraph = None
SimpleDocTemplate = None
Spacer = None
Table = None
TableStyle = None

BRAND_BLUE = None
BRAND_TEAL = None
INK = None
TEXT = None
MUTED = None
BG = None
LINE = None
SOFT = None
WHITE = None

PAGE_W = 1080
PAGE_H = 1780
MARGIN_X = 54
CONTENT_W = PAGE_W - MARGIN_X * 2
GAP = 18


def ensure_reportlab():
    global colors, TA_CENTER, A4, ParagraphStyle, getSampleStyleSheet
    global Image, KeepTogether, ListFlowable, ListItem, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    global BRAND_BLUE, BRAND_TEAL, INK, TEXT, MUTED, BG, LINE, SOFT, WHITE
    if colors is not None:
        return
    from reportlab.lib import colors as reportlab_colors
    from reportlab.lib.enums import TA_CENTER as reportlab_ta_center
    from reportlab.lib.styles import ParagraphStyle as reportlab_paragraph_style, getSampleStyleSheet as reportlab_get_styles
    from reportlab.platypus import Image as reportlab_image, KeepTogether as reportlab_keep_together, ListFlowable as reportlab_list_flowable, ListItem as reportlab_list_item, Paragraph as reportlab_paragraph, SimpleDocTemplate as reportlab_doc, Spacer as reportlab_spacer, Table as reportlab_table, TableStyle as reportlab_table_style

    colors = reportlab_colors
    TA_CENTER = reportlab_ta_center
    ParagraphStyle = reportlab_paragraph_style
    getSampleStyleSheet = reportlab_get_styles
    Image = reportlab_image
    KeepTogether = reportlab_keep_together
    ListFlowable = reportlab_list_flowable
    ListItem = reportlab_list_item
    Paragraph = reportlab_paragraph
    SimpleDocTemplate = reportlab_doc
    Spacer = reportlab_spacer
    Table = reportlab_table
    TableStyle = reportlab_table_style
    BRAND_BLUE = colors.HexColor('#1588D8')
    BRAND_TEAL = colors.HexColor('#18B9B7')
    INK = colors.HexColor('#142A35')
    TEXT = colors.HexColor('#43525C')
    MUTED = colors.HexColor('#6F8794')
    BG = colors.HexColor('#EEF6F8')
    LINE = colors.HexColor('#D8E9EF')
    SOFT = colors.HexColor('#F7FBFD')
    WHITE = colors.white


CJK_RE = re.compile(r'[\u3400-\u9fff\uf900-\ufaff]')


def clean_text(value, fallback=''):
    if value is None:
        return fallback
    text = str(value).strip()
    if text.lower() in {'none', 'null', 'undefined', 'nan'}:
        return fallback
    return text


def client_text(value, fallback=''):
    text = clean_text(value)
    if not text:
        return fallback
    if CJK_RE.search(text):
        return fallback
    return text


def trim_text(value, max_chars=120, fallback=''):
    text = client_text(value, fallback)
    text = re.sub(r'\s+', ' ', text).strip()
    if len(text) <= max_chars:
        return text
    cut = text[: max_chars - 1].rsplit(' ', 1)[0] or text[: max_chars - 1]
    return cut.rstrip('.,; ') + '.'


def split_lines(value, limit=None, max_chars=64):
    lines = []
    for line in clean_text(value).splitlines():
        line = trim_text(line.strip('-• \t'), max_chars, '')
        if line:
            lines.append(line)
    return lines[:limit] if limit else lines


def safe_filename_part(value):
    text = re.sub(r'[^A-Za-z0-9_-]+', '_', clean_text(value))
    return text.strip('_')


def format_price(value):
    text = clean_text(value)
    if not text:
        return 'To be confirmed'
    lowered = text.lower()
    if any(word in lowered for word in ['confirm', 'final', 'based']):
        return client_text(text, 'To be confirmed') or 'To be confirmed'
    compact = text.replace(',', '')
    match = re.search(r'(?:usd|us\$|\$)?\s*([0-9]+(?:\.[0-9]+)?)', compact, flags=re.I)
    if match:
        amount = float(match.group(1))
        return f'USD {int(amount):,}' if amount.is_integer() else f'USD {amount:,.2f}'
    return client_text(text, 'To be confirmed') or 'To be confirmed'


def poi_client_name(poi):
    return client_text(poi.get('name_en'), '')


def existing_local_path(value):
    path = clean_text(value)
    if not path:
        return ''
    candidate = Path(path)
    return str(candidate) if candidate.exists() else ''


def poi_image_path(poi):
    for key in ['image_path', 'image_placeholder']:
        path = existing_local_path(poi.get(key))
        if path:
            return path
    images = poi.get('images')
    if isinstance(images, list):
        for item in sorted(images, key=lambda item: not bool(item.get('isMain'))):
            path = existing_local_path(item.get('url'))
            if path:
                return path
    return ''


def first_day_image(day_pois):
    for poi in day_pois or []:
        path = poi_image_path(poi)
        if path:
            return path
    return ''


def first_route_image(day_pois):
    for pois in day_pois or []:
        path = first_day_image(pois)
        if path:
            return path
    return ''


def bullet_text_from_pois(day_pois, limit=3):
    names = [poi_client_name(poi) for poi in day_pois or [] if poi_client_name(poi)]
    return '\n'.join(names[:limit])


def build_pdf_draft_from_proposal(*, title, client_name, group_size, days, nights, route_cities, intro, itinerary, day_pois, includes, excludes, pricing_result, customer_type='Comprehensive Client'):
    quote = format_price((pricing_result or {}).get('client_quote'))
    route = client_text(route_cities, 'the selected route')
    client = client_text(client_name, 'Valued Client') or 'Valued Client'
    itinerary_rows = []
    for idx, row in enumerate(itinerary or [], start=1):
        pois = day_pois[idx - 1] if idx - 1 < len(day_pois or []) else []
        city = client_text(row.get('City'), '')
        theme = client_text(row.get('Theme'), 'Curated Health Travel')
        arrangement = trim_text(row.get('Arrangement'), 150, "Curated arrangements with medical, wellness, and concierge support.")
        tags = []
        for poi in pois[:3]:
            for tag in re.split(r'[,;/、，；]+', clean_text(poi.get('tags'))):
                tag = client_text(tag.strip(), '')
                if tag and tag not in tags:
                    tags.append(tag)
        itinerary_rows.append({
            'day': idx,
            'dayTitle': f'Day {idx}',
            'city': city,
            'title': trim_text(f'{city} - {theme}' if city else theme, 48, 'Curated Health Travel'),
            'description': arrangement,
            'pois': bullet_text_from_pois(pois),
            'tags': ' / '.join(tags[:4]) or 'Medical / Wellness / Concierge',
            'imagePath': first_day_image(pois),
        })
    medical_pois = [poi for pois in day_pois or [] for poi in pois if any(word in clean_text(poi.get('category')) + clean_text(poi.get('tags')) for word in ['医疗', '体检', '康养', 'Health', 'Medical', 'Wellness'])]
    medical_name = poi_client_name(medical_pois[0]) if medical_pois else 'Comprehensive Health Checkup'
    return {
        'cover': {
            'title': 'Personalized Health Travel Plan',
            'clientName': client,
            'publishedDate': date.today().isoformat(),
            'validity': '30 days',
            'heroTitle': client_text(title, 'Premium Medical Tourism Proposal') or 'Premium Medical Tourism Proposal',
            'subtitle': 'A private health travel journey combining comprehensive medical checkup, local concierge support, and curated city experiences.',
            'introText': trim_text(intro, 230, 'A private health travel journey combining comprehensive medical checkup, local concierge support, and curated city experiences.'),
            'quotedPrice': quote,
            'heroImagePath': first_route_image(day_pois),
        },
        'overview': {
            'tripType': 'Health Checkup + City Travel',
            'duration': f'{days} Days / {nights} Nights',
            'medicalCore': medical_name or 'Comprehensive Health Checkup',
            'service': 'Transfer, Interpreter, Hotel, Local Support',
            'summary': trim_text(f'This personalized health travel plan is designed for {client}, combining coordinated medical checkup services, private local support, and curated travel experiences across {route}. The itinerary balances healthcare efficiency, comfort, and selected local experiences into one easy-to-understand plan.', 330),
        },
        'recommendation': {
            'title': 'Consultant Recommendation',
            'body': 'Our recommendation is to keep the medical schedule clear and well-supported, while pairing it with light, high-quality local experiences. This structure gives the client confidence in the healthcare coordination, a comfortable daily rhythm, and a practical next step for confirmation.',
            'medicalCore': medical_name or 'Comprehensive Health Checkup',
            'highlights': 'Appointment coordination, interpretation support, private transfers, and a recovery-friendly travel rhythm.',
            'targetUsers': 'International clients seeking reliable medical coordination, private support, and a smooth health travel experience.',
            'serviceNote': 'HZH coordinates appointments, ground support, interpretation, transfers, hotels, and itinerary adjustments according to final availability.',
            'imagePath': first_route_image(day_pois),
        },
        'itinerary': itinerary_rows,
        'included': client_text(includes, '') or 'Health checkup appointment coordination\nPrivate transfer or local transportation support\nInterpreter / concierge support if selected\nHotel arrangement if included\nCurated wellness or city experiences',
        'notIncluded': client_text(excludes, '') or 'International flights\nVisa and travel insurance\nPersonal expenses\nOptional medical add-ons\nHotel upgrade or holiday surcharge',
        'pricing': {
            'clientFacingPrice': quote,
            'pricingNote': 'Final price may vary depending on travel dates, hotel availability, medical appointment schedule, and optional upgrades.',
            'finalNote': 'Final quotation should be confirmed after the client approves route, service scope, and travel dates.',
        },
        'nextSteps': {
            'items': 'Travel dates\nNumber of travelers\nHealth checkup package\nHotel preference\nDietary or mobility requirements\nFinal quotation confirmation',
            'cta': 'Contact HZH consultant to confirm availability',
            'note': 'Confirm the key details so HZH can check availability and prepare the final service arrangement.',
        },
    }


def pdf_styles():
    ensure_reportlab()
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle('H1', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=36, leading=40, textColor=INK, spaceAfter=9))
    styles.add(ParagraphStyle('H2', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=21, leading=24, textColor=INK, spaceAfter=7))
    styles.add(ParagraphStyle('CardTitle', parent=styles['Heading3'], fontName='Helvetica-Bold', fontSize=13.8, leading=16.5, textColor=INK, spaceAfter=4))
    styles.add(ParagraphStyle('Eyebrow', parent=styles['BodyText'], fontName='Helvetica-Bold', fontSize=9.5, leading=12, textColor=BRAND_BLUE, spaceAfter=6))
    styles.add(ParagraphStyle('Body', parent=styles['BodyText'], fontName='Helvetica', fontSize=12, leading=17.4, textColor=TEXT))
    styles.add(ParagraphStyle('Small', parent=styles['BodyText'], fontName='Helvetica', fontSize=10.6, leading=14.4, textColor=TEXT))
    styles.add(ParagraphStyle('Muted', parent=styles['BodyText'], fontName='Helvetica', fontSize=10, leading=13.6, textColor=MUTED))
    styles.add(ParagraphStyle('Price', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=30, leading=34, textColor=BRAND_BLUE, alignment=TA_CENTER))
    styles.add(ParagraphStyle('CTA', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=16, leading=20, textColor=WHITE, alignment=TA_CENTER))
    return styles


def paragraph(text, style, fallback=''):
    return Paragraph(html.escape(client_text(text, fallback)).replace('\n', '<br/>'), style)


def bullets(lines, style, limit=5):
    items = [ListItem(paragraph(line, style), leftIndent=8) for line in split_lines(lines, limit)]
    if not items:
        items = [ListItem(paragraph('To be confirmed', style), leftIndent=8)]
    return ListFlowable(items, bulletType='bullet', leftIndent=12, bulletFontSize=5)


def cover_image(path, width, height):
    path = existing_local_path(path)
    if not path:
        return None
    try:
        from PIL import Image as PilImage, ImageOps
        img = PilImage.open(path).convert('RGB')
        fitted = ImageOps.fit(
            img,
            (int(width * 2), int(height * 2)),
            method=PilImage.Resampling.LANCZOS,
            centering=(0.5, 0.5),
        )
        buf = BytesIO()
        fitted.save(buf, format='JPEG', quality=88)
        buf.seek(0)
        result = Image(buf, width=width, height=height)
    except Exception:
        result = Image(path, width=width, height=height)
    result.hAlign = 'CENTER'
    return result


def box(content, width, bg=None, padding=16, border=True, height=None):
    ensure_reportlab()
    bg = bg or WHITE
    t = Table([[content]], colWidths=[width], rowHeights=[height] if height else None)
    commands = [
        ('BACKGROUND', (0, 0), (-1, -1), bg),
        ('LEFTPADDING', (0, 0), (-1, -1), padding),
        ('RIGHTPADDING', (0, 0), (-1, -1), padding),
        ('TOPPADDING', (0, 0), (-1, -1), padding),
        ('BOTTOMPADDING', (0, 0), (-1, -1), padding),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]
    if border:
        commands.append(('BOX', (0, 0), (-1, -1), 0.7, LINE))
    t.setStyle(TableStyle(commands))
    return t


def image_box(path, width, height, label='Image Area'):
    ensure_reportlab()
    img = cover_image(path, width, height)
    if img:
        return img
    ph = Table([[paragraph(f'{label}\nCurated HZH visual', pdf_styles()['Muted'])]], colWidths=[width], rowHeights=[height])
    ph.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F7FBFC')),
        ('BOX', (0, 0), (-1, -1), 0.6, colors.HexColor('#DCEBF0')),
        ('INNERGRID', (0, 0), (-1, -1), 0.2, colors.HexColor('#EEF6F8')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ]))
    return ph


def add_bg(canvas, doc):
    canvas.saveState()
    canvas.setFillColor(BG)
    page_w, page_h = doc.pagesize
    canvas.rect(0, 0, page_w, page_h, fill=1, stroke=0)
    canvas.setFillColor(WHITE)
    canvas.rect(20, 20, page_w - 40, page_h - 40, fill=1, stroke=0)
    canvas.setStrokeColor(LINE)
    canvas.rect(20, 20, page_w - 40, page_h - 40, fill=0, stroke=1)
    canvas.restoreState()


def header(logo_path, cover, styles):
    logo = Image(str(logo_path), width=235, height=52) if logo_path and Path(logo_path).exists() else paragraph('HangZhou Health', styles['CardTitle'])
    meta = paragraph(f"{cover.get('title') or 'Personalized Health Travel Plan'}\nCreated for: {cover.get('clientName') or 'Valued Client'} / Published: {cover.get('publishedDate') or date.today().isoformat()} / Valid: {cover.get('validity') or '30 days'}", styles['Muted'])
    t = Table([[logo, meta]], colWidths=[280, CONTENT_W - 280])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), WHITE),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 18),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LINEBELOW', (0, 0), (-1, -1), 0.8, LINE),
        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
    ]))
    return t


def stat_card(label, value, styles, width):
    return box([paragraph(label, styles['Muted']), paragraph(value, styles['CardTitle'], 'To be confirmed')], width, WHITE, 13)


def generate_customer_pdf(pdf_draft, logo_path=None):
    ensure_reportlab()
    styles = pdf_styles()
    draft = pdf_draft or {}
    cover = draft.get('cover', {})
    overview = draft.get('overview', {})
    rec = draft.get('recommendation', {})
    pricing = draft.get('pricing', {})
    next_steps = draft.get('nextSteps', {})
    all_days = draft.get('itinerary') or []
    day_count = len(all_days)
    itinerary_cols = 4 if day_count <= 4 else (3 if day_count <= 6 else 4)
    itinerary_rows = max(1, math.ceil(day_count / itinerary_cols))
    if day_count <= 4:
        itinerary_card_h, itinerary_image_h = 292, 140
    elif day_count <= 6:
        itinerary_card_h, itinerary_image_h = 252, 112
    else:
        itinerary_card_h, itinerary_image_h = 238, 104
    page_h = PAGE_H + 160 + max(0, itinerary_rows - 1) * 55
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=(PAGE_W, page_h), rightMargin=MARGIN_X, leftMargin=MARGIN_X, topMargin=42, bottomMargin=42)
    story = []

    story.append(header(logo_path, cover, styles))
    story.append(Spacer(1, 26))

    hero_w = CONTENT_W - 326 - GAP
    hero_left = [
        paragraph('HZH MEDICAL TOURISM PROPOSAL', styles['Eyebrow']),
        paragraph(cover.get('heroTitle'), styles['H1'], 'Premium Medical Tourism Proposal'),
        paragraph(cover.get('subtitle') or cover.get('introText'), styles['Body'], 'A private health travel journey combining comprehensive medical checkup, local concierge support, and curated city experiences.'),
    ]
    price = format_price(cover.get('quotedPrice') or pricing.get('clientFacingPrice'))
    price_card = [
        paragraph('PACKAGE PRICE', styles['Muted']),
        paragraph(price, styles['Price']),
        paragraph('Final price is subject to availability and final service confirmation.', styles['Muted']),
        Spacer(1, 12),
        paragraph('Quick Scope', styles['CardTitle']),
        bullets('Medical coordination\nPrivate local support\nHotel / transfer planning', styles['Small'], 3),
    ]
    hero = Table([[box(hero_left, hero_w, SOFT, 26, height=230), box(price_card, 326, WHITE, 25, height=230)]], colWidths=[hero_w + GAP, 326])
    hero.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'), ('LEFTPADDING', (0, 0), (-1, -1), 0), ('RIGHTPADDING', (0, 0), (-1, -1), 0)]))
    story.append(hero)
    story.append(Spacer(1, 24))

    stat_w = (CONTENT_W - 3 * 14) / 4
    stats = Table([[stat_card('Trip Type', overview.get('tripType'), styles, stat_w), stat_card('Duration', overview.get('duration'), styles, stat_w), stat_card('Medical Core', overview.get('medicalCore'), styles, stat_w), stat_card('Service', overview.get('service'), styles, stat_w)]], colWidths=[stat_w + 14, stat_w + 14, stat_w + 14, stat_w])
    stats.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'), ('LEFTPADDING', (0, 0), (-1, -1), 0), ('RIGHTPADDING', (0, 0), (-1, -1), 0)]))
    story.append(stats)
    story.append(Spacer(1, 22))

    img_w = 430
    rec_w = CONTENT_W - img_w - GAP
    visual = Table([[image_box(cover.get('heroImagePath') or rec.get('imagePath'), img_w, 238, 'Curated Medical Travel Visual'), box([paragraph('Consultant Recommendation', styles['H2']), paragraph(trim_text(rec.get('body'), 235, 'This plan balances medical efficiency, comfort, and curated local experiences for a smooth client journey.'), styles['Body']), Spacer(1, 8), paragraph('Plan Summary', styles['CardTitle']), paragraph(trim_text(overview.get('summary'), 215, 'A concise medical tourism proposal with health checkup coordination, private support, selected experiences, and a clear next step.'), styles['Small'])], rec_w, WHITE, 19, height=238)]], colWidths=[img_w + GAP, rec_w])
    visual.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'), ('LEFTPADDING', (0, 0), (-1, -1), 0), ('RIGHTPADDING', (0, 0), (-1, -1), 0)]))
    story.append(visual)
    story.append(Spacer(1, 22))

    half_w = (CONTENT_W - GAP) / 2
    med = [paragraph('Medical Core', styles['H2']), paragraph(rec.get('medicalCore'), styles['CardTitle'], 'Comprehensive Health Checkup'), paragraph(rec.get('highlights'), styles['Small'], 'Appointment coordination, interpretation support, private transfers, and a recovery-friendly travel rhythm.')]
    service = [paragraph('Service Highlights', styles['H2']), bullets(draft.get('included'), styles['Small'], 5)]
    med_grid = Table([[box(med, half_w, WHITE, 18, height=176), box(service, half_w, SOFT, 18, height=176)]], colWidths=[half_w + GAP, half_w])
    med_grid.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'), ('LEFTPADDING', (0, 0), (-1, -1), 0), ('RIGHTPADDING', (0, 0), (-1, -1), 0)]))
    story.append(med_grid)
    story.append(Spacer(1, 22))

    story.append(paragraph('Simplified Itinerary', styles['H2']))
    story.append(Spacer(1, 4))
    days = all_days
    cols = itinerary_cols
    day_gap = 14
    day_w = (CONTENT_W - (cols - 1) * day_gap) / cols
    day_tables = []
    for start in range(0, len(days), cols):
        row_cells = []
        for day in days[start:start + cols]:
            description_limit = 74 if day_count > 6 else 92
            content = [
                image_box(day.get('imagePath'), day_w - 22, itinerary_image_h, 'Day Image'),
                Spacer(1, 8),
                paragraph(day.get('dayTitle'), styles['Eyebrow'], f"Day {day.get('day')}"),
                paragraph(day.get('title'), styles['CardTitle'], 'Curated Health Travel'),
                paragraph(trim_text(day.get('description'), description_limit, 'Curated medical, wellness, and concierge arrangements.'), styles['Small']),
            ]
            row_cells.append(box(content, day_w, WHITE, 11, height=itinerary_card_h))
        if row_cells:
            col_widths = [day_w + day_gap] * (len(row_cells) - 1) + [day_w]
            row_table = Table([row_cells], colWidths=col_widths, rowHeights=[itinerary_card_h])
            row_table.hAlign = 'LEFT'
            row_table.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'), ('LEFTPADDING', (0, 0), (-1, -1), 0), ('RIGHTPADDING', (0, 0), (-1, -1), 0), ('TOPPADDING', (0, 0), (-1, -1), 0), ('BOTTOMPADDING', (0, 0), (-1, -1), 0)]))
            day_tables.append(row_table)
    if not day_tables:
        day_tables = [box([paragraph('Itinerary to be confirmed', styles['CardTitle']), paragraph('Daily travel details can be confirmed after travel dates are selected.', styles['Small'])], CONTENT_W, WHITE, 16, height=130)]
    for row_index, day_table in enumerate(day_tables):
        story.append(day_table)
        if row_index < len(day_tables) - 1:
            story.append(Spacer(1, 14))
    story.append(Spacer(1, 24))

    third_w = (CONTENT_W - 2 * GAP) / 3
    scope = [paragraph('Not Included / To Confirm', styles['CardTitle']), bullets(draft.get('notIncluded'), styles['Small'], 5)]
    confirmation_notes = '\n'.join([line for line in [pricing.get('finalNote'), 'Hotel, transfer, and medical appointment availability are subject to final confirmation.', 'HZH consultant will coordinate the final service arrangement.', 'Optional upgrades can be adjusted before final quotation.'] if clean_text(line)])
    confirm_box = [paragraph('Final Confirmation Notes', styles['CardTitle']), bullets(confirmation_notes, styles['Small'], 5)]
    next_box = [paragraph('Next Steps', styles['CardTitle']), bullets(next_steps.get('items'), styles['Small'], 5), Spacer(1, 8), paragraph(next_steps.get('note') or next_steps.get('cta'), styles['Muted'], 'Please contact your HZH consultant for final confirmation.')]
    bottom = Table([[box(scope, third_w, WHITE, 16, height=214), box(confirm_box, third_w, WHITE, 16, height=214), box(next_box, third_w, SOFT, 16, height=214)]], colWidths=[third_w + GAP, third_w + GAP, third_w], rowHeights=[214])
    bottom.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP'), ('LEFTPADDING', (0, 0), (-1, -1), 0), ('RIGHTPADDING', (0, 0), (-1, -1), 0)]))
    story.append(bottom)

    doc.build(story, onFirstPage=add_bg, onLaterPages=add_bg)
    return buf.getvalue()


def customer_pdf_filename(pdf_draft):
    name = safe_filename_part((pdf_draft or {}).get('cover', {}).get('clientName', ''))
    today = date.today().strftime('%Y%m%d')
    return f"HZH_Health_Travel_Proposal_{name + '_' if name and name.lower() != 'valued_client' else ''}{today}.pdf"
