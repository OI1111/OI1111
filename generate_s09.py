from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

# ── helpers ────────────────────────────────────────────────────────────────

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def add_heading(doc, text, level=1, color='000000'):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(14 if level == 1 else 12 if level == 2 else 11)
    run.font.color.rgb = RGBColor.from_string(color)
    return p

def add_body(doc, text, bold=False, italic=False, size=10):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    return p

def add_labelled_block(doc, label, text, label_color='1F3864'):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    lr = p.add_run(label + '  ')
    lr.bold = True
    lr.font.size = Pt(10)
    lr.font.color.rgb = RGBColor.from_string(label_color)
    tr = p.add_run(text)
    tr.font.size = Pt(10)
    return p

def add_box_block(doc, label, text, bg='F2F2F2'):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.style = 'Table Grid'
    cell = tbl.rows[0].cells[0]
    set_cell_bg(cell, bg)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    lr = p.add_run(label + '  ')
    lr.bold = True
    lr.font.size = Pt(10)
    tr = p.add_run(text)
    tr.font.size = Pt(10)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def add_two_col_table(doc, headers, rows, col_widths=None):
    n = len(headers)
    tbl = doc.add_table(rows=1+len(rows), cols=n)
    tbl.style = 'Table Grid'
    tbl.alignment = WD_TABLE_ALIGNMENT.LEFT
    # header row
    for i, h in enumerate(headers):
        cell = tbl.rows[0].cells[i]
        set_cell_bg(cell, '1F3864')
        p = cell.paragraphs[0]
        run = p.add_run(h)
        run.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        run.font.size = Pt(9)
    # data rows
    for ri, row in enumerate(rows):
        for ci, val in enumerate(row):
            cell = tbl.rows[ri+1].cells[ci]
            if ri % 2 == 0:
                set_cell_bg(cell, 'EBF0F7')
            p = cell.paragraphs[0]
            run = p.add_run(str(val))
            run.font.size = Pt(9)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def add_divider(doc):
    p = doc.add_paragraph('─' * 80)
    p.runs[0].font.color.rgb = RGBColor(0xCC, 0xCC, 0xCC)
    p.runs[0].font.size = Pt(8)

def add_theory_header_sb(doc, label, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(2)
    lr = p.add_run(f'Theory | {label} |  ')
    lr.bold = True
    lr.font.size = Pt(11)
    lr.font.color.rgb = RGBColor.from_string('7F7F7F')
    tr = p.add_run(title)
    tr.bold = True
    tr.font.size = Pt(13)
    tr.font.color.rgb = RGBColor.from_string('1F3864')
    add_divider(doc)

def add_theory_header_tm(doc, minutes, label, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after = Pt(2)
    lr = p.add_run(f'{minutes} MIN | {label} | ')
    lr.bold = True
    lr.font.size = Pt(11)
    lr.font.color.rgb = RGBColor.from_string('7F7F7F')
    tr = p.add_run(title)
    tr.bold = True
    tr.font.size = Pt(13)
    tr.font.color.rgb = RGBColor.from_string('1F3864')
    add_divider(doc)

def add_quote(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.left_indent = Inches(0.3)
    run = p.add_run(f'"{text}"')
    run.italic = True
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor.from_string('1F3864')

def add_build_header(doc, number, title):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(f'Build {number:02d} | {title}')
    run.bold = True
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor.from_string('1F3864')
    add_divider(doc)

def add_reflect(doc, questions):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    run = p.add_run('REFLECT')
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor.from_string('1F3864')
    for q in questions:
        p2 = doc.add_paragraph(q, style='List Number')
        p2.runs[0].font.size = Pt(10)
        p2.paragraph_format.space_after = Pt(4)

def add_transition(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(4)
    lr = p.add_run('TRANSITION  ')
    lr.bold = True
    lr.italic = True
    lr.font.size = Pt(10)
    lr.font.color.rgb = RGBColor.from_string('1F3864')
    tr = p.add_run(text)
    tr.italic = True
    tr.font.size = Pt(10)

# ═══════════════════════════════════════════════════════════════════════════
#  STUDENT BOOK
# ═══════════════════════════════════════════════════════════════════════════

def build_sb():
    doc = Document()
    # page margins
    for section in doc.sections:
        section.top_margin = Inches(0.9)
        section.bottom_margin = Inches(0.9)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # ── COVER BLOCK ────────────────────────────────────────────────────────
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('Bar-Skills')
    r.bold = True; r.font.size = Pt(18)
    r.font.color.rgb = RGBColor.from_string('1F3864')

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('Craft Bartender Course')
    r.font.size = Pt(12)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('SESSION 09')
    r.bold = True; r.font.size = Pt(22)
    r.font.color.rgb = RGBColor.from_string('1F3864')

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('Menu Engineering')
    r.bold = True; r.font.size = Pt(16)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('Cost, Design, and the Profitable Bar')
    r.font.size = Pt(12); r.italic = True

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('STUDENT BOOK')
    r.bold = True; r.font.size = Pt(12)

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('bar-skills.com | info@bar-skills.com')
    r.font.size = Pt(9)
    r.font.color.rgb = RGBColor.from_string('7F7F7F')

    add_divider(doc)

    # ── WELCOME ────────────────────────────────────────────────────────────
    add_heading(doc, 'Welcome to Session 09', level=2)
    add_body(doc,
        'Sessions 01 through 08 gave you the technique, the classics, the balance equation, the original build, '
        'speed under pressure, and the capstone. Session 09 asks a different question: once you know how to build it — '
        'how do you sell it? A menu is not a list of drinks. It is the most powerful sales tool behind the bar, and '
        'most bars never use it correctly.')

    add_body(doc,
        'This session introduces menu engineering: the discipline of designing a drink program that guides guest '
        'decisions, maximises margin, and reflects the brand. You will read two real menus, identify exactly what '
        'is working and what is not, and then go behind the bar to prove the principle in four builds.')

    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    r = p.add_run('TRAINING MODE | SERVICE MODE')
    r.bold = True; r.font.size = Pt(10)
    r.font.color.rgb = RGBColor.from_string('1F3864')

    add_body(doc,
        'You will build four drinks today: a Black Russian, a White Russian, a Classic Cosmopolitan from a real '
        'bar menu, and an Elevated Cosmopolitan that demonstrates how three deliberate modifications can justify '
        'a higher price point while improving your margin. The same precision standard applies. Measure everything. '
        'Straw-taste before you present.')

    add_divider(doc)

    # ══════════════════════════════════════════════════
    # T1 — What Is Menu Engineering?
    # ══════════════════════════════════════════════════
    add_theory_header_sb(doc, 'T1', 'What Is Menu Engineering?')

    add_labelled_block(doc, 'CORE PRINCIPLE',
        'Menu engineering is the discipline of designing a drink list to maximise both guest satisfaction and profitability. '
        'It is not decoration — it is strategy.')

    add_heading(doc, 'WHAT YOU NEED TO REMEMBER', level=3)
    add_body(doc,
        'Menu engineering was formalised by Kasavana and Smith in 1982 for food menus. The hospitality industry adopted '
        'the principles for beverage programs and they have guided profitable bar design ever since.')
    add_body(doc,
        'The difference between a drink list and an engineered menu: a drink list tells guests what is available. '
        'An engineered menu guides them toward decisions that satisfy them and generate margin. Both look the same '
        'to the untrained eye. The revenue difference is measurable within a week.')
    add_body(doc,
        'Four tools drive menu engineering: pricing strategy, visual design, language, and item selection. Used '
        'together, they turn a static page into an active sales conversation.')
    add_body(doc,
        'Every item on a menu has two numbers that matter: what it costs to make, and how often it sells. Engineering '
        'connects both. An item that costs little to make but never sells is not a good menu item. An item that sells '
        'constantly but has a poor margin is not either.')
    add_body(doc,
        'A bar with 40 cocktails is not a better bar than one with 12. A shorter, well-engineered menu outperforms '
        'a long undisciplined list every time — in sales, in consistency, and in guest decision speed.')

    add_two_col_table(doc,
        ['TERM', 'DEFINITION'],
        [
            ['Menu Engineering', 'The discipline of designing a beverage list to guide guest choice and maximise profitability.'],
            ['Pour Cost %', 'The percentage of a drink\'s selling price that goes toward ingredient cost. Target: 18–22% for cocktails.'],
            ['Menu Real Estate', 'The visual space on a menu. Some positions attract the eye first. That placement is deliberate — or it should be.'],
            ['Anchor Pricing', 'A deliberate high-price item that makes mid-range items feel like better value by comparison.'],
            ['Decoy Item', 'An item priced to make other items more attractive. Often not expected to sell well — it exists to reframe perception.'],
        ]
    )

    add_box_block(doc, 'NOTE',
        'A bartender who understands menu engineering is a different kind of professional. You can execute a spec, '
        'but you can also evaluate a menu, identify a cost problem, and recommend a structural change. That is a '
        'consulting skill that commands a different conversation with an employer or client.')

    add_box_block(doc, 'PRO TIP',
        'The next time you sit with any bar menu, ignore the drinks for a moment and look at the page. Where does '
        'your eye go first? What item is in that position? That placement is either deliberate strategy or wasted '
        'real estate. You will now know the difference.', bg='FFF8DC')

    add_reflect(doc, [
        'What is the difference between a drink list and an engineered menu? Give one specific decision that separates them.',
        'A bar has 40 cocktails. Name two specific problems this creates — one for the guest and one for the business.',
    ])

    # ══════════════════════════════════════════════════
    # T2 — Your Costs
    # ══════════════════════════════════════════════════
    add_theory_header_sb(doc, 'T2', 'Your Costs — Pour Cost, Wine, and Labour')

    add_labelled_block(doc, 'CORE PRINCIPLE',
        'Every item on a menu has a true cost. Most bars only count the liquid. The real cost includes labour, waste, '
        'and execution time. Knowing the full number is what separates a profitable bar from a merely busy one.')

    add_heading(doc, 'WHAT YOU NEED TO REMEMBER', level=3)

    add_body(doc, 'POUR COST %', bold=True)
    add_body(doc,
        'Formula: (Ingredient cost ÷ Selling price) × 100\n'
        'Industry targets: Cocktails 18–22%  |  Beer 22–28%  |  Wine 28–35%\n'
        'Session standard: 2oz alcohol per cocktail, 1oz for a single spirit, 6oz wine glass.')
    add_body(doc,
        'Example: Classic Cosmopolitan ingredients cost $2.10. Selling price $15. '
        'Pour cost = (2.10 ÷ 15) × 100 = 14%. This is within target — and leaves room to improve margin with '
        'strategic modifications, which you will see in Build 04.')

    add_body(doc, 'WINE BY THE GLASS (BTG) LOGIC', bold=True)
    add_body(doc,
        'A 750ml bottle yields approximately four pours of 6oz, accounting for spillage and overpour. '
        'The rule: BTG price × 4 should equal or slightly exceed the bottle menu price. '
        'The glass price reflects the open-bottle risk. If BTG × 4 falls below the bottle price, guests '
        'are financially incentivised to order the bottle, which reduces transaction count and complicates service.')
    add_body(doc,
        'Example: Bottle at $45 → BTG at $12 × 4 = $48. Correct. '
        'Red flag: Bottle at $58 → BTG at $14 × 4 = $56. Under-priced by glass — the bar loses $2 of potential revenue every four glasses poured.')

    add_body(doc, 'LABOUR COST', bold=True)
    add_body(doc,
        'Industry target: 28–35% of total revenue. Labour is not just the bartender\'s wage. It includes build time, '
        'training time, error correction time, and ticket time. A cocktail with three spirit pours takes 40–60% longer '
        'to build than a single-spirit build. That time has a cost on every order.')
    add_body(doc,
        'A menu that is slow to execute costs money on every ticket. Complexity is justified only when the selling '
        'price reflects both the ingredient cost and the labour cost of execution. If it does not — the drink is '
        'selling at a loss that does not appear in the pour cost calculation.')

    add_two_col_table(doc,
        ['COST CATEGORY', 'TARGET', 'RED FLAG'],
        [
            ['Cocktail pour cost', '18–22%', 'Over 28%'],
            ['Wine pour cost (BTG)', '28–35%', 'Under 25% or over 40%'],
            ['Beer pour cost', '22–28%', 'Over 32%'],
            ['Labour cost', '28–35% of revenue', 'Over 38%'],
            ['Combined beverage cost', 'Under 30%', 'Over 35%'],
        ]
    )

    add_box_block(doc, 'NOTE',
        'A cocktail with a 12% pour cost is not automatically your best performer. If it takes 8 minutes to build, '
        'the labour cost may eliminate the margin advantage. Pour cost tells you what you spent on the bottle. '
        'It does not tell you what you spent on the bartender\'s time.')

    add_box_block(doc, 'PRO TIP',
        'Pick any cocktail from a menu. Estimate its ingredient cost. Then count how many distinct steps it takes '
        'to build. If the build takes longer than 3 minutes for a $15 drink, someone has not done the labour maths.', bg='FFF8DC')

    add_reflect(doc, [
        'A cocktail costs $2.80 in ingredients and sells for $16. Calculate the pour cost percentage. Is it within target?',
        'A wine bottle is $52 on the menu. The 6oz glass is $13. Using the BTG rule, is this correctly priced? What would you adjust?',
    ])

    # ══════════════════════════════════════════════════
    # T3 — The Menu Engineering Matrix
    # ══════════════════════════════════════════════════
    add_theory_header_sb(doc, 'T3', 'The Menu Engineering Matrix')

    add_labelled_block(doc, 'CORE PRINCIPLE',
        'Every item on your menu belongs in one of four categories. Knowing which is which tells you exactly what '
        'to do with it — promote, reprice, reposition, or remove.')

    add_heading(doc, 'WHAT YOU NEED TO REMEMBER', level=3)
    add_body(doc,
        'The matrix classifies every drink on two axes: profitability (high or low margin) and popularity '
        '(high or low sales volume). The combination produces four categories, each with a clear action.')
    add_body(doc,
        'STARS — High profit, high popularity. These are your best-performing drinks. Protect them, feature them, '
        'and do not change the spec without a strong reason. A Star is the result of good engineering. Do not undo it.')
    add_body(doc,
        'PLOWHORSES — Low profit, high popularity. Guests love them but the margin is weak. Options: raise the price, '
        'reduce the ingredient cost without changing the guest experience, or move them to a less prominent menu position '
        'to slow their velocity while you adjust the economics.')
    add_body(doc,
        'PUZZLES — High profit, low popularity. The drink makes money when it sells — it just does not sell enough. '
        'This is usually a positioning or visibility problem, not a recipe problem. Options: move it to a better '
        'menu position, improve the description, rename it, or ask whether staff are recommending it.')
    add_body(doc,
        'DOGS — Low profit, low popularity. These deserve the hardest review. Options: remove, redesign, or reprice. '
        'Keeping a Dog on the menu uses space, creates inventory, complicates training, and generates no return.')

    add_two_col_table(doc,
        ['CATEGORY', 'PROFIT', 'POPULARITY', 'ACTION'],
        [
            ['STAR', 'High', 'High', 'Promote. Feature. Do not change without reason.'],
            ['PLOWHORSE', 'Low', 'High', 'Reprice or reduce cost. Reposition if needed.'],
            ['PUZZLE', 'High', 'Low', 'Move on menu. Improve description. Train staff to recommend.'],
            ['DOG', 'Low', 'Low', 'Remove or redesign. Do not invest further without data.'],
        ]
    )

    add_heading(doc, 'Real examples from menus reviewed in this session:', level=3)
    add_two_col_table(doc,
        ['ITEM (CANNERY MENU)', 'ESTIMATED CATEGORY', 'REASON'],
        [
            ['Aperol Spritz — $16', 'Star candidate', 'Low ingredient cost. High trend appeal. Premium price.'],
            ['Old Fashioned — $13', 'Plowhorse', 'Popular classic. Likely under-priced vs. spirit cost.'],
            ['Drunken Starfish — $15', 'Puzzle', 'Strong name and spec. Buried mid-list. Not finding its audience.'],
            ['Fresh Berry Mojito — $13', 'Dog candidate', 'Lowest price. Highest labour. Fresh berry cost unpredictable.'],
            ['Espresso Martini — $16', 'Star', 'High perceived value. Simple spec. Strong trend performance.'],
        ]
    )

    add_box_block(doc, 'NOTE',
        'The matrix is a thinking tool, not a final verdict. You need sales data to confirm classification. '
        'But applying the matrix forces the right questions about every item before you have the data. '
        'Most bars never ask those questions at all.')

    add_box_block(doc, 'PRO TIP',
        'Apply the matrix to any menu before a meeting about that menu. You will have identified three or four '
        'specific decisions to discuss before anyone else in the room has opened a spreadsheet.', bg='FFF8DC')

    add_reflect(doc, [
        'A cocktail sells 40 times per week at a 14% pour cost. What matrix category is it most likely in? What is your recommended action?',
        'A manager wants to add 8 new cocktails to an already long menu. Using matrix logic, what two questions do you ask before agreeing?',
    ])

    # ══════════════════════════════════════════════════
    # T4 — Menu Design Principles
    # ══════════════════════════════════════════════════
    add_theory_header_sb(doc, 'T4', 'Menu Design Principles')

    add_labelled_block(doc, 'CORE PRINCIPLE',
        'Where an item is placed, what it is called, and how it is described are all decisions that affect sales. '
        'A menu is a visual and psychological tool. Every element either works for you or against you.')

    add_heading(doc, 'WHAT YOU NEED TO REMEMBER', level=3)

    add_body(doc, 'VISUAL REAL ESTATE', bold=True)
    add_body(doc,
        'Guests do not read a menu top to bottom. Eye-tracking research shows attention moves to the top-right of '
        'a panel first, then to the top-left, then to the center. This is called the Golden Triangle. '
        'Your highest-margin item belongs in the highest-value position. If your best cocktail is buried in the '
        'middle of a column, it is selling against its own location.')
    add_body(doc,
        'Boxes, borders, and bold text direct the eye. Use them for items you want sold — not for decoration. '
        'If everything is highlighted, nothing is. More than three visual callouts on a page cancel each other out.')

    add_body(doc, 'PRICING PSYCHOLOGY', bold=True)
    add_body(doc,
        'No dollar signs. The $ symbol activates price sensitivity. Listing prices as numerals only (15, not $15.00) '
        'reduces the friction of ordering. Anchor pricing: one deliberately high-priced item makes mid-range items '
        'feel like value. A $26 cocktail makes the $17 cocktail feel reasonable. Flat pricing bands — every cocktail '
        '$14 to $16 — give guests no guidance on what is worth ordering.')

    add_body(doc, 'LANGUAGE AND DESCRIPTION', bold=True)
    add_body(doc,
        'An ingredient list is not a description. "Vodka, lime, cranberry" tells a guest what is in the glass — '
        'not why they want it. Sensory language increases sales. Research by Wansink (2001) showed menus using '
        'evocative descriptions sold 27% more of those items than the same items described as ingredient lists. '
        'One word of texture or mood is enough: clean, bright, smoky, crushable, indulgent.')
    add_body(doc,
        'The description should match the brand voice. A Moroccan-themed bar should not describe its cocktails in '
        'generic language. Every word on a menu is either brand work or noise.')

    add_body(doc, 'SECTION NAMING AND ARCHITECTURE', bold=True)
    add_body(doc,
        'Maximum three cocktail sections. More than three creates decision paralysis. Section names must communicate '
        'instantly — guests scan, they do not study. "Cannery Cocktails" vs "Cannery Signature" tells the guest nothing. '
        '"CRAFT COCKTAILS" vs "CLASSICS" does the job in one word each.')

    add_two_col_table(doc,
        ['ELEMENT', 'STRONG PRACTICE', 'WEAK PRACTICE'],
        [
            ['Pricing', 'No dollar signs. Anchor item at premium.', 'Flat band. $ symbol. Prices buried in text.'],
            ['Language', 'Sensory, brand-matched description.', 'Ingredient list only.'],
            ['Visual design', 'Box or bold on 1–2 hero items.', 'Everything highlighted — or nothing.'],
            ['Section naming', '"CRAFT COCKTAILS" / "CLASSICS"', '"Cannery Cocktails" / "Cannery Signature"'],
            ['Item count', '10–14 cocktails per section maximum', '20+ items across 4+ overlapping sections'],
        ]
    )

    add_box_block(doc, 'NOTE',
        'The menu you hand a guest is doing sales work before you say a word. A well-designed menu creates a guest '
        'who arrives at the bar already knowing what they want. A poorly designed menu creates a guest who asks the '
        'server to explain everything — and that explanation costs time and table turns.')

    add_box_block(doc, 'PRO TIP',
        'Count the visual entry points on any menu. Every bold item, every box, every NEW! badge is a callout. '
        'If there are more than three, the page is fighting itself. The guest\'s eye does not know where to land.', bg='FFF8DC')

    add_reflect(doc, [
        'Why does removing dollar signs from a menu increase average spend? Name the psychological mechanism.',
        'A menu has six cocktail sections and 32 items total. Name two consequences for the guest and two for the bar.',
    ])

    # ══════════════════════════════════════════════════
    # T5 — Inventory, Par Levels, and Control
    # ══════════════════════════════════════════════════
    add_theory_header_sb(doc, 'T5', 'Inventory, Par Levels, and Control')

    add_labelled_block(doc, 'CORE PRINCIPLE',
        'A well-engineered menu only works if the inventory behind it is controlled. Par levels, FIFO, and variance '
        'tracking are not administrative tasks — they are the operating system of a profitable bar.')

    add_heading(doc, 'WHAT YOU NEED TO REMEMBER', level=3)

    add_body(doc, 'PAR LEVEL', bold=True)
    add_body(doc,
        'A par level is the minimum amount of a product you need on hand to last until the next delivery, plus a '
        'safety buffer. It prevents running out mid-service.')
    add_body(doc,
        'Formula: (Daily usage × days between orders) + safety stock = par level\n'
        'Example: You sell 4 bottles of vodka per day and order every 3 days. Safety stock: 2 bottles. '
        'Par = (4 × 3) + 2 = 14 bottles minimum on hand at all times.')
    add_body(doc,
        'Par levels are directly connected to your menu. High-selling cocktails drive high par on their spirit. '
        'A cocktail that has been on the menu for three months and rarely orders is a low-par item — and a '
        'signal that it belongs in the Dog category of your matrix.')

    add_body(doc, 'MENU ENGINEERING AND PURCHASING', bold=True)
    add_body(doc,
        'Every cocktail on your menu requires at least one bottle on the shelf. A menu with 40 cocktails and '
        '60 different bottles is 60 par-level commitments. A shorter, engineered menu with 12 cocktails and '
        '20 shared bottles is dramatically easier to manage, count, and control.')
    add_body(doc,
        'Dead inventory is money on the shelf. A bottle of specialty liqueur purchased for one cocktail that '
        'never sells sits at cost price until it is used, discarded, or written off. Every Dog on your menu '
        'is a potential dead inventory item.')

    add_body(doc, 'FIFO — FIRST IN, FIRST OUT', bold=True)
    add_body(doc,
        'New stock goes behind old stock. The oldest product is always poured first. Critical for: fresh citrus '
        'juice (24-hour maximum), house-made syrups, open wine bottles (72-hour maximum refrigerated), '
        'and any perishable prep. FIFO is a quality standard as much as a cost control tool.')

    add_body(doc, 'VARIANCE TRACKING', bold=True)
    add_body(doc,
        'Theoretical inventory is what the POS says you should have poured. Actual inventory is what the '
        'physical count shows. The gap between them is variance. Variance comes from: over-pouring, spillage, '
        'theft, incorrect recipe execution, or unrecorded complimentary pours. A standardised menu with '
        'consistent recipes reduces variance because every build is the same, every time.')

    add_two_col_table(doc,
        ['TERM', 'DEFINITION'],
        [
            ['Par Level', 'Minimum stock on hand to last until next delivery plus safety buffer.'],
            ['SKU', 'Stock Keeping Unit. One unique product. Each bottle on the back bar is one SKU.'],
            ['FIFO', 'First In, First Out. Oldest stock poured first, always.'],
            ['Dead Inventory', 'Product that sits unsold. Costs money without generating revenue.'],
            ['Variance', 'Gap between what the POS says was sold and what the physical count confirms.'],
            ['Theoretical Inventory', 'What the system calculates you should have based on recorded sales.'],
        ]
    )

    add_box_block(doc, 'NOTE',
        'A cocktail that performs poorly does not just fail on the matrix. It creates dead inventory, complicates '
        'par calculations, slows physical counts, and ties up cash in product that is not generating return. '
        'A Dog on the menu is a problem in multiple departments simultaneously.')

    add_box_block(doc, 'PRO TIP',
        'Before redesigning any menu, count the SKUs the current cocktail list requires. Then count how many '
        'appear in more than one cocktail. The ones used in only one drink are your inventory risk items. '
        'Those are the first candidates for removal or substitution.', bg='FFF8DC')

    add_reflect(doc, [
        'A bar orders vodka every 3 days and sells 3 bottles per day. They want 1 bottle safety stock. What is the par level? What happens if they fall below it mid-service?',
        'A cocktail has been on the menu for 3 months and orders fewer than 2 per week. Name three ways this underperformance affects the business beyond its matrix classification.',
    ])

    # ══════════════════════════════════════════════════
    # T6 — Cocktail Engineering in Practice
    # ══════════════════════════════════════════════════
    add_theory_header_sb(doc, 'T6', 'Cocktail Engineering in Practice')

    add_labelled_block(doc, 'CORE PRINCIPLE',
        'The difference between a cocktail list and a profitable menu is execution. Three deliberate modifications '
        'to a single drink can justify a price increase of 25–30% while adding under $0.50 in ingredient cost. '
        'That is cocktail engineering — and it is why people pay professionals to build menus, not assemble lists.')

    add_heading(doc, 'WHAT YOU NEED TO REMEMBER', level=3)

    add_body(doc, 'MENU COMPOSITION — CATEGORY COVERAGE', bold=True)
    add_body(doc,
        'A well-built cocktail menu covers all major drinking occasions. If a category is missing, the bar loses '
        'an order every time a guest at that table wants that style of drink. This is not about variety for its '
        'own sake — it is about not losing a sale.')

    add_two_col_table(doc,
        ['CATEGORY', 'EXAMPLE BUILDS', 'GUEST IT SERVES'],
        [
            ['Sour (shaken, citrus-forward)', 'Cosmopolitan, Daiquiri, Margarita', 'Refreshing, bright drinkers'],
            ['Spirit-Forward (stirred)', 'Negroni, Old Fashioned, Manhattan', 'Serious drinkers, digestif crowd'],
            ['Sparkling / Light', 'Aperol Spritz, French 75, Prosecco Spritz', 'Easy-drinking, social occasions'],
            ['Sweet / Indulgent', 'Espresso Martini, Mudslide, Chocolate Martini', 'After-dinner, dessert drinkers'],
            ['Long / Refreshing', 'Moscow Mule, Mojito, Tom Collins', 'Casual, warm-weather, session drinkers'],
            ['Zero-Proof', 'N/A Sour, Shrub Mocktail, Sparkling Juice', 'Non-drinking guests — every table'],
        ]
    )

    add_body(doc,
        'A menu missing the zero-proof category loses every table with one non-drinking guest. A menu without '
        'a spirit-forward option loses the guest who wants something serious. Missing sweet or sparkling loses '
        'the after-dinner and celebration orders. Category coverage is not about pleasing everyone — it is about '
        'not writing off an entire type of sale.')

    add_body(doc, 'THE MODIFICATION PRINCIPLE — THREE TOOLS, ONE LESSON', bold=True)
    add_body(doc,
        'Three modifications exist that cost under $0.50 each and justify a price increase of $3–5 per drink. '
        'Used together, they transform a standard cocktail into a menu item guests photograph, remember, and '
        'describe to other people. That word-of-mouth is not accidental — it is designed.')

    add_two_col_table(doc,
        ['MODIFICATION', 'COST', 'WHAT IT ADDS', 'PRICE JUSTIFICATION'],
        [
            ['Dehydrated citrus wheel (lime / lemon / orange)', '~$0.10–0.15', 'Visual premium. Shelf life: weeks. Made in advance.', '+$1–2'],
            ['Aromatic spray or expressed peel (bitters / citrus oil)', '~$0.05', 'Aroma before first sip. Guest experiences the drink before tasting it.', '+$1–2'],
            ['Spirit or syrup infusion (24–48h cold process)', '~$0.10–0.20', 'Depth and uniqueness. A flavour no other bar can replicate without your recipe.', '+$2–3'],
        ]
    )

    add_body(doc,
        'Any one of these elevates a $15 drink to $17–18. All three together create a $19–21 cocktail. '
        'The ingredient cost increase is under $0.50. The margin improves. The guest experience is categorically '
        'different. That is the difference between a list and a menu. That is why clients pay for menu design.')

    add_box_block(doc, 'NOTE',
        'Every modification must earn its place. A garnish that does not contribute aroma, flavour, or texture '
        'is visual cost — not guest value. A dehydrated wheel that has no connection to the drink\'s flavour '
        'profile is decoration. A dehydrated lime wheel on a citrus-forward cocktail is architecture.')

    add_box_block(doc, 'PRO TIP',
        'Look at any $22 cocktail on any craft bar menu and break it down. Almost always: one infused or '
        'house-made component, one aromatic garnish, one visual element. The spirit itself may cost the same '
        'as a $14 drink. The modifications are the entire price difference.', bg='FFF8DC')

    add_reflect(doc, [
        'A cocktail menu has no zero-proof section and no sparkling category. Name two guest types who leave without ordering and explain what this means for that table\'s total spend.',
        'You are asked to elevate a $15 Cosmopolitan to a $19 cocktail without adding more than $0.50 in ingredient cost. Describe two modifications you would make and what each one contributes to the guest experience.',
    ])

    # ══════════════════════════════════════════════════
    # PRACTICAL SECTION
    # ══════════════════════════════════════════════════
    add_divider(doc)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('PRACTICAL')
    r.bold = True; r.font.size = Pt(16)
    r.font.color.rgb = RGBColor.from_string('1F3864')
    add_divider(doc)

    add_body(doc,
        'Four builds. Each one is connected to the theory. The first two demonstrate how a single addition '
        'creates a new product and justifies a higher price. The third shows you a real menu cocktail — costed '
        'and classified. The fourth shows you what happens when you engineer it.')

    # ── Build 01: Black Russian ──────────────────────────────────────────
    add_build_header(doc, 1, 'Black Russian')
    add_body(doc,
        'The Black Russian is the baseline of today\'s practical. Two ingredients, built in a glass, no shake, '
        'no strain. Its value is in what it teaches: a simple build with a well-controlled pour cost is not a '
        'lesser cocktail — it is the most efficient one on your menu.')

    add_body(doc, 'BEFORE YOU TOUCH ANYTHING', bold=True)
    add_body(doc,
        'State the spec from memory. Name, both ingredients in oz and ml, build order, glass, ice, garnish. '
        'Then calculate the pour cost before you pour anything.')

    add_two_col_table(doc,
        ['INGREDIENT', 'AMOUNT', 'NOTES'],
        [
            ['Vodka (house)', '2oz / 60ml', 'Ketel One, Smirnoff, or house selection'],
            ['Coffee Liqueur', '1oz / 30ml', 'Kahlúa or Bolivar'],
        ]
    )

    add_body(doc, 'Glass: Rocks  |  Ice: Large cube or cubed ice  |  Method: Build  |  Garnish: None')
    add_body(doc,
        'Build order: Add ice to rocks glass. Measure and pour vodka. Add coffee liqueur. '
        'Stir 3–4 slow rotations to combine. Straw-taste. Present.')

    add_two_col_table(doc,
        ['COST ELEMENT', 'AMOUNT', 'CALCULATION'],
        [
            ['Vodka (house) 2oz', '$1.00/oz', '$2.00'],
            ['Coffee Liqueur 1oz', '$0.80/oz', '$0.80'],
            ['Total ingredient cost', '', '$2.80'],
            ['Selling price', '', '$14'],
            ['Pour cost %', '(2.80 ÷ 14) × 100', '20%  ✓ On target'],
        ]
    )

    # ── Build 02: White Russian ──────────────────────────────────────────
    add_build_header(doc, 2, 'White Russian')
    add_body(doc,
        'The White Russian is a menu engineering lesson in one pour. Same base as the Black Russian. '
        'One additional ingredient. A measurably different guest experience. A justified price increase. '
        'This is what a menu extension looks like — not a new concept, a deliberate variation.')

    add_body(doc, 'BEFORE YOU TOUCH ANYTHING', bold=True)
    add_body(doc,
        'State how this spec differs from Build 01. Name the one change and explain what it contributes '
        'to the drink before you pour anything.')

    add_two_col_table(doc,
        ['INGREDIENT', 'AMOUNT', 'NOTES'],
        [
            ['Vodka (house)', '2oz / 60ml', 'Same as Build 01'],
            ['Coffee Liqueur', '1oz / 30ml', 'Same as Build 01'],
            ['Heavy Cream', '1oz / 30ml', 'Float — do not stir after'],
        ]
    )

    add_body(doc, 'Glass: Rocks  |  Ice: Large cube or cubed ice  |  Method: Build + float  |  Garnish: None')
    add_body(doc,
        'Build order: Build exactly as the Black Russian. Then float the cream slowly over the back of a bar spoon. '
        'Do not stir. The cream layers on top. The guest drinks through the cream into the coffee-and-vodka base.')

    add_two_col_table(doc,
        ['COST ELEMENT', 'AMOUNT', 'CALCULATION'],
        [
            ['Vodka + Coffee Liqueur', '', '$2.80 (same as Build 01)'],
            ['Heavy Cream 1oz', '$0.25/oz', '$0.25'],
            ['Total ingredient cost', '', '$3.05'],
            ['Selling price', '', '$15'],
            ['Pour cost %', '(3.05 ÷ 15) × 100', '20.3%  ✓ On target'],
            ['Additional cost vs Build 01', '', '$0.25'],
            ['Additional price vs Build 01', '', '$1.00'],
            ['Additional profit per drink', '', '$0.75'],
        ]
    )

    add_box_block(doc, 'THE LESSON',
        '$0.25 more in cost. $1.00 more in price. $0.75 more in profit per drink. The cream float changes the '
        'texture, the visual, and the guest\'s experience of the drink. That change has a calculable value.')

    # ── Build 03: Classic Cosmopolitan ──────────────────────────────────
    add_build_header(doc, 3, 'Classic Cosmopolitan')
    add_body(doc,
        'This is the Cosmopolitan as listed on the Cannery Kitchen menu at $15. You will build it to spec, '
        'cost it, and classify it on the matrix. Then you will identify what is holding it back on the menu — '
        'not the recipe, but the engineering around it.')

    add_body(doc, 'BEFORE YOU TOUCH ANYTHING', bold=True)
    add_body(doc,
        'State the spec. Then state where you think this drink sits on the matrix and why — before you '
        'build it, before you see the cost calculation.')

    add_two_col_table(doc,
        ['INGREDIENT', 'AMOUNT', 'NOTES'],
        [
            ['Vodka (house)', '1.5oz / 45ml', 'Citrus-neutral. Do not use flavoured vodka here.'],
            ['Triple Sec', '0.5oz / 15ml', 'Cointreau preferred. Adjust price accordingly if used.'],
            ['Fresh Lime Juice', '0.5oz / 15ml', 'Squeeze to order. Not bottled.'],
            ['Cranberry Juice', '0.5oz / 15ml', 'Just enough for colour. Not a dominant flavour.'],
        ]
    )

    add_body(doc, 'Glass: Chilled coupe  |  Ice: None (strained)  |  Method: Shake, double strain  |  Garnish: Expressed lemon twist')
    add_body(doc,
        'Build order: Add all ingredients to a shaker with ice. Shake hard for 12–15 seconds. '
        'Double-strain into a pre-chilled coupe. Express a lemon twist over the glass surface and place on the rim.')

    add_two_col_table(doc,
        ['COST ELEMENT', 'AMOUNT', 'CALCULATION'],
        [
            ['Vodka 1.5oz', '$1.00/oz', '$1.50'],
            ['Triple Sec 0.5oz', '$0.80/oz', '$0.40'],
            ['Fresh Lime Juice 0.5oz', 'est.', '$0.15'],
            ['Cranberry Juice 0.5oz', 'est.', '$0.05'],
            ['Total ingredient cost', '', '$2.10'],
            ['Selling price (Cannery)', '', '$15'],
            ['Pour cost %', '(2.10 ÷ 15) × 100', '14%  — Well below target'],
        ]
    )

    add_box_block(doc, 'MATRIX CLASSIFICATION',
        'Pour cost 14% — excellent. But this drink is listed in the "Cannery Signature" section between generic classics '
        'with no visual callout and a generic ingredient-list description. High margin, low visibility = PUZZLE. '
        'The recipe is not the problem. The engineering around it is.')

    # ── Build 04: Elevated Cosmopolitan ─────────────────────────────────
    add_build_header(doc, 4, 'Elevated Cosmopolitan')
    add_body(doc,
        'The same drink. Three modifications. A different menu position, a different description, and a selling '
        'price of $19. The ingredient cost increase is $0.42. The profit increase per drink is $3.58. '
        'This is why menus are designed — not assembled.')

    add_body(doc, 'PRE-PREPARATION NOTE', bold=True)
    add_body(doc,
        'The lemon peel infusion is prepared by the instructor 24–48 hours before the session. Pack a sealed '
        'glass jar with fresh lemon peels. Cover completely with vodka. Refrigerate. Strain through fine mesh '
        'before class. The lemon peels are then squeezed for their juice, which becomes the citrus component '
        'of the build. Nothing is wasted.')

    add_two_col_table(doc,
        ['INGREDIENT', 'AMOUNT', 'NOTES'],
        [
            ['Lemon Peel Infused Vodka', '1.5oz / 45ml', 'Prepared 24–48h prior. Same base cost.'],
            ['Cointreau', '0.5oz / 15ml', 'Do not substitute Triple Sec for this build.'],
            ['Fresh Lemon Juice', '0.5oz / 15ml', 'Squeezed from the infusion lemons. Byproduct — near zero cost.'],
            ['White Cranberry Juice', '0.5oz / 15ml', 'Lighter colour. More refined visual than standard cranberry.'],
            ['Orange Bitters', '2 dashes or spray', 'Dash into shaker OR atomize 2 sprays inside the chilled coupe before straining.'],
            ['Dehydrated Citrus Wheel', '1 piece', 'Lime, lemon, or orange. Prepared in advance. Lasts weeks.'],
            ['Expressed Orange Peel', '1 strip', 'Express oils over finished drink. Do not drop in — place on rim.'],
        ]
    )

    add_body(doc, 'Glass: Chilled coupe  |  Ice: None (strained)  |  Method: Shake, double strain  |  Garnish: Dehydrated wheel + expressed orange peel')
    add_body(doc, 'Build order:', bold=True)

    steps = [
        'Pre-chill the coupe.',
        'If using the bitters spray: 2 pumps inside the chilled coupe before straining. Set aside.',
        'Add infused vodka, Cointreau, lemon juice, and white cranberry to a shaker with ice.',
        'Shake hard for 12–15 seconds.',
        'Double-strain into the prepared coupe.',
        'Express orange peel over the surface of the drink. Hold the peel skin-side down and pinch firmly over the glass. The oils mist across the surface. Place the peel on the rim.',
        'Position the dehydrated citrus wheel on the rim opposite the peel. Straw-taste. Present.',
    ]
    for i, step in enumerate(steps, 1):
        p = doc.add_paragraph(f'{i}.  {step}', style='List Number')
        p.runs[0].font.size = Pt(10)
        p.paragraph_format.space_after = Pt(3)

    add_two_col_table(doc,
        ['VERSION', 'INGREDIENT COST', 'SELLING PRICE', 'POUR COST %', 'ADDITIONAL PROFIT'],
        [
            ['Classic Cosmopolitan', '$2.10', '$15', '14.0%', 'Baseline'],
            ['+ Infused Vodka only', '$2.10', '$17', '12.4%', '+$1.90 / drink'],
            ['+ Cointreau upgrade', '$2.45', '$18', '13.6%', '+$2.55 / drink'],
            ['All 3 modifications', '$2.52', '$19', '13.3%', '+$3.58 / drink'],
        ]
    )

    add_box_block(doc, 'THE LESSON',
        'Cost increase: $0.42. Price increase: $4.00. Additional profit per drink: $3.58. Pour cost improved from '
        '14% to 13.3%. Every modification contributed to the guest experience — aroma, visual, flavour depth. '
        'None of them were decoration. That is cocktail engineering.')

    # ── Menu Analysis Exercise ───────────────────────────────────────────
    add_divider(doc)
    add_heading(doc, 'Menu Analysis Exercise', level=2)
    add_body(doc,
        'Choose one cocktail from either menu reviewed in this session — the Cannery Kitchen or The Sultan\'s Tent. '
        'Complete all four steps below.')

    add_two_col_table(doc,
        ['STEP', 'YOUR WORK'],
        [
            ['1. Name the cocktail and its menu price.', ''],
            ['2. Estimate the ingredient cost (use session standard: 2oz alcohol, 1oz single spirit).', ''],
            ['3. Calculate the pour cost %. Classify it on the matrix with a reason.', ''],
            ['4. Identify one modification under $0.50 that justifies a $2–3 price increase. Describe what it contributes.', ''],
            ['5. Write a new description for the drink — maximum 15 words, sensory language, brand-matched.', ''],
        ]
    )

    # ── Knowledge Test ───────────────────────────────────────────────────
    add_divider(doc)
    add_heading(doc, 'Knowledge Test', level=2)

    test_qs = [
        'Write the pour cost formula from memory.',
        'A wine bottle is $48 on the menu. The 6oz glass is $11. Is this correctly priced? What would you change?',
        'Name the four cells of the menu engineering matrix and the primary action for each.',
        'What does FIFO stand for? Where does it matter most in a bar operation?',
        'The Elevated Cosmopolitan costs $0.42 more to make than the Classic and sells for $4.00 more. What is the additional profit per drink?',
    ]
    for i, q in enumerate(test_qs, 1):
        p = doc.add_paragraph(f'{i}.  {q}', style='List Number')
        p.runs[0].font.size = Pt(10)
        p.paragraph_format.space_after = Pt(14)

    # footer
    add_divider(doc)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('bar-skills.com | info@bar-skills.com')
    r.font.size = Pt(9)
    r.font.color.rgb = RGBColor.from_string('7F7F7F')

    doc.save('/home/user/OI1111/BarSkills_S09_SB.docx')
    print('SB saved.')


# ═══════════════════════════════════════════════════════════════════════════
#  TEACHER'S MANUAL
# ═══════════════════════════════════════════════════════════════════════════

def build_tm():
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.9)
        section.bottom_margin = Inches(0.9)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # ── COVER ──────────────────────────────────────────────────────────────
    for text, size, bold in [
        ('Bar-Skills', 18, True),
        ('Craft Bartender Course', 12, False),
        ('SESSION 09', 22, True),
        ('Menu Engineering', 16, True),
        ('Cost, Design, and the Profitable Bar', 12, False),
        ("TEACHER'S MANUAL", 12, True),
        ('Oleks Iurchenko | bar-skills.com | info@bar-skills.com', 9, False),
    ]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(text)
        r.bold = bold
        r.font.size = Pt(size)
        if text.startswith('SESSION') or text == 'Bar-Skills':
            r.font.color.rgb = RGBColor.from_string('1F3864')
        if text == 'Oleks Iurchenko | bar-skills.com | info@bar-skills.com':
            r.font.color.rgb = RGBColor.from_string('7F7F7F')

    add_divider(doc)

    # ── SESSION CONTROL ────────────────────────────────────────────────────
    add_heading(doc, 'Session Control', level=2)

    for label, text in [
        ('OBJECTIVE',
         'The student understands menu engineering as a commercial discipline, can calculate pour cost for '
         'cocktails and wine, applies the four-cell matrix to real menu items, and executes four builds — '
         'including a two-stage Cosmopolitan that demonstrates how deliberate modifications improve both guest '
         'experience and bar margin simultaneously.'),
        ('OUTCOME',
         'Student can state the pour cost formula and calculate it for any build. Student can classify a drink '
         'on the matrix with a rationale. Student can execute the Black Russian, White Russian, Classic Cosmopolitan, '
         'and Elevated Cosmopolitan to the same standard as all previous session builds. Student can name at least '
         'two modifications that justify a higher price without proportional cost increase.'),
        ('FORMAT', '30 min Theory + 60 min Practical'),
        ('PRE-VIDEO', '[To be populated by Oleks]'),
        ('POST-VIDEO', '[To be populated by Oleks]'),
    ]:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        lr = p.add_run(label + '  ')
        lr.bold = True; lr.font.size = Pt(10)
        lr.font.color.rgb = RGBColor.from_string('1F3864')
        tr = p.add_run(text)
        tr.font.size = Pt(10)

    doc.add_paragraph()

    # timing table
    tbl = doc.add_table(rows=9, cols=4)
    tbl.style = 'Table Grid'
    headers = [('THEORY (30 MIN)', '1F3864'), ('TIME', '1F3864'), ('PRACTICAL (60 MIN)', '1F3864'), ('TIME', '1F3864')]
    for i, (h, col) in enumerate(headers):
        cell = tbl.rows[0].cells[i]
        set_cell_bg(cell, '1F3864')
        p = cell.paragraphs[0]
        r = p.add_run(h)
        r.bold = True; r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    theory_rows = [
        ('T1: What Is Menu Engineering?', '5 min', 'Station Orientation', '5 min'),
        ('T2: Your Costs — Pour Cost, Wine, and Labour', '6 min', 'Build 01: Black Russian', '10 min'),
        ('T3: The Menu Engineering Matrix', '5 min', 'Build 02: White Russian', '8 min'),
        ('T4: Menu Design Principles', '5 min', 'Build 03: Classic Cosmopolitan', '12 min'),
        ('T5: Inventory, Par Levels, and Control', '5 min', 'Build 04: Elevated Cosmopolitan', '18 min'),
        ('T6: Cocktail Engineering in Practice', '4 min', 'Menu Analysis Exercise', '5 min'),
        ('', '', 'Knowledge Test', '2 min'),
        ('TOTAL', '30 min', 'TOTAL', '60 min'),
    ]
    for ri, row in enumerate(theory_rows):
        for ci, val in enumerate(row):
            cell = tbl.rows[ri+1].cells[ci]
            if ri % 2 == 0:
                set_cell_bg(cell, 'EBF0F7')
            if ri == 7:
                set_cell_bg(cell, 'D6E4F0')
            p = cell.paragraphs[0]
            r = p.add_run(val)
            r.font.size = Pt(9)
            if ri == 7:
                r.bold = True

    doc.add_paragraph()
    add_body(doc,
        'If theory runs long, compress T5 and T6 — carry inventory and modification detail into the practical '
        'debrief between builds. The builds teach faster than talking. Build 04 is the session\'s main argument. '
        'Do not cut it short.')

    add_divider(doc)

    # ══════════════════════════════════════════════════
    # TM T1
    # ══════════════════════════════════════════════════
    add_theory_header_tm(doc, 5, 'T1', 'What Is Menu Engineering?')

    add_labelled_block(doc, 'CORE PRINCIPLE',
        'Menu engineering is the discipline of designing a drink list to maximise both guest satisfaction and profitability. '
        'It is not decoration — it is strategy.')

    add_heading(doc, 'WHAT TO SAY', level=3)
    add_body(doc,
        'Menu engineering was formalised by Kasavana and Smith in 1982 for food menus. The hospitality industry '
        'adopted the principles for beverage programs and they have shaped every well-run bar program since. '
        'Most bartenders have never heard of it. After today, you will not be most bartenders.')
    add_body(doc,
        'The distinction to establish first: a drink list tells guests what is available. An engineered menu guides '
        'them toward decisions that satisfy them and generate margin. Both look the same to an untrained eye. '
        'The revenue difference is measurable within a week of implementing one over the other.')
    add_body(doc,
        'Four tools: pricing strategy, visual design, language, and item selection. None of them are about the recipe. '
        'Menu engineering assumes you already know how to make the drinks — it decides how to sell them.')
    add_body(doc,
        'Shorter menus outperform longer ones. A bar with 40 cocktails is not a better bar than one with 12. '
        'It is a bar with 40 training commitments, 40 inventory lines, and a guest who cannot make a decision.')

    add_heading(doc, 'WHAT TO SHOW', level=3)
    add_body(doc,
        'Place the two menus from this session on the table — the Cannery Kitchen and The Sultan\'s Tent. '
        'Ask: which of these is a drink list and which is an engineered menu? Let the student answer before '
        'you tell them. The Sultan\'s Tent has a brand story in every cocktail name — that is engineering. '
        'The Cannery has two sections with nearly identical names — that is not.')
    add_body(doc,
        'Then ask: without tasting anything, which menu makes you want to order a cocktail? Why? '
        'The answer tells you what visual and language engineering does — before a single drink is made.')

    add_two_col_table(doc,
        ['TERM', 'DEFINITION'],
        [
            ['Menu Engineering', 'Designing a beverage list to guide guest choice and maximise profitability.'],
            ['Pour Cost %', 'Ingredient cost as a percentage of selling price. Target 18–22% for cocktails.'],
            ['Menu Real Estate', 'Visual space on a menu. Some positions attract the eye first. That is either deliberate or wasted.'],
            ['Anchor Pricing', 'A premium-priced item that makes mid-range items feel like value by comparison.'],
            ['Decoy Item', 'Priced to make other items more attractive. May not be expected to sell — it reframes perception.'],
        ]
    )

    add_box_block(doc, 'WATCH OUT',
        'Students who bartend may push back: "I just make the drinks." Acknowledge it. Then ask: have you ever '
        'worked on a slow night and wondered why no one was ordering? Or watched a new menu launch and suddenly '
        'certain drinks started selling without anyone changing the recipe? That is engineering — or the absence of it.')

    add_quote(doc,
        'A bartender who understands menu engineering is a different kind of professional. You can execute a spec, '
        'but you can also evaluate a menu, identify a cost problem, and recommend a change. That is a consulting '
        'skill that commands a different conversation with every employer you will ever have.')

    add_body(doc, 'DISCUSSION PROMPT', bold=True)
    add_body(doc,
        'Look at the Cannery menu. Find the two section names "Cannery Cocktails" and "Cannery Signature." '
        'What is the difference between them? Without reading the descriptions — could a guest tell? '
        'What does that confusion cost the bar?')

    add_transition(doc,
        'Now that we understand what menu engineering is, we need to understand the numbers it works with. '
        'Everything in menu engineering starts with knowing what a drink actually costs.')

    # ══════════════════════════════════════════════════
    # TM T2
    # ══════════════════════════════════════════════════
    add_theory_header_tm(doc, 6, 'T2', 'Your Costs — Pour Cost, Wine, and Labour')

    add_labelled_block(doc, 'CORE PRINCIPLE',
        'Every item on a menu has a true cost. Most bars only count the liquid. The real cost includes labour, '
        'waste, and execution time. Knowing the full number separates a profitable bar from a merely busy one.')

    add_heading(doc, 'WHAT TO SAY', level=3)
    add_body(doc,
        'Pour cost is the foundational calculation. (Ingredient cost ÷ Selling price) × 100. Session standard: '
        '2oz alcohol per cocktail, 1oz for a single spirit pour, 6oz wine glass. These are our working assumptions '
        'for every calculation in this session and every exercise.')
    add_body(doc,
        'Industry targets: cocktails 18–22%, beer 22–28%, wine 28–35%. These are not rules — they are benchmarks. '
        'A 14% pour cost looks great until you factor in a 7-minute build. A 22% pour cost that takes 90 seconds '
        'to execute may be more profitable in real terms.')
    add_body(doc,
        'Wine BTG logic: 750ml = approximately 4 pours of 6oz. BTG price × 4 should equal or slightly exceed the '
        'bottle price. The bar takes on open-bottle risk. If BTG × 4 falls below bottle price, guests are '
        'incentivised to order the bottle — which reduces transaction count.')
    add_body(doc,
        'Labour cost target: 28–35% of total revenue. A 3-spirit cocktail at $18 may look strong on pour cost '
        'but takes 40–60% longer to build than a single-spirit build at $15. Add training time, error rate, and '
        'ticket time — the $15 drink may be more profitable after labour.')

    add_heading(doc, 'WHAT TO SHOW', level=3)
    add_body(doc,
        'Work through the Classic Cosmopolitan calculation live. Write it out: $2.10 cost ÷ $15 price × 100 = 14%. '
        'Ask: is this good? Then ask: what if it takes 6 minutes to make because the bar has to squeeze fresh juice '
        'to order during service? What does the pour cost look like then? (Answer: still 14%, but the labour cost '
        'on 6 minutes of bartender time may cost the bar more than the 4 points of pour cost they saved.)')
    add_body(doc,
        'Show the Sultan\'s Tent Moroccan Mojito: $16 selling price, two spirits (Bacardi White + Bacardi Black). '
        'At 1oz each: two spirit pours at ~$0.80–$1.00 per oz = $1.60–$2.00 in spirits alone plus juice, syrup, '
        'rose water. Estimate total: $3.50–$4.50. Pour cost: 22–28%. Borderline at best. At a busy bar this drink '
        'is a liability — and it has three garnishes.')

    add_two_col_table(doc,
        ['COST CATEGORY', 'TARGET', 'RED FLAG'],
        [
            ['Cocktail pour cost', '18–22%', 'Over 28%'],
            ['Wine pour cost (BTG)', '28–35%', 'Under 25% or over 40%'],
            ['Beer pour cost', '22–28%', 'Over 32%'],
            ['Labour cost', '28–35% of revenue', 'Over 38%'],
            ['Combined beverage cost', 'Under 30%', 'Over 35%'],
        ]
    )

    add_box_block(doc, 'WATCH OUT',
        'Students may ask: "How do I know what the bar pays for a bottle?" Explain: you use the wholesale cost, not '
        'the retail price. In training, use estimates based on typical wholesale — roughly 30–40% of retail. '
        'For this session, the costs provided in the build cards are working estimates. The formula and the thinking '
        'are what matters, not the exact number.')

    add_quote(doc,
        'Pour cost tells you what you spent on the bottle. It does not tell you what you spent on the bartender\'s '
        'time. A cocktail that takes 8 minutes to build and costs 13% in pour cost may be losing money on every order.')

    add_body(doc, 'DISCUSSION PROMPT', bold=True)
    add_body(doc,
        'The Sultan\'s Mule contains Arak, Soho liqueur, and Vodka — three spirits — plus ginger beer, '
        'mango nectar, and lime. It sells for $18. Walk through the pour cost estimate together. '
        'Then ask: what would need to change to make this drink profitable at scale?')

    add_transition(doc,
        'We have the cost calculation tools. Now we need a framework that tells us what to do with every item on '
        'the menu once we know its cost and how often it sells.')

    # ══════════════════════════════════════════════════
    # TM T3
    # ══════════════════════════════════════════════════
    add_theory_header_tm(doc, 5, 'T3', 'The Menu Engineering Matrix')

    add_labelled_block(doc, 'CORE PRINCIPLE',
        'Every item on a menu belongs in one of four categories. Knowing which is which tells you exactly what '
        'to do with it — promote, reprice, reposition, or remove.')

    add_heading(doc, 'WHAT TO SAY', level=3)
    add_body(doc,
        'The matrix has two axes: profitability (high or low) and popularity (high or low). The combination '
        'produces four categories. Each one has a specific action — and "do nothing" is not one of them.')
    add_body(doc,
        'Stars: High profit, high popularity. Protect them. Feature them. Do not change the spec because you '
        'are bored of making it. The guest is not bored of ordering it.')
    add_body(doc,
        'Plowhorses: Low profit, high popularity. The most dangerous category. These drinks are eating your '
        'margin while appearing to perform. Options: raise the price by $1–2, reduce one ingredient cost without '
        'changing the experience, or move them to a less prominent position to slow their velocity.')
    add_body(doc,
        'Puzzles: High profit, low popularity. The drink makes money when it sells — it just is not selling. '
        'Almost always a positioning or description problem. Move it, rename it, ask staff to recommend it.')
    add_body(doc,
        'Dogs: Low profit, low popularity. Remove or redesign. Do not let sentiment keep a Dog on a menu. '
        'Every Dog is costing the bar in inventory, training, and space.')

    add_heading(doc, 'WHAT TO SHOW', level=3)
    add_body(doc,
        'Put the Cannery menu in front of the student. Walk through 4–5 items together using cost estimates and '
        'visibility on the page. Classify each one out loud. The Fresh Berry Mojito at $13 with fresh berries is '
        'a Dog candidate — lowest price, highest ingredient cost, highest labour. The Aperol Spritz at $16 with '
        'low ingredient cost is a Star candidate. Make the student argue for each classification.')

    add_two_col_table(doc,
        ['CATEGORY', 'PROFIT', 'POPULARITY', 'ACTION'],
        [
            ['STAR', 'High', 'High', 'Promote. Feature. Do not change without reason.'],
            ['PLOWHORSE', 'Low', 'High', 'Reprice or reduce cost. Consider repositioning.'],
            ['PUZZLE', 'High', 'Low', 'Move on menu. Better description. Train staff to recommend.'],
            ['DOG', 'Low', 'Low', 'Remove or redesign. Do not invest further without data.'],
        ]
    )

    add_box_block(doc, 'WATCH OUT',
        'Students may classify based on how much they personally like a drink. Redirect: the matrix is not about '
        'your taste — it is about what the guest orders and what the margin produces. If a drink you love never '
        'sells at an acceptable margin, it is a Dog. That is not an opinion — it is a number.')

    add_quote(doc,
        'The matrix is a thinking tool, not a verdict. You need sales data to confirm. But applying it forces you '
        'to ask the right questions about every item before you have the data. Most bars never ask those questions at all.')

    add_body(doc, 'DISCUSSION PROMPT', bold=True)
    add_body(doc,
        'Look at the Cannery Daiquiri at $14. It is a simple two-ingredient shaken build with high margin. '
        'But it is not featured, not described, and buried mid-list. What category is it most likely in? '
        'What single change would have the most impact?')

    add_transition(doc,
        'The matrix tells us what to do with each item. Menu design principles tell us how to do it on the page.')

    # ══════════════════════════════════════════════════
    # TM T4
    # ══════════════════════════════════════════════════
    add_theory_header_tm(doc, 5, 'T4', 'Menu Design Principles')

    add_labelled_block(doc, 'CORE PRINCIPLE',
        'Where an item is placed, what it is called, and how it is described are all decisions that affect sales. '
        'Every element either works for you or against you.')

    add_heading(doc, 'WHAT TO SAY', level=3)
    add_body(doc,
        'Golden Triangle: the eye moves to top-right first, then top-left, then center. Your highest-margin item '
        'goes in the highest-attention position. This is not a preference — it is documented in eye-tracking research '
        'on menu behaviour.')
    add_body(doc,
        'No dollar signs. The $ symbol activates the part of the brain associated with pain and loss. Removing it '
        'consistently increases average spend. Prices listed as numerals only (15, not $15.00) feel lower even '
        'when the number is identical.')
    add_body(doc,
        'Anchor pricing works because humans judge value relationally, not absolutely. A $26 cocktail on a menu '
        'of $14–17 drinks does not sell well — but it makes every other item feel reasonable. That is its job.')
    add_body(doc,
        'Sensory language sells. Wansink (2001) showed menus using descriptive language sold 27% more of those '
        'items than the same items listed as ingredient inventories. One word of texture or mood is enough. '
        '"Bright and crushable" is a description. "Gin, strawberry, lemon, soda" is a shopping list.')
    add_body(doc,
        'Section architecture: maximum three sections. Section names must communicate instantly. Scan — do not read. '
        '"Cannery Cocktails" vs "Cannery Signature" communicates nothing. "CRAFT COCKTAILS" vs "CLASSICS" works in two words.')

    add_heading(doc, 'WHAT TO SHOW', level=3)
    add_body(doc,
        'Point to the Sultan\'s Tent menu. Count the cocktail sections: The Sultan\'s Classics, Sparklers, Martini, '
        'Sultan\'s Mixology. Ask: what is the difference between a Classic and a Mixology cocktail? '
        'If the student cannot answer immediately — the guest certainly cannot. That confusion costs a decision.')
    add_body(doc,
        'Then look at the descriptions. Point to "Old Moroccan Fashion — Rye infused with Moroccan spices, '
        'Angostura bitters, Orange bitters." Now compare to "Casablanca Sunrise — Reposado tequila, Maraschino '
        'liqueur, Amaro Lucano, orange and lemon juice." Ask: which one makes you want to order it? '
        'The first tells a story. The second is a recipe card.')

    add_two_col_table(doc,
        ['ELEMENT', 'STRONG PRACTICE', 'WEAK PRACTICE'],
        [
            ['Pricing', 'No $ signs. Anchor item at premium.', 'Flat band. $ symbol. Prices buried.'],
            ['Language', 'Sensory, brand-matched description.', 'Ingredient list only.'],
            ['Visual design', 'Box or bold on 1–2 hero items.', 'Everything highlighted or nothing.'],
            ['Section naming', '"CRAFT COCKTAILS" / "CLASSICS"', '"Cannery Cocktails" / "Cannery Signature"'],
            ['Item count', '10–14 cocktails maximum', '20+ items across 4+ sections'],
        ]
    )

    add_box_block(doc, 'WATCH OUT',
        'Students may suggest adding more callouts to fix a menu — more NEW! badges, more boxes, more bold. '
        'The opposite is true. Every callout added dilutes every other callout. More signals produce less response. '
        'The most important design decision is often what to remove.')

    add_quote(doc,
        'The menu you hand a guest is doing sales work before you say a word. A well-designed menu creates a guest '
        'who arrives at the bar already knowing what they want. A poorly designed menu creates a guest who asks '
        'the server to explain everything — and that explanation costs time and table turns.')

    add_body(doc, 'DISCUSSION PROMPT', bold=True)
    add_body(doc,
        'If you were redesigning the Cannery cocktail section tonight — one change, that is all you can make — '
        'what would it be and why? Make the student commit to a single answer.')

    add_transition(doc,
        'Design tells guests what to order. Inventory control makes sure you have what they ordered. '
        'That is where we go next.')

    # ══════════════════════════════════════════════════
    # TM T5
    # ══════════════════════════════════════════════════
    add_theory_header_tm(doc, 5, 'T5', 'Inventory, Par Levels, and Control')

    add_labelled_block(doc, 'CORE PRINCIPLE',
        'A well-engineered menu only works if the inventory behind it is controlled. Par levels, FIFO, and variance '
        'tracking are not administrative tasks — they are the operating system of a profitable bar.')

    add_heading(doc, 'WHAT TO SAY', level=3)
    add_body(doc,
        'Par level: the minimum stock on hand to last until the next delivery, plus a safety buffer. '
        'Formula: (Daily usage × days between orders) + safety stock. '
        'Example: 4 bottles of vodka per day, ordering every 3 days, 2-bottle safety stock = par of 14. '
        'Running below par mid-service means running out during a rush — which means 86\'ing a cocktail and '
        'losing every subsequent order of that drink for the night.')
    add_body(doc,
        'Menu engineering connection: your best-selling cocktails drive high par on their spirits. '
        'A cocktail that rarely sells has a low par and slow inventory turnover. That slow-moving bottle is '
        'cash sitting on a shelf. Dogs on the menu create dead inventory. Remove the Dog, free the capital.')
    add_body(doc,
        'FIFO: First In, First Out. New stock behind old stock, always. Most critical for fresh juice '
        '(24-hour window), syrups (varies by type), and open wine (72 hours maximum refrigerated). '
        'FIFO is a quality standard as much as a cost control mechanism.')
    add_body(doc,
        'Variance: theoretical inventory is what the POS says you should have poured. Actual is the count. '
        'The gap is variance. Causes: over-pouring, spillage, theft, incorrect recipe execution, unrecorded comps. '
        'A standardised menu with consistent recipes narrows variance because every build is identical.')

    add_heading(doc, 'WHAT TO SHOW', level=3)
    add_body(doc,
        'Ask the student: how many different bottles does the Sultan\'s Tent cocktail menu require? Walk through '
        'counting unique spirits, liqueurs, and modifiers. The number will be significant. Then ask: how many of '
        'those appear in more than one cocktail? The ones that appear in only one drink are the inventory risk. '
        'If that cocktail gets pulled, the bottle goes dead.')
    add_body(doc,
        'Then show the contrast: the Black Russian and White Russian in today\'s practical share two bottles. '
        'The only additional SKU for Build 02 is heavy cream. Two products, two different menu items, '
        'one shared inventory commitment. That is efficient engineering.')

    add_two_col_table(doc,
        ['TERM', 'DEFINITION', 'BAR IMPACT'],
        [
            ['Par Level', 'Min stock to last until next delivery + safety buffer.', 'Running below par = 86\'d menu items mid-service.'],
            ['SKU', 'One unique product on the back bar.', 'Fewer SKUs = faster counts, less dead stock.'],
            ['FIFO', 'Oldest stock poured first, always.', 'Critical for juice, syrups, open wine.'],
            ['Dead Inventory', 'Product that sits unsold.', 'Cash tied up. Counts slower. Margin drag.'],
            ['Variance', 'Gap between theoretical and actual inventory.', 'Indicates over-pour, spillage, or theft.'],
        ]
    )

    add_box_block(doc, 'WATCH OUT',
        'Do not let this section become abstract. If the student\'s eyes glaze over, bring it back to a drink: '
        '"Remember the Sultan\'s Mule with 3 spirits? If that drink sells 3 times a week, those 3 bottles '
        'turn slowly. Now remove the drink — those bottles might not move at all. That is what dead inventory '
        'feels like in a physical count."')

    add_quote(doc,
        'A Dog on the menu is a problem in multiple departments simultaneously. It underperforms on the matrix, '
        'creates dead inventory, complicates par calculations, and slows counts. Removing one bad drink is not '
        'a menu decision — it is an operational decision.')

    add_body(doc, 'DISCUSSION PROMPT', bold=True)
    add_body(doc,
        'Name one cocktail from either menu that you would remove tonight based purely on its likely inventory impact. '
        'What would you do with the bottles it required?')

    add_transition(doc,
        'We have the theory. We have the tools. Now we apply all of it in one practical example — '
        'the same cocktail, engineered three ways, with the numbers on the table.')

    # ══════════════════════════════════════════════════
    # TM T6
    # ══════════════════════════════════════════════════
    add_theory_header_tm(doc, 4, 'T6', 'Cocktail Engineering in Practice')

    add_labelled_block(doc, 'CORE PRINCIPLE',
        'Three deliberate modifications to a single drink can justify a 25–30% price increase while adding under '
        '$0.50 in ingredient cost. That is cocktail engineering — and it is why menus are designed, not assembled.')

    add_heading(doc, 'WHAT TO SAY', level=3)
    add_body(doc,
        'Menu composition: a well-built cocktail menu covers all major drinking occasions. Minimum one entry per '
        'category. If a category is missing, the bar loses a sale every time a guest at that table wants that '
        'style of drink. Not an inconvenience — a lost revenue line.')
    add_body(doc,
        'The six categories: Sour (shaken, citrus-forward), Spirit-Forward (stirred), Sparkling/Light, '
        'Sweet/Indulgent, Long/Refreshing, Zero-Proof. A menu without zero-proof loses every table with one '
        'non-drinker. A menu without spirit-forward loses the serious drinker. Missing categories are missing sales.')
    add_body(doc,
        'The modification principle: three types of modification cost under $0.50 each and justify a $3–5 price '
        'increase. Dehydrated citrus (visual premium, weeks of shelf life), aromatic spray or expressed peel '
        '(aroma before first sip — the guest experiences the drink before tasting it), and infusion (flavour '
        'depth and a menu story no other bar can copy without your recipe).')
    add_body(doc,
        'The Cosmopolitan elevation proves all three simultaneously. Same base spirit. Same architecture. '
        'Three modifications totalling $0.42 in additional cost. $4 in additional price. Pour cost improves. '
        'Guest experience is categorically different. That is the argument for professional menu design.')

    add_heading(doc, 'WHAT TO SHOW', level=3)
    add_body(doc,
        'Hold up the jar of lemon peel infused vodka. Ask: what did this cost? Answer: the vodka was already '
        'purchased for other cocktails. The lemon peels would have gone to the juice station or waste. '
        'The infusion took 30 seconds to prepare and 48 hours to develop. The unit cost per drink is negligible.')
    add_body(doc,
        'Hold up a dehydrated citrus wheel. Ask: what did this cost? Answer: oven at low heat, 4–6 hours, '
        'fruit that was going to be juiced anyway. Unit cost per wheel: $0.10–0.15. Guest perception: premium, '
        'craft, deliberate. That gap between cost and perception is margin.')
    add_body(doc,
        'Show the bitters atomizer. One spray over the inside of the coupe before straining. Ask the student '
        'to smell the glass before the drink goes in. Then after. That aroma delivery — before the first sip — '
        'changes the entire experience. The guest smells complexity before they taste anything. That is worth $1.')

    add_two_col_table(doc,
        ['MODIFICATION', 'COST', 'CONTRIBUTION', 'PRICE JUSTIFICATION'],
        [
            ['Dehydrated citrus wheel', '~$0.12', 'Visual premium. Weeks of shelf life.', '+$1–2'],
            ['Aromatic spray / expressed peel', '~$0.05', 'Aroma before first sip. Engages sense before taste.', '+$1–2'],
            ['Spirit or syrup infusion (24–48h)', '~$0.15–0.25', 'Depth. Uniqueness. Brand story.', '+$2–3'],
            ['All three combined', '~$0.42', 'Full elevation. Photographable. Memorable. Recommended.', '+$4–5'],
        ]
    )

    add_box_block(doc, 'WATCH OUT',
        'If a student asks "what if guests don\'t notice?" — that is the wrong question. Every guest who receives '
        'a drink with a dehydrated wheel, an expressed peel aroma, and depth from an infusion notices something '
        'is different, even if they cannot name it. They feel the difference. They photograph it. They recommend it. '
        'Guests do not need to understand the technique to respond to it.')

    add_quote(doc,
        'This is why people pay money to create a menu instead of just a list. A list tells guests what is available. '
        'A menu tells them a story — and the story has a price.')

    add_body(doc, 'DISCUSSION PROMPT', bold=True)
    add_body(doc,
        'Ask: if you were opening your own bar tomorrow and you could only have six cocktails on the menu — '
        'one from each category — what would each one be? Make the student name all six and explain the category '
        'it covers. This is the session\'s closing exercise before the practical begins.')

    add_transition(doc,
        'Theory is done. The numbers are on the table. Now go behind the bar and prove them.')

    add_divider(doc)

    # ══════════════════════════════════════════════════
    # TM PRACTICAL NOTES
    # ══════════════════════════════════════════════════
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('PRACTICAL — INSTRUCTOR NOTES')
    r.bold = True; r.font.size = Pt(14)
    r.font.color.rgb = RGBColor.from_string('1F3864')
    add_divider(doc)

    add_heading(doc, 'Station Setup — 5 min', level=2)
    add_body(doc, 'Required pre-session preparation:', bold=True)
    setup_items = [
        'Lemon peel infused vodka: prepare 24–48h before class. Pack fresh lemon peels in a sealed glass jar, cover with vodka, refrigerate. Strain through fine mesh sieve before class. Reserve the strained peels for juicing.',
        'Dehydrated citrus wheels: slice limes, lemons, and/or oranges to 4mm thickness. Place on a wire rack in an oven at 90°C / 195°F for 4–6 hours or until dry and pliable. Cool completely. Store airtight. Prepare the day before.',
        'Bitters atomizer: load a small spray bottle with orange bitters or Peychaud\'s. Test the mist pattern before class.',
        'Pre-chill two coupes per student for Builds 03 and 04.',
        'Have both printed menus (Cannery Kitchen and Sultan\'s Tent) available at the station.',
        'White cranberry juice, Cointreau, coffee liqueur, heavy cream, vodka (house), fresh limes and lemons for squeezing.',
    ]
    for item in setup_items:
        p = doc.add_paragraph(f'— {item}')
        p.runs[0].font.size = Pt(10)
        p.paragraph_format.space_after = Pt(3)

    add_heading(doc, 'Build 01 — Black Russian (10 min)', level=2)
    add_body(doc,
        'The teaching point is efficiency. Two ingredients, no shaking, no straining. Walk through the cost calculation '
        'live as the student builds: $2.80 total cost ÷ $14 selling price × 100 = 20%. On target. Ask: is this a '
        'Star, Plowhorse, Puzzle, or Dog? Answer depends on sales volume — but pour cost says it could be a Star. '
        'Instruct the student to straw-taste and call the balance before presenting.')

    add_two_col_table(doc,
        ['EVALUATION POINT', 'STANDARD'],
        [
            ['Measurement', '2oz vodka exactly. 1oz coffee liqueur exactly.'],
            ['Ice', 'Glass properly iced before pouring.'],
            ['Stir', '3–4 slow rotations. Not a shake.'],
            ['Straw-taste', 'Called before presenting.'],
            ['Cost calculation', 'Student states $2.80 cost, $14 price, 20% pour cost.'],
        ]
    )

    add_heading(doc, 'Build 02 — White Russian (8 min)', level=2)
    add_body(doc,
        'Establish the menu extension principle before pouring. Ask: what is different? One ingredient — the cream float. '
        'What does it add? Texture, visual layering, a different flavour progression. What does it cost? $0.25 more. '
        'What can we charge? $1 more. What is the additional profit? $0.75 per drink. '
        'That is a menu extension decision made with data.')
    add_body(doc,
        'Critical technique: the cream must be floated, not stirred. Use the back of a bar spoon held just above '
        'the surface. Pour slowly. The cream should sit in a distinct layer. If it sinks or mixes, the temperature '
        'is wrong (cream too warm or drink under-chilled) or the pour was too fast. Rebuild and explain why.')

    add_two_col_table(doc,
        ['EVALUATION POINT', 'STANDARD'],
        [
            ['Base build', 'Identical to Build 01 — same measurements, same sequence.'],
            ['Cream float', 'Poured over back of spoon. Distinct layer visible.'],
            ['No stir after cream', 'Do not incorporate. The visual and texture layer is the point.'],
            ['Cost delta stated', 'Student states +$0.25 cost, +$1 price, +$0.75 profit per drink.'],
        ]
    )

    add_heading(doc, 'Build 03 — Classic Cosmopolitan (12 min)', level=2)
    add_body(doc,
        'Before building, have the student look at the Cannery menu and find the Cosmopolitan. Ask: where is it on '
        'the page? What section? What does the description say? Is it featured in any way? Answer: it is listed in '
        '"Cannery Signature" between a Mojito and a Paloma with a simple ingredient list. No callout. No story. '
        'That is the problem this build diagnoses.')
    add_body(doc,
        'Build the drink to full spec. Calculate the cost live: $2.10 ÷ $15 = 14%. Ask: is this a Star or a Puzzle? '
        'Answer: excellent pour cost, but the menu positioning makes it a Puzzle. The recipe is not the problem — '
        'the engineering around it is. That is the setup for Build 04.')

    add_two_col_table(doc,
        ['EVALUATION POINT', 'STANDARD'],
        [
            ['Measurement', '1.5oz vodka + 0.5oz Triple Sec + 0.5oz lime + 0.5oz cranberry. 2oz total alcohol.'],
            ['Shake', '12–15 seconds. Hard shake. Fully chilled.'],
            ['Double strain', 'Fine mesh + Hawthorne strainer. No ice chips.'],
            ['Coupe', 'Pre-chilled. Wash line clean.'],
            ['Garnish', 'Expressed lemon twist. Pinched over glass, oil visible on surface, placed on rim.'],
            ['Matrix call', 'Student classifies and justifies: Puzzle. High margin, poor positioning.'],
        ]
    )

    add_heading(doc, 'Build 04 — Elevated Cosmopolitan (18 min)', level=2)
    add_body(doc,
        'This is the session\'s central argument, executed behind the bar. Walk through each modification before '
        'the student builds it, and explain what each one does before they can see or taste it.')
    add_body(doc, 'Modification 1 — Infused vodka:', bold=True)
    add_body(doc,
        'Hold up the jar of infused vodka. Show the colour — slightly more golden than uninfused. Let the student '
        'smell it before tasting. Then nose it alongside the house vodka. Ask: what is different? The lemon peel '
        'oils have bonded to the spirit. The cost is identical to the uninfused vodka — same bottle, same price, '
        '30 seconds of prep 48 hours ago.')
    add_body(doc, 'Modification 2 — Dehydrated citrus wheel:', bold=True)
    add_body(doc,
        'Hold up the wheel. Ask: what does this cost? Walk through the calculation: fruit cost per wheel approximately '
        '$0.10–0.15, made from fruit that was going to be juiced anyway, shelf life weeks. What does it signal to '
        'the guest? Craft, preparation, intentionality. That perception adds $1–2 to the guest\'s sense of value '
        'before they taste the drink.')
    add_body(doc, 'Modification 3 — Bitters spray:', bold=True)
    add_body(doc,
        'Have the student hold the empty pre-chilled coupe. Spray 2 pumps of orange bitters inside. Have them smell '
        'the glass before any liquid goes in. Ask: what do you get? The aroma reaches the guest before the first sip. '
        'This is the most cost-efficient modification on the list — under $0.05 per drink, and it changes the sensory '
        'experience before the guest tastes anything.')
    add_body(doc,
        'Now build the full elevated version. Walk through the cost comparison table. State the numbers out loud: '
        '$2.10 becomes $2.52. $15 becomes $19. Pour cost goes from 14% to 13.3%. Profit per drink increases by '
        '$3.58. The modifications did not just improve the guest experience — they improved the margin.')

    add_two_col_table(doc,
        ['EVALUATION POINT', 'STANDARD'],
        [
            ['Infused vodka', '1.5oz. Measured exactly. Student can describe what the infusion adds.'],
            ['Bitters spray', '2 pumps inside coupe before straining. Student noses the glass and describes.'],
            ['Shake and double strain', 'Same standard as Build 03.'],
            ['Dehydrated wheel garnish', 'Positioned on rim. Not dropped in the drink.'],
            ['Expressed orange peel', 'Pinched over surface with visible oil mist. Placed on rim opposite wheel.'],
            ['Cost comparison', 'Student states all four rows of the comparison table from memory.'],
        ]
    )

    add_heading(doc, 'Menu Analysis Exercise (5 min)', level=2)
    add_body(doc,
        'Student selects one cocktail from either menu. They calculate the pour cost, classify it on the matrix, '
        'identify one modification under $0.50 with a justified price increase, and write a new 15-word maximum '
        'description in brand voice. Review the description together: does it use sensory language? Does it match '
        'the brand? Does it tell you why to order it, not just what is in it?')

    add_heading(doc, 'Knowledge Test (2 min) — Answer Key', level=2)
    add_two_col_table(doc,
        ['QUESTION', 'CORRECT ANSWER'],
        [
            ['Pour cost formula', '(Ingredient cost ÷ Selling price) × 100'],
            ['Wine bottle $48, glass $11. Correctly priced?', '$11 × 4 = $44. Under-priced by glass. Raise BTG to $12–13 or lower bottle to $44.'],
            ['Four matrix cells', 'Star (high/high), Plowhorse (low/high), Puzzle (high/low), Dog (low/low). One action for each.'],
            ['FIFO — where most critical?', 'First In, First Out. Most critical for fresh juice (24h), syrups, open wine (72h refrigerated).'],
            ['Additional profit — Elevated Cosmo', '$19 − $2.52 = $16.48 profit vs $15 − $2.10 = $12.90. Additional profit = $3.58 per drink.'],
        ]
    )

    # footer
    add_divider(doc)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run('Oleks Iurchenko | bar-skills.com | info@bar-skills.com')
    r.font.size = Pt(9)
    r.font.color.rgb = RGBColor.from_string('7F7F7F')

    doc.save('/home/user/OI1111/BarSkills_S09_TM.docx')
    print('TM saved.')


build_sb()
build_tm()
print('Done.')
