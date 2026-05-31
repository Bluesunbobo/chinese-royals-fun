import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# Define project working directory
proj_dir = "/Users/rennuoya/Desktop/RPA/Chinese Royals Fun"

def add_header(slide, title_text):
    # Colors
    GREEN = RGBColor(22, 101, 52)
    GOLD = RGBColor(181, 148, 50)
    
    # 1. Left decorative bar (Gold)
    left_bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0.6), Inches(0.4), Inches(0.08), Inches(0.6)
    )
    left_bar.fill.solid()
    left_bar.fill.fore_color.rgb = GOLD
    left_bar.line.fill.background() # no border
    
    # 2. Slide Title
    title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.35), Inches(11.5), Inches(0.7))
    tf = title_box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.name = "SimHei"
    p.font.size = Pt(28)
    p.font.bold = True
    p.font.color.rgb = GREEN
    
    # 3. Horizontal subtle divider line
    div_line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0.6), Inches(1.15), Inches(12.133), Inches(0.02)
    )
    div_line.fill.solid()
    div_line.fill.fore_color.rgb = RGBColor(226, 232, 240) # light gray border
    div_line.line.fill.background()

def set_slide_background(slide, color_rgb):
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = color_rgb

def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    # Common layouts
    blank_layout = prs.slide_layouts[6]
    
    # Brand Colors
    DARK_GREEN = RGBColor(11, 60, 29) # Premium Dark Green for cover/end
    CREAM = RGBColor(250, 249, 246) # Warm Cream for content background
    GOLD = RGBColor(181, 148, 50) # Antique Gold
    WHITE = RGBColor(255, 255, 255)
    TEXT_MAIN = RGBColor(30, 41, 59) # Slate Dark Gray
    TEXT_MUTED = RGBColor(71, 85, 105) # Slate Soft Gray
    
    # ==========================================
    # SLIDE 1: COVER SLIDE (Dark Green Background)
    # ==========================================
    slide1 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide1, DARK_GREEN)
    
    # Top Gold accent line
    top_accent = slide1.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(0), Inches(13.333), Inches(0.15)
    )
    top_accent.fill.solid()
    top_accent.fill.fore_color.rgb = GOLD
    top_accent.line.fill.background()
    
    # Title Box
    title_box = slide1.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.333), Inches(1.8))
    tf = title_box.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    p1.text = "御苑福膳"
    p1.font.name = "SimHei"
    p1.font.size = Pt(64)
    p1.font.bold = True
    p1.font.color.rgb = GOLD
    p1.alignment = PP_ALIGN.LEFT
    p1.space_after = Pt(20)
    
    p2 = tf.add_paragraph()
    p2.text = "专注本土超级食物 · 开启健康膳食新风尚"
    p2.font.name = "Microsoft YaHei"
    p2.font.size = Pt(24)
    p2.font.color.rgb = WHITE
    p2.alignment = PP_ALIGN.LEFT
    
    # Subtitle / Company Box
    company_box = slide1.shapes.add_textbox(Inches(1.0), Inches(5.0), Inches(11.333), Inches(1.5))
    tf2 = company_box.text_frame
    tf2.word_wrap = True
    
    pc1 = tf2.paragraphs[0]
    pc1.text = "北京中卫神州生物科技有限公司 创立"
    pc1.font.name = "Microsoft YaHei"
    pc1.font.size = Pt(16)
    pc1.font.color.rgb = RGBColor(200, 200, 200)
    pc1.space_after = Pt(8)
    
    pc2 = tf2.add_paragraph()
    pc2.text = "北京御苑福膳文化传播有限公司 运营管理"
    pc2.font.name = "Microsoft YaHei"
    pc2.font.size = Pt(16)
    pc2.font.color.rgb = RGBColor(200, 200, 200)

    # ==========================================
    # SLIDE 2: BRAND ORIGIN & RESEARCH BACKGROUND (Cream Background)
    # ==========================================
    slide2 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide2, CREAM)
    add_header(slide2, "品牌起源与中医药科研背景")
    
    # Left Box - Description Card
    left_card = slide2.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.6), Inches(1.5), Inches(5.8), Inches(5.2)
    )
    left_card.fill.solid()
    left_card.fill.fore_color.rgb = WHITE
    left_card.line.color.rgb = RGBColor(226, 232, 240)
    
    left_tf = left_card.text_frame
    left_tf.margin_left = Inches(0.4)
    left_tf.margin_right = Inches(0.4)
    left_tf.margin_top = Inches(0.4)
    left_tf.word_wrap = True
    
    p = left_tf.paragraphs[0]
    p.text = "深厚中医药学科研背景"
    p.font.name = "SimHei"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = DARK_GREEN
    p.space_after = Pt(20)
    
    p = left_tf.add_paragraph()
    p.text = "“御苑福膳”由北京中卫神州生物科技有限公司创立。公司成立于2002年，原隶属于中国医学科学院药用植物研究所（国家级药用植物研究权威机构）下属企业。"
    p.font.name = "Microsoft YaHei"
    p.font.size = Pt(15)
    p.font.color.rgb = TEXT_MAIN
    p.space_after = Pt(15)
    
    p = left_tf.add_paragraph()
    p.text = "现由中卫神州全资子公司北京御苑福膳文化传播有限公司进行专业的品牌运营与文化推广。公司依托中医药科研数十载的历史积淀，致力于打造面向全民大健康与健康膳食的餐饮标杆品牌。"
    p.font.name = "Microsoft YaHei"
    p.font.size = Pt(15)
    p.font.color.rgb = TEXT_MAIN
    
    # Right Box - Expert Guidance List
    right_card = slide2.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(6.8), Inches(1.5), Inches(5.9), Inches(5.2)
    )
    right_card.fill.solid()
    right_card.fill.fore_color.rgb = WHITE
    right_card.line.color.rgb = RGBColor(226, 232, 240)
    
    right_tf = right_card.text_frame
    right_tf.margin_left = Inches(0.4)
    right_tf.margin_right = Inches(0.4)
    right_tf.margin_top = Inches(0.4)
    right_tf.word_wrap = True
    
    p = right_tf.paragraphs[0]
    p.text = "国医大师与权威专家学术指导"
    p.font.name = "SimHei"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = DARK_GREEN
    p.space_after = Pt(20)
    
    bullet_points = [
        "中药学泰斗、中国工程院院士肖培根教授长期担纲学术与应用指导专家，确保草本膳食科学底蕴；",
        "国医大师、中国工程院院士王琦教授等业内泰斗级专家亲自莅临并对时令膳食开发提供科学理论支撑；",
        "汇聚多位资深中医中药学者、现代营养师与国家级烹饪大师，跨界联手构建“健康餐桌”系统工程；",
        "倡导“吃得科学、做得分外精致”，为公众的体质调理与亚健康改善提供切实有效的膳食方案。"
    ]
    for bp in bullet_points:
        p = right_tf.add_paragraph()
        p.text = "• " + bp
        p.font.name = "Microsoft YaHei"
        p.font.size = Pt(14)
        p.font.color.rgb = TEXT_MUTED
        p.space_after = Pt(12)

    # ==========================================
    # SLIDE 3: DIETARY CONCEPT (Cream Background)
    # ==========================================
    slide3 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide3, CREAM)
    add_header(slide3, "核心理念：本土超级食物与新鲜烹饪")
    
    # Left Column
    left_card = slide3.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.6), Inches(1.5), Inches(5.8), Inches(5.2)
    )
    left_card.fill.solid()
    left_card.fill.fore_color.rgb = WHITE
    left_card.line.color.rgb = RGBColor(226, 232, 240)
    
    left_tf = left_card.text_frame
    left_tf.margin_left = Inches(0.4)
    left_tf.margin_right = Inches(0.4)
    left_tf.margin_top = Inches(0.4)
    left_tf.word_wrap = True
    
    p = left_tf.paragraphs[0]
    p.text = "新鲜食材 · 新鲜烹饪"
    p.font.name = "SimHei"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = DARK_GREEN
    p.space_after = Pt(20)
    
    points_left = [
        "坚持本真膳食体验，拒绝工业预制与死板调味，完全遵循“新鲜食材、现场即时烹饪”的工艺准则；",
        "精选并推广高营养素密度的中国本土超级食物，将日常朴素食材以极致的烹饪艺术做出千姿百态的花样；",
        "紧密围绕自然四季、二十四节气的时令更替进行菜谱轮换，牢牢锁住新鲜食材的原始生机、活性成分与充沛营养；",
        "以健康为底色，把大众餐桌的日常饮食升级为兼具美味享受与健康调理的精致艺术体验。"
    ]
    for pt in points_left:
        p = left_tf.add_paragraph()
        p.text = "• " + pt
        p.font.name = "Microsoft YaHei"
        p.font.size = Pt(14)
        p.font.color.rgb = TEXT_MAIN
        p.space_after = Pt(12)
        
    # Right Column
    right_card = slide3.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(6.8), Inches(1.5), Inches(5.9), Inches(5.2)
    )
    right_card.fill.solid()
    right_card.fill.fore_color.rgb = WHITE
    right_card.line.color.rgb = RGBColor(226, 232, 240)
    
    right_tf = right_card.text_frame
    right_tf.margin_left = Inches(0.4)
    right_tf.margin_right = Inches(0.4)
    right_tf.margin_top = Inches(0.4)
    right_tf.word_wrap = True
    
    p = right_tf.paragraphs[0]
    p.text = "中药服食与时令食疗创新"
    p.font.name = "SimHei"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = DARK_GREEN
    p.space_after = Pt(20)
    
    points_right = [
        "首创“鲜药入膳”核心专利模式，打破传统干药材煲汤熬煮的陈旧思维，探索生鲜本草与现代烹饪的结合；",
        "直接依托高科技鲜中药材大棚与生态基地，现采现用藿香、紫苏、玉竹、地黄、鲜灵芝等名贵时令本草；",
        "科学融合中医药饮食理论、中医服食养生学及现代食疗临床营养成果，研发四季调理养生方；",
        "精心构筑科学、安全的亚健康防御膳食屏障，将博大精深的中医文化推向千家万户的大众家庭餐桌。"
    ]
    for pt in points_right:
        p = right_tf.add_paragraph()
        p.text = "• " + pt
        p.font.name = "Microsoft YaHei"
        p.font.size = Pt(14)
        p.font.color.rgb = TEXT_MAIN
        p.space_after = Pt(12)

    # ==========================================
    # SLIDE 4: ECOLOGICAL BASES (Cream Background)
    # ==========================================
    slide4 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide4, CREAM)
    add_header(slide4, "生态种植基地与官方认定示范窗口")
    
    # Text card at top
    top_card = slide4.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.6), Inches(1.5), Inches(12.1), Inches(1.8)
    )
    top_card.fill.solid()
    top_card.fill.fore_color.rgb = WHITE
    top_card.line.color.rgb = RGBColor(226, 232, 240)
    
    top_tf = top_card.text_frame
    top_tf.margin_left = Inches(0.4)
    top_tf.margin_top = Inches(0.3)
    top_tf.word_wrap = True
    
    p = top_tf.paragraphs[0]
    p.text = "中医药文化旅游示范基地（北京市中医管理局官方认定）"
    p.font.name = "SimHei"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = DARK_GREEN
    p.space_after = Pt(8)
    
    p = top_tf.add_paragraph()
    p.text = "该基地由北京市中医管理局与北京市旅游局联合评定、授权并颁发官方示范牌照。它是经政府相关部门严格评审与行业权威认定的中医药养生与膳食科普示范标杆窗口，汇聚百余种鲜活本草植物，是一座融植物科普、中药识药与健康膳食体验于一体的权威生态示范园区。"
    p.font.name = "Microsoft YaHei"
    p.font.size = Pt(14)
    p.font.color.rgb = TEXT_MAIN
    
    # 8 Bases Grid
    grid_start_y = Inches(3.6)
    grid_w = Inches(2.9)
    grid_h = Inches(1.5)
    gap_x = Inches(0.16)
    gap_y = Inches(0.15)
    
    bases = [
        ("国家级科研基地", "依托中国医学科学院药植所，开展鲜药材种质资源研究与繁育工作"),
        ("鲜药材大棚基地", "数十亩现代化生态大棚，现场带露采摘，供应即采即用的盎然鲜活本草"),
        ("百草示范园", "北京市中医管理局颁牌，融科普、识药、健康膳食为一体的示范中心"),
        ("原生态灵芝基地", "深山林区环境抚育，活体孢子切片采收，饱含灵芝活性多糖营养"),
        ("宁夏枸杞基地", "黄河冲积平原种植，肉厚汁甜，含丰富类胡萝卜素，清晨采摘鲜用"),
        ("阿拉善肉苁蓉基地", "沙漠边缘纯天然梭梭寄生，手工鲜取，提供极高活性的传统滋补动力"),
        ("承德金莲花基地", "高山草甸无污染带露采摘，泡茶入膳清热解毒，香气持久芬芳"),
        ("青藏玛咖基地", "高海拔强日照极寒山区抚育，汲取雪域精华，能量充沛天然滋养")
    ]
    
    for i, (title, desc) in enumerate(bases):
        row = i // 4
        col = i % 4
        x = Inches(0.6) + col * (grid_w + gap_x)
        y = grid_start_y + row * (grid_h + gap_y)
        
        card = slide4.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            x, y, grid_w, grid_h
        )
        card.fill.solid()
        card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = RGBColor(22, 101, 52)
        card.line.width = Pt(1)
        
        tf = card.text_frame
        tf.margin_left = Inches(0.2)
        tf.margin_right = Inches(0.2)
        tf.margin_top = Inches(0.15)
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = "SimHei"
        p.font.size = Pt(14)
        p.font.bold = True
        p.font.color.rgb = GOLD
        p.space_after = Pt(5)
        
        p = tf.add_paragraph()
        p.text = desc
        p.font.name = "Microsoft YaHei"
        p.font.size = Pt(10)
        p.font.color.rgb = TEXT_MUTED

    # ==========================================
    # SLIDE 5: HONORS & QUALIFICATIONS (Cream Background)
    # ==========================================
    slide5 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide5, CREAM)
    add_header(slide5, "品牌资质与社会荣誉认可")
    
    # 5 Honors layout - 5 cards
    honor_w = Inches(3.8)
    honor_h = Inches(2.4)
    gap_x = Inches(0.3)
    gap_y = Inches(0.3)
    
    honors_list = [
        ("非物质文化遗产", "传统膳食技艺非遗传承保护单位", "多年来始终坚持中华古老膳食技艺的精髓继承与现代保护，非物质文化遗产的传承，见证了御苑福膳深厚的文化底蕴。"),
        ("老字号荣誉", "传统膳食品牌与老字号荣誉", "荣获中医药及健康餐饮行业内的老字号名号，代表着数十载持之以恒的卓越声誉、工匠精神以及官方与社会的信任。"),
        ("行业荣誉资质", "中医药食同源与特色康养餐饮认证", "被行业协会与健康养生协会联合认证为药食同源与特色健康餐饮示范店，体现出深厚的草本膳食科学研究底蕴。"),
        ("首都窗口行业示范单位", "2008年北京迎奥运技能竞赛示范", "在北京市劳动和社会保障局、首都窗口行业奥运培训工作协调小组办公室主办的技能大赛中，被授予“首都窗口行业技能示范单位”。"),
        ("安全生产标准化企业", "北京市海淀区安全生产监督管理局颁发", "作为北京市海淀区餐饮业安全生产标准化的典范，连续多年荣获安全生产达标资质，用极高安全标准保障每一位食客的健康安全。")
    ]
    
    # Positions: 3 on top row, 2 centered on bottom row
    for i, (title, subtitle, desc) in enumerate(honors_list):
        if i < 3:
            x = Inches(0.6) + i * (honor_w + gap_x)
            y = Inches(1.5)
        else:
            # Centered on bottom row
            x = Inches(2.65) + (i - 3) * (honor_w + gap_x)
            y = Inches(4.2)
            
        card = slide5.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            x, y, honor_w, honor_h
        )
        card.fill.solid()
        card.fill.fore_color.rgb = WHITE
        card.line.color.rgb = RGBColor(226, 232, 240)
        
        tf = card.text_frame
        tf.margin_left = Inches(0.25)
        tf.margin_right = Inches(0.25)
        tf.margin_top = Inches(0.2)
        tf.word_wrap = True
        
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = "SimHei"
        p.font.size = Pt(16)
        p.font.bold = True
        p.font.color.rgb = DARK_GREEN
        p.space_after = Pt(4)
        
        p = tf.add_paragraph()
        p.text = subtitle
        p.font.name = "Microsoft YaHei"
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = GOLD
        p.space_after = Pt(8)
        
        p = tf.add_paragraph()
        p.text = desc
        p.font.name = "Microsoft YaHei"
        p.font.size = Pt(10)
        p.font.color.rgb = TEXT_MUTED

    # ==========================================
    # SLIDE 6: CELEBRITY FOOTPRINTS (Cream Background)
    # ==========================================
    slide6 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide6, CREAM)
    add_header(slide6, "社会各界与中医药学专家交流足迹")
    
    # Left Box - Political leaders
    left_card = slide6.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.6), Inches(1.5), Inches(5.8), Inches(5.2)
    )
    left_card.fill.solid()
    left_card.fill.fore_color.rgb = WHITE
    left_card.line.color.rgb = RGBColor(226, 232, 240)
    
    left_tf = left_card.text_frame
    left_tf.margin_left = Inches(0.4)
    left_tf.margin_right = Inches(0.4)
    left_tf.margin_top = Inches(0.4)
    left_tf.word_wrap = True
    
    p = left_tf.paragraphs[0]
    p.text = "党政领导与社会名流莅临视察"
    p.font.name = "SimHei"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = DARK_GREEN
    p.space_after = Pt(20)
    
    leaders = [
        ("全国人大常委会副委员长 何维", "莅临中卫御苑福膳，关心并考察健康膳食与康养产业交流工作；"),
        ("全国政协原副主席 张思卿 & 孙孚凌", "视察指导，对中草药健康膳食在大众层面的传承推广予以厚望；"),
        ("原卫生部部长 张文康 & 高强、副部长 顾英奇", "亲临指导，鼓励传承好非遗技艺，将草本健康安全地融入日常饮食；"),
        ("甘肃省委原书记 陆浩、海南省委原书记 卫留成等", "先后莅临考察，赞许御苑福膳“以膳食守护大众健康”的企业使命；"),
        ("毛泽东主席原机要秘书 张玉凤女士与刘爱民先生", "亲临现场，品鉴时令健康鲜药膳并予以极高评价。")
    ]
    for title, desc in leaders:
        p = left_tf.add_paragraph()
        p.text = f"★ {title}："
        p.font.name = "SimHei"
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = GOLD
        p = left_tf.add_paragraph()
        p.text = desc
        p.font.name = "Microsoft YaHei"
        p.font.size = Pt(12)
        p.font.color.rgb = TEXT_MAIN
        p.space_after = Pt(8)
        
    # Right Box - Experts & Masters
    right_card = slide6.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(6.8), Inches(1.5), Inches(5.9), Inches(5.2)
    )
    right_card.fill.solid()
    right_card.fill.fore_color.rgb = WHITE
    right_card.line.color.rgb = RGBColor(226, 232, 240)
    
    right_tf = right_card.text_frame
    right_tf.margin_left = Inches(0.4)
    right_tf.margin_right = Inches(0.4)
    right_tf.margin_top = Inches(0.4)
    right_tf.word_wrap = True
    
    p = right_tf.paragraphs[0]
    p.text = "院士泰斗与国宴大师技术合作"
    p.font.name = "SimHei"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = DARK_GREEN
    p.space_after = Pt(20)
    
    experts = [
        ("中药学权威 肖培根院士", "中国工程院院士、中药学泰斗亲自指导，从植物药理角度确保膳食的调理功效；"),
        ("国医大师 王琦院士", "中国工程院院士、国医大师王琦对“药食同源”的日常科学应用与标准化给予高度赞扬；"),
        ("医学泰斗 刘德培 & 顾方舟院士", "原工程院副院长刘德培、“脊灰疫苗之父”顾方舟对餐桌大健康工程表示极高认可；"),
        ("国宴泰斗 侯仲华 & 孟宪斌大师", "原钓鱼台国宾馆副总厨师长侯仲华、特级国宴大师孟宪斌亲临指导烹饪技术，将传统草本与国宴级别膳食艺术结合；"),
        ("国际与多边友好交流", "承接联合国北北合作组织、西藏活佛及多国大使与贵宾考察，向世界宣传中国膳食智慧。")
    ]
    for title, desc in experts:
        p = right_tf.add_paragraph()
        p.text = f"★ {title}："
        p.font.name = "SimHei"
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = GOLD
        p = right_tf.add_paragraph()
        p.text = desc
        p.font.name = "Microsoft YaHei"
        p.font.size = Pt(12)
        p.font.color.rgb = TEXT_MAIN
        p.space_after = Pt(8)

    # ==========================================
    # SLIDE 7: SOCIAL RESPONSIBILITY & MEDIA (Cream Background)
    # ==========================================
    slide7 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide7, CREAM)
    add_header(slide7, "大健康政策响应、社会责任与媒体聚焦")
    
    # Left Column
    left_card = slide7.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.6), Inches(1.5), Inches(5.8), Inches(5.2)
    )
    left_card.fill.solid()
    left_card.fill.fore_color.rgb = WHITE
    left_card.line.color.rgb = RGBColor(226, 232, 240)
    
    left_tf = left_card.text_frame
    left_tf.margin_left = Inches(0.4)
    left_tf.margin_right = Inches(0.4)
    left_tf.margin_top = Inches(0.4)
    left_tf.word_wrap = True
    
    p = left_tf.paragraphs[0]
    p.text = "积极响应政策与践行社会责任"
    p.font.name = "SimHei"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = DARK_GREEN
    p.space_after = Pt(15)
    
    p = left_tf.add_paragraph()
    p.text = "「大健康战略与体重管理年」"
    p.font.name = "SimHei"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = GOLD
    p.space_after = Pt(5)
    
    p = left_tf.add_paragraph()
    p.text = "积极响应国家“大健康”战略，切实落实国家“体重管理年”倡议。开设“教你吃、教你做”时令食育工坊。每逢清明节（低卡青团）、儿童节（趣味营养餐）、端午节（草木香草粽）、中秋节（低糖低脂养生月饼）等，带领大众科学体重管理，吃出本真健康。"
    p.font.name = "Microsoft YaHei"
    p.font.size = Pt(12.5)
    p.font.color.rgb = TEXT_MAIN
    p.space_after = Pt(15)
    
    p = left_tf.add_paragraph()
    p.text = "「“爱心拐杖”云端老年健康食堂」"
    p.font.name = "SimHei"
    p.font.size = Pt(14)
    p.font.bold = True
    p.font.color.rgb = GOLD
    p.space_after = Pt(5)
    
    p = left_tf.add_paragraph()
    p.text = "作为首批定点保障单位，为海淀区马连洼街道的长者研发量身定制软烂易消化、控盐控脂的适老营养膳食。用实际行动守护老年人的舌尖幸福，为健康中国战略落地贡献社区温情力量。"
    p.font.name = "Microsoft YaHei"
    p.font.size = Pt(12.5)
    p.font.color.rgb = TEXT_MAIN
    
    # Right Column - Media Coverage
    right_card = slide7.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(6.8), Inches(1.5), Inches(5.9), Inches(5.2)
    )
    right_card.fill.solid()
    right_card.fill.fore_color.rgb = WHITE
    right_card.line.color.rgb = RGBColor(226, 232, 240)
    
    right_tf = right_card.text_frame
    right_tf.margin_left = Inches(0.4)
    right_tf.margin_right = Inches(0.4)
    right_tf.margin_top = Inches(0.4)
    right_tf.word_wrap = True
    
    p = right_tf.paragraphs[0]
    p.text = "主流媒体聚焦与社会公信力"
    p.font.name = "SimHei"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = DARK_GREEN
    p.space_after = Pt(15)
    
    media = [
        ("CCTV-1《夕阳红》清明/谷雨/立夏/小满美食", "2025年特别节目多次专访与专题报道，在全国范围内科普节气养生饮食与膳食调和方法；"),
        ("CGTN西班牙语频道 专题向全球播送", "“民生中国：药食同源的传统新实践”专题片向全球播放，代表中国向世界宣传传统美食文化现代化创新；"),
        ("CCTV央视频 假日美食录制特别直播", "央视知名主持人王曦梁亲临现场，进行药食同源假日健康美食制作直播，观看量极大；"),
        ("BRTV北京电视台《气象观天下》连续专题", "现场录制惊蛰、清明、端午、冬至等多期特别气象民俗美食，点赞侯仲华大师技术指导；"),
        ("《中国日报》&《中新网》多篇图文专题", "对马连洼云端老年食堂进行社会公益纪实报道，专访朱怀彬董事长“走南闯北长垣人”奋斗生平与老字号匠心。")
    ]
    for name, desc in media:
        p = right_tf.add_paragraph()
        p.text = f"● {name}："
        p.font.name = "SimHei"
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = GOLD
        p = right_tf.add_paragraph()
        p.text = desc
        p.font.name = "Microsoft YaHei"
        p.font.size = Pt(11)
        p.font.color.rgb = TEXT_MUTED
        p.space_after = Pt(8)

    # ==========================================
    # SLIDE 8: END SLIDE (Dark Green Background)
    # ==========================================
    slide8 = prs.slides.add_slide(blank_layout)
    set_slide_background(slide8, DARK_GREEN)
    
    # Bottom Gold accent line
    bottom_accent = slide8.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(7.35), Inches(13.333), Inches(0.15)
    )
    bottom_accent.fill.solid()
    bottom_accent.fill.fore_color.rgb = GOLD
    bottom_accent.line.fill.background()
    
    # Text Box
    end_box = slide8.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.333), Inches(3.0))
    tf_end = end_box.text_frame
    tf_end.word_wrap = True
    
    p = tf_end.paragraphs[0]
    p.text = "坚持新鲜食材与新鲜烹饪"
    p.font.name = "SimHei"
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = GOLD
    p.alignment = PP_ALIGN.CENTER
    p.space_after = Pt(15)
    
    p = tf_end.add_paragraph()
    p.text = "以时令膳食守护大众餐桌健康"
    p.font.name = "Microsoft YaHei"
    p.font.size = Pt(28)
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER
    p.space_after = Pt(40)
    
    p = tf_end.add_paragraph()
    p.text = "谢谢观看"
    p.font.name = "SimHei"
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = GOLD
    p.alignment = PP_ALIGN.CENTER
    
    # Save presentation
    filepath = os.path.join(proj_dir, "御苑福膳_品牌介绍.pptx")
    prs.save(filepath)
    print(f"Presentation generated successfully at {filepath}")

if __name__ == "__main__":
    create_presentation()
