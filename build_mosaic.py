# -*- coding: utf-8 -*-
"""生成镶嵌综合实践项目设计 Word：镶嵌综合实践_数智赋能项目设计.docx"""
import os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from PIL import Image, ImageDraw, ImageFont

BASE="/Users/chengqp/Desktop/tan"; IMG=os.path.join(BASE,"图"); os.makedirs(IMG,exist_ok=True)
OUT=os.path.join(BASE,"镶嵌综合实践_数智赋能项目设计.docx")
SONG,HEI,KAI="宋体","黑体","楷体"
FP="/System/Library/Fonts/Hiragino Sans GB.ttc"
def F(s): return ImageFont.truetype(FP,s)

def setf(r,cn=SONG,size=10.5,bold=False,color=None):
    r.font.name=cn; r._element.rPr.rFonts.set(qn('w:eastAsia'),cn); r.font.size=Pt(size); r.font.bold=bold
    if color: r.font.color.rgb=color
def para(doc,ind=2,after=2,line=1.5,align=None):
    p=doc.add_paragraph(); pf=p.paragraph_format
    if ind: pf.first_line_indent=Pt(10.5*ind)
    pf.space_after=Pt(after); pf.space_before=Pt(0); pf.line_spacing=line; pf.line_spacing_rule=WD_LINE_SPACING.MULTIPLE
    if align: p.alignment=align
    return p
def runs(p,segs,size=10.5):
    for t,st in segs:
        r=p.add_run(t); setf(r,SONG,size)
        if st=='b': r.font.bold=True
        elif st=='sup': r.font.superscript=True
    return p
def H(doc,t,size=12,before=8,after=4):
    p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(before); p.paragraph_format.space_after=Pt(after)
    r=p.add_run(t); setf(r,HEI,size,bold=True); return p

# 图1：项目六环节路径
def draw_ct(d,box,text,font,fill=(40,40,40)):
    x0,y0,x1,y1=box; ls=text.split("\n")
    hs=[d.textbbox((0,0),l,font=font)[3]-d.textbbox((0,0),l,font=font)[1] for l in ls]
    cy=(y0+y1)/2-(sum(hs)+(len(ls)-1)*6)/2
    for l,h in zip(ls,hs):
        w=d.textbbox((0,0),l,font=font)[2]-d.textbbox((0,0),l,font=font)[0]
        d.text(((x0+x1)/2-w/2,cy),l,font=font,fill=fill); cy+=h+6
def make_fig1():
    W,H=1580,560; im=Image.new("RGB",(W,H),"white"); d=ImageDraw.Draw(im)
    d.text((40,24),"「用数学识破 AI」镶嵌审辩课六环节 —— ★环节⑤为“非 AI 不可”的核心",font=F(28),fill=(15,15,15))
    boxes=[("① 赏纹样","越王勾践剑菱形纹\n引入二方/四方连续"),
           ("② 辨变换","识别单位纹样与\n平移·旋转·轴对称"),
           ("③ 还原埃舍尔","拆解鸽子瓷砖\n→基本图形→密铺"),
           ("④ 提猜想·裁判","先猜哪些能密铺\n再用角和=360°验证"),
           ("⑤ 审AI作品★","审辩 AI 生成的“镶嵌”\n用360°戳穿伪密铺"),
           ("⑥ 共创作品","AI+数学共造合法镶嵌\n提交·互评·展示")]
    n=len(boxes); mg=36; gap=16; bw=(W-2*mg-(n-1)*gap)/n; top=150; bh=280
    cols=[(232,244,253),(232,245,233),(255,243,224),(255,236,239),(237,231,246),(224,247,250)]
    lc=[(21,101,192),(56,142,60),(230,126,34),(200,60,90),(123,31,162),(0,131,143)]
    for i,(ti,bo) in enumerate(boxes):
        x0=mg+i*(bw+gap); x1=x0+bw; y0,y1=top,top+bh
        d.rounded_rectangle([x0,y0,x1,y1],radius=14,fill=cols[i],outline=lc[i],width=3)
        d.rounded_rectangle([x0,y0,x1,y0+44],radius=14,fill=lc[i])
        tw=d.textbbox((0,0),ti,font=F(23))[2]; d.text(((x0+x1)/2-tw/2,y0+7),ti,font=F(23),fill=(255,255,255))
        draw_ct(d,(x0+6,y0+52,x1-6,y1-8),bo,F(18))
        if i<n-1:
            ay=top+bh/2; a1=x1+gap-2
            d.line([x1+2,ay,a1-7,ay],fill=(120,120,120),width=4)
            d.polygon([(a1-7,ay-6),(a1,ay),(a1-7,ay+6)],fill=(120,120,120))
    by=top+bh+26
    d.rounded_rectangle([mg,by,W-mg,by+56],radius=12,fill=(250,250,250),outline=(180,180,180),width=2)
    tg="创新点：数学握有 AI 无法自证的判定权——AI 靠视觉“画”，我们靠 360°“判”；不信 AI 的眼，只信数学的理"
    tw=d.textbbox((0,0),tg,font=F(21))[2]; d.text((W/2-tw/2,by+15),tg,font=F(21),fill=(30,30,30))
    p=os.path.join(IMG,"图1_镶嵌项目路径.png"); im.save(p,dpi=(200,200)); return p
FIG1=make_fig1(); print("图1:",FIG1)


# ================= 文档 =================
doc=Document(); sec=doc.sections[0]
sec.page_width=Cm(21.0); sec.page_height=Cm(29.7)
sec.top_margin=Cm(2.3); sec.bottom_margin=Cm(2.3); sec.left_margin=Cm(2.6); sec.right_margin=Cm(2.6)
nm=doc.styles['Normal']; nm.font.name=SONG; nm.font.size=Pt(10.5); nm._element.rPr.rFonts.set(qn('w:eastAsia'),SONG)

t=doc.add_paragraph(); t.alignment=WD_ALIGN_PARAGRAPH.CENTER; t.paragraph_format.space_after=Pt(4)
r=t.add_run("“AI造图·数学裁判”：镶嵌图案中的数智赋能综合实践"); setf(r,HEI,16,bold=True)
s=doc.add_paragraph(); s.alignment=WD_ALIGN_PARAGRAPH.CENTER; s.paragraph_format.space_after=Pt(8)
r=s.add_run("——依托沪教版七上“综合与实践·从传统连续纹样到现代镶嵌图案”（约6课时）"); setf(r,KAI,11)

H(doc,"一、项目定位与课标锚点")
para(doc); runs(doc.paragraphs[-1],[
    ("本项目依托教材七年级上册“综合与实践”活动《从传统连续纹样到现代镶嵌图案》，属项目式学习。"
     "数学内核为图形的运动——平移、旋转、轴对称及其组合，二方连续、四方连续纹样，以及镶嵌（密铺）；"
     "跨艺术与传统文化（越王勾践剑菱形纹、埃舍尔镶嵌）。核心素养落点：",'')])
runs(doc.paragraphs[-1],[("几何直观、空间观念、创新意识",'b'),("（主），兼及应用意识。",'')])
para(doc); runs(doc.paragraphs[-1],[
    ("数智赋能定位一句话：",''),("“AI造图、数学裁判”",'b'),
    ("——让生成式人工智能生成“镶嵌”图案（这是AI真正擅长、也靠得住的事），"
     "再用数学去判定它到底用了哪些变换、是不是真正的密铺（无缝隙、不重叠、拼接点处角和为360°），"
     "从而把AI用在其所长，把数学立为裁判，避免“为技术而技术”。",'')])

H(doc,"二、选题四问自检")
para(doc,ind=0); runs(doc.paragraphs[-1],[
    ("现实?✅ 纹样与镶嵌广见于建筑、纺织、包装、地砖。　跨学科?✅ 融艺术、传统文化。　"
     "学生可解?✅ 依托平移/旋转/轴对称与角度计算，七年级可及。　能提出问题?✅ 由“AI画得对不对”自然生问。",'')])

H(doc,"三、驱动问题")
para(doc); runs(doc.paragraphs[-1],[
    ("AI能画出真正的镶嵌图案吗?",'b'),
    ("——请你用数学给AI的“密铺”作品做鉴定：指出它用了哪些变换、哪里其实不合法，并和AI一起造出一幅"
     "真正无缝、不重叠的镶嵌作品。",'')])

H(doc,"四、数学内核与关键结论（“裁判”的标准）")
para(doc); runs(doc.paragraphs[-1],[
    ("① 变换语言：平移、旋转、轴对称（及滑动反射）是纹样生成的基本动作；单位纹样经变换连续排布成纹样。"
     "② 密铺判据：图形能无缝隙、不重叠地铺满平面，其充要的直观条件是——",''),
    ("拼接点处各角之和恰为360°",'b'),
    ("。据此可判正多边形能否单独密铺（见表1）。这一“角和=360°”正是审辩AI作品的硬标准。",'')])
# 表1 正多边形密铺
tb=doc.add_table(rows=5,cols=4); tb.alignment=WD_TABLE_ALIGNMENT.CENTER; tb.style='Table Grid'
d1=[["正多边形","一个内角","能否恰好拼成360°","能否单独密铺"],
    ["正三角形","60°","60°×6=360°","能"],
    ["正方形","90°","90°×4=360°","能"],
    ["正六边形","120°","120°×3=360°","能"],
    ["正五边形","108°","108°×3=324°，×4=432°","不能"]]
for i,row in enumerate(d1):
    for j,v in enumerate(row):
        c=tb.cell(i,j); c.text=""; pp=c.paragraphs[0]; pp.alignment=WD_ALIGN_PARAGRAPH.CENTER
        rr=pp.add_run(v); setf(rr,HEI if i==0 else SONG,9,bold=(i==0))
cp=doc.add_paragraph(); cp.alignment=WD_ALIGN_PARAGRAPH.CENTER; cp.paragraph_format.space_before=Pt(2); cp.paragraph_format.space_after=Pt(6)
r=cp.add_run("表1　正多边形能否单独密铺的角度判据（学生“当裁判”的依据）"); setf(r,KAI,9)
print("上半就绪")


# ================= 文档 =================
doc=Document(); sec=doc.sections[0]
sec.page_width=Cm(21.0); sec.page_height=Cm(29.7)
sec.top_margin=Cm(2.4); sec.bottom_margin=Cm(2.4); sec.left_margin=Cm(2.6); sec.right_margin=Cm(2.6)
nm=doc.styles['Normal']; nm.font.name=SONG; nm.font.size=Pt(10.5); nm._element.rPr.rFonts.set(qn('w:eastAsia'),SONG)

p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(4)
r=p.add_run("当AI画错了镶嵌——用数学识破生成式AI的“伪密铺”"); setf(r,HEI,15.5,bold=True)
p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(4)
r=p.add_run("——一节由AI“伪密铺”驱动、指向几何推理与理性质疑的综合与实践（以《从传统连续纹样到现代镶嵌图案》为例）"); setf(r,HEI,11.5,bold=True)
p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.space_after=Pt(8)
r=p.add_run("〔作者姓名〕（上海外国语大学云间中学，上海　201600）　沪教版七年级上册·综合与实践·约6课时（课内4＋课外2）"); setf(r,KAI,9)

H(doc,"一、立意：AI的“伪密铺”，逼出一堂更深的数学推理课")
para(doc); runs(doc.paragraphs[-1],[
    ("本设计先做一次“反过来”的自省：镶嵌这节课若只把AI当作找图、生成图案的工具，那么拿掉AI、"
     "用教材原有的图片与剪纸照样能上——事实上，教材的空中课堂原版正是全程无AI完成的。"
     "这说明“AI帮我们学镶嵌”只是锦上添花，不构成真正的创新。",''),
])
para(doc); runs(doc.paragraphs[-1],[
    ("因此本设计换一个关系：",''),
    ("以“审辩AI的伪密铺”为非它不可的真实问题情境，发展学生的几何直观、空间观念、推理能力与创新意识；"
     "“学会审辩AI”是这堂数学课自然生长出的、有时代价值的迁移收获，而非数学要去服务的终点。",'b'),
    ("生成式AI靠视觉统计“画”、而非几何规则“算”，能生成“以假乱真却数学非法”的镶嵌图，并理直气壮地宣称无缝——"
     "这个“看着对、其实错”的情境，恰好逼出一堂更深的数学推理课：学生用密铺的数学标准"
     "（拼接点各角之和为360°、无缝不重叠）判定、说理、修正。数学在此握有AI无法自证的判定权。"
     "全程依托上海市中小学“三个助手”平台组织与留痕。",''),
])
H(doc,"二、创新点与“非AI不可”的论证（反测）")
para(doc); runs(doc.paragraphs[-1],[
    ("为什么这一次AI“非它不可”?因为课的目标本身就是“审辩AI的生成物”——抽掉AI，这件事就不存在了。"
     "老师手画的错图，学生知道是“故意考我”；而AI是真不懂、真自信地错，且是学生将来会天天使用的工具。"
     "教他们“不信AI的眼、只信数学的理”，是AI时代的真实素养，只有借真实的AI才能训练。"
     "镶嵌，是这把尺子里最锋利的一把：密铺合法性是硬邦邦的360°，没有模糊地带，最适合戳穿“看着对、其实错”。",''),
])
para(doc); runs(doc.paragraphs[-1],[
    ("一句话创新点：",''),
    ("在AI时代，用一道数学硬标准（密铺=360°）训练学生识破生成式AI“视觉合理却数学非法”的输出——"
     "这是数学独有的、AI无法自证的判定权。",'b'),
])

H(doc,"三、课标与素养定位（诚实厘清两个层次）")
para(doc); runs(doc.paragraphs[-1],[
    ("领域为“综合与实践”，以项目式学习展开，关联“图形的运动”（平移、旋转、轴对称）与密铺",''),("[2]",'sup'),
    ("。数学核心素养落点：",'b'),("几何直观、空间观念、推理能力、创新意识（含质疑）",'b'),
    ("。须厘清两个层次以免拔高：其一，课堂上真正发展的是数学素养——“用360°规则判定真假”是推理能力，"
     "“质疑并发现AI的破绽”是创新意识，二者都是课标明列的表现；其二，“学会审辩AI生成内容”是这些数学素养"
     "在AI时代的迁移价值（媒介素养），并不另算作课标数学素养。故本课以数学素养为目的、以审辩AI情境为触发，"
     "前者是本体，后者是收获，二者不混淆",''),("[1]",'sup'),("。",''),
])

H(doc,"四、驱动问题")
para(doc); runs(doc.paragraphs[-1],[
    ("“AI画的镶嵌，能信吗?”——面对AI生成的、看上去天衣无缝的“镶嵌”图，我们能不能用数学把它的破绽找出来、说清它错在哪、再造出真正合法的镶嵌?",'b'),
])

H(doc,"五、数学内核：镶嵌（密铺）的数学判定")
para(doc); runs(doc.paragraphs[-1],[
    ("镶嵌合法的两条数学标准：①无缝隙、不重叠地铺满平面；②在每个拼接点，围绕该点的各内角之和恰为360°。"
     "据此可判定正多边形能否单独密铺（见表1），这也是学生审辩AI作品的“数学尺子”。",''),
])
t=doc.add_table(rows=6,cols=4); t.alignment=WD_TABLE_ALIGNMENT.CENTER; t.style='Table Grid'
d1=[["正多边形","每个内角","能否单独密铺","数学判定（拼接点角和=360°）"],
    ["正三角形","60°","能","60°×6 = 360°"],
    ["正方形","90°","能","90°×4 = 360°"],
    ["正五边形","108°","不能","108°×3=324°，×4=432°，无整数倍等于360°"],
    ["正六边形","120°","能","120°×3 = 360°"],
    ["正八边形","135°","不能（须与正方形组合）","135°×2 + 90° = 360°（半正密铺）"]]
for i,row in enumerate(d1):
    for j,v in enumerate(row):
        c=t.cell(i,j); c.text=""; pp=c.paragraphs[0]
        pp.alignment=WD_ALIGN_PARAGRAPH.CENTER if j!=3 else WD_ALIGN_PARAGRAPH.LEFT
        rr=pp.add_run(v); setf(rr,HEI if i==0 else SONG,9,bold=(i==0))
cp=doc.add_paragraph(); cp.alignment=WD_ALIGN_PARAGRAPH.CENTER; cp.paragraph_format.space_before=Pt(2); cp.paragraph_format.space_after=Pt(6)
r=cp.add_run("表1　正多边形能否单独密铺的数学判定（学生审辩AI作品的依据）"); setf(r,KAI,9)
print("上半就绪")


H(doc,"六、项目流程（六环节，见图1）")
pic=doc.add_paragraph(); pic.alignment=WD_ALIGN_PARAGRAPH.CENTER; pic.paragraph_format.space_after=Pt(2)
pic.add_run().add_picture(FIG1,width=Cm(15.8))
cp=doc.add_paragraph(); cp.alignment=WD_ALIGN_PARAGRAPH.CENTER; cp.paragraph_format.space_after=Pt(6)
r=cp.add_run("图1　镶嵌审辩课六环节路径：AI“伪密铺”为情境、数学判定为主线（依托教材活动串）"); setf(r,KAI,9)

def envp(doc,name,segs):
    p=para(doc,ind=2,after=2)
    r=p.add_run(name); setf(r,HEI,10.5,bold=True)
    for tx,st in segs:
        rr=p.add_run(tx); setf(rr,SONG,10.5)
        if st=='b': rr.font.bold=True

envp(doc,"环节①　赏纹样（1课时）。",[
    ("承原课件：由越王勾践剑菱形纹引入连续纹样，赏唐草纹（二方连续）、清代福寿纹葫芦盒（四方连续）。",''),
    ("（AI在此仅为辅助，可用教材图片替代，非必需。）",''),
])
envp(doc,"环节②　辨变换（1课时）。",[
    ("承教材活动1、2：识别二方/四方连续纹样的单位纹样及其平移、旋转、轴对称等变换方式。",''),
    ("小试审辩：",'b'),("上传一个纹样让AI判定“单位纹样与所用变换”；AI常把旋转误作平移、或找错单位纹样，"
     "学生用变换的定义纠正它。教学助手推同类互动题并统计。（此处的审辩用AI误判或同伴错答均可，AI非必需，是环节⑤的热身。）",''),
])
envp(doc,"环节③　还原埃舍尔（2课内）。",[
    ("承教材活动3：将埃舍尔鸽子瓷砖的不规则单位图案，经平移、轴对称、旋转还原为基本几何图形，再四方连续密铺。",''),
    ("由此自然引出“到底怎样才算真正密铺”，过渡到环节④的数学判定。",''),
])
envp(doc,"环节④　提猜想·学裁判（0.5课时）。",[
    ("先不给结论，抛问让学生提出猜想：",''),
    ("“哪些正多边形能单独密铺?”",'b'),
    ("学生凭直觉先猜（如“三角形、正方形能，五边形大概不行”），经教学助手提交猜想并统计分布；"
     "再引导用“拼接点各角之和须为360°、无缝隙不重叠”这把数学尺子逐一验证、证实或证伪（表1）。"
     "学生由此经历“提出猜想—数学验证”的完整过程，这把尺子也成为下一环节审辩AI的依据——"
     "没有它，审辩只会流于“看着像不像”。",''),
])
envp(doc,"环节⑤　审AI作品（★全课“非AI不可”的核心，1.5课时）。",[
    ("教师课前用备课助手让AI生成若干“镶嵌／四方连续”图案（含缺陷者），投屏共审：",''),
    ("这幅用了哪些变换?是不是真的无缝、不重叠?拼接点角和够360°吗?",'b'),
    ("学生逐幅鉴定、指出AI错在哪、说明修正方向；每人的判断经教学助手作答留痕、统计（全员被看见）。"
     "本环节是全课唯一“非AI不可”处——审辩对象必须是真实AI的生成物：它“真不懂却自信地错”，"
     "这种认知冲突是教师手画的错图替代不了的，而这恰是学生将来面对AI时最该有的警觉。",''),
])
envp(doc,"环节⑥　共创作品（2课外）。",[
    ("小组把一个正多边形“变形”为鸽子/动物等单位图案，再用平移等变换四方连续密铺，并用数学验证其确实合法；"
     "成果经作业辅导助手提交、同伴互评与展示。（变形可用AI，也可沿用教材剪纸法，AI非必需。）",''),
])

H(doc,"七、“三个助手”的分工")
t=doc.add_table(rows=4,cols=3); t.alignment=WD_TABLE_ALIGNMENT.CENTER; t.style='Table Grid'
d2=[["阶段","助手","主要任务"],
    ["课前","备课助手","生成传统纹样素材库；用引导语让AI生成“待审的镶嵌图”（含缺陷）；准备分层任务单"],
    ["课中","教学助手","推送“识别变换”互动题并实时统计；白板标注；投屏共审AI作品，收集每人的鉴定判断"],
    ["课后","作业辅导助手","接收小组镶嵌作品与设计说明；组织同伴互评（点赞、评论）与个性化反馈"]]
for i,row in enumerate(d2):
    for j,v in enumerate(row):
        c=t.cell(i,j); c.text=""; pp=c.paragraphs[0]
        pp.alignment=WD_ALIGN_PARAGRAPH.CENTER if j!=2 else WD_ALIGN_PARAGRAPH.LEFT
        rr=pp.add_run(v); setf(rr,HEI if i==0 else SONG,9,bold=(i==0))
cp=doc.add_paragraph(); cp.alignment=WD_ALIGN_PARAGRAPH.CENTER; cp.paragraph_format.space_before=Pt(2); cp.paragraph_format.space_after=Pt(6)
r=cp.add_run("表2　“三个助手”在项目各阶段的分工"); setf(r,KAI,9)
print("中段就绪")


H(doc,"八、AI审辩任务示例（环节⑤用）")
for seg in [
    ("例1（正五边形密铺）。",["AI生成或声称“用正五边形可以密铺”。学生审辩：正五边形每个内角108°，"
        "108°×3=324°<360°，108°×4=432°>360°，找不到整数个恰好拼满360°，故正五边形不能单独密铺——AI错。"]),
    ("例2（看着无缝，实则有缝/重叠）。",["AI生成的图“远看像密铺”。学生放大检查拼接处：若出现缝隙或重叠，"
        "即不满足“无缝隙、不重叠”，判为不合法，并指出问题出在哪个拼接点。"]),
    ("例3（混淆变换方式）。",["AI把某纹样的构成说成“平移”。学生依变换定义审辩：若对应图案有翻转，"
        "应是轴对称；若绕某点转动，应是旋转——纠正AI的识别。"]),
]:
    p=para(doc,ind=2,after=2); r=p.add_run(seg[0]); setf(r,HEI,10.5,bold=True)
    rr=p.add_run(seg[1][0]); setf(rr,SONG,10.5)

H(doc,"九、物化成果")
para(doc); runs(doc.paragraphs[-1],[
    ("个人：纹样“单位纹样＋变换方式”识别单、AI镶嵌作品《鉴定报告》（指出合法性、错在哪、如何改）。"
     "小组：一幅合法的镶嵌作品＋《设计说明》（写清单位纹样、所用变换、拼接点角和为360°的验证）。",''),
])

H(doc,"十、过程性评价量规（维度×证据来源×权重）")
para(doc); runs(doc.paragraphs[-1],[
    ("量规对齐教材建议的三方面评价——个人“学习任务完成情况”“活动体验与习得”、小组“成员互评”",''),("[2]",'sup'),
    ("，并借“三个助手”把证据落到实处。",''),
])
t=doc.add_table(rows=6,cols=3); t.alignment=WD_TABLE_ALIGNMENT.CENTER; t.style='Table Grid'
d3=[["评价维度（对应素养）","权重","证据来源（成果／助手）"],
    ["密铺合法性的数学判定（空间观念·推理）","30%","AI作品《鉴定报告》中的角度计算与判定"],
    ["审辩AI：识破伪密铺、指错与修正（核心·批判质疑）","30%","《鉴定报告》、课堂共审发言 → 教学助手统计"],
    ["识别单位纹样与变换（几何直观）","15%","环节②互动作答 → 教学助手统计、识别单"],
    ["镶嵌作品创作与设计说明（创新意识）","15%","小组作品与《设计说明》 → 作业辅导助手"],
    ["协作与互评（交流表达）","10%","讨论区发言、同伴互评"]]
for i,row in enumerate(d3):
    for j,v in enumerate(row):
        c=t.cell(i,j); c.text=""; pp=c.paragraphs[0]
        pp.alignment=WD_ALIGN_PARAGRAPH.CENTER if j!=2 else WD_ALIGN_PARAGRAPH.LEFT
        rr=pp.add_run(v); setf(rr,HEI if i==0 else SONG,9,bold=(i==0))
cp=doc.add_paragraph(); cp.alignment=WD_ALIGN_PARAGRAPH.CENTER; cp.paragraph_format.space_before=Pt(2); cp.paragraph_format.space_after=Pt(6)
r=cp.add_run("表3　镶嵌项目过程性评价量规（维度×证据来源×权重，合计100%）"); setf(r,KAI,9)

H(doc,"十一、实施建议与边界")
para(doc); runs(doc.paragraphs[-1],[
    ("其一，AI生成的图案须教师课前预览、遴选，确保“待审图”既有缺陷可辨、又不误导；"
     "其二，一切合法性以数学判定为准，防止课堂滑向“只比谁的图好看”；"
     "其三，AI擅长出图、不擅长严谨判断，切勿让AI来“判对错”，判定权始终在数学与学生手里；"
     "其四，课时按教材建议课内4＋课外2，环节⑤为重心；"
     "其五，学生端操作在“三个助手”平台内完成，注意用网与数据规范；"
     "其六，本课有意把AI集中在环节⑤（审辩其生成物）这一“非它不可”处，①②⑥的AI均为辅助、可用教材图片或剪纸替代——"
     "如此“数智赋能”方落在实处而非装饰，也经得起“拿掉AI还能不能上”的反问。",''),
])

H(doc,"十二、模拟课堂实录片段（预演，非真实数据）")
para(doc); runs(doc.paragraphs[-1],[
    ("以下为设计阶段的模拟预演，用于打磨环节、预设AI的典型输出与学生反应；其中的对话与统计均为模拟，"
     "正式实施后须以真实课堂实录和平台数据替换。",''),
])

def dlg(doc,who,text,ai=False):
    p=para(doc,ind=1,after=1)
    r=p.add_run(who+"："); setf(r,HEI,10,bold=True,color=(RGBColor(0x1F,0x6F,0xB2) if ai else None))
    r=p.add_run(text); setf(r,SONG,10)
def stat(doc,text):
    p=para(doc,ind=1,after=2); r=p.add_run("〔模拟平台数据〕"); setf(r,KAI,9,color=RGBColor(0x99,0x66,0x00))
    r=p.add_run(text); setf(r,KAI,9,color=RGBColor(0x99,0x66,0x00))

p=para(doc,ind=2,after=1); r=p.add_run("片段1｜环节②　审辩AI对“变换方式”的识别"); setf(r,HEI,10.5,bold=True)
dlg(doc,"教师","（投屏一条“回纹”二方连续）先请AI说说它的单位纹样和变换方式。")
dlg(doc,"AI","这条纹样是把一个“回”字形单位纹样不断平移排列而成的。",ai=True)
dlg(doc,"教师","同意吗?各组在教学助手上作答“它用了哪些变换”。")
stat(doc,"平移 28人 ／ 平移＋旋转(180°) 15人 ／ 含轴对称 4人（多数与AI一样只看到平移）。")
dlg(doc,"学生1","我觉得不全是平移，相邻两个“回”是反着的，应该有180°旋转。")
dlg(doc,"教师","上来比一比。（拖单位纹样平移→不重合；旋转180°→重合）")
dlg(doc,"学生1","所以是“平移＋旋转”，AI只说对了一半。")
dlg(doc,"教师","AI帮我们起了个头，但对不对，得我们用变换的定义量一量。")

p=para(doc,ind=2,after=1); r=p.add_run("片段2｜环节④　先提猜想，再用数学验证"); setf(r,HEI,10.5,bold=True)
dlg(doc,"教师","先别急着算。你们猜，哪些正多边形能单独密铺?在教学助手上写下你的猜想。")
stat(doc,"正三角形能 41人／正方形能 45人／正五边形能 22人／正六边形能 26人（正五、正六分歧大，正是好问题）。")
dlg(doc,"教师","分歧最大的是正五边形和正六边形。怎么才能不靠猜、把它定下来?")
dlg(doc,"学生","看拼接点凑不凑得满360°!")
dlg(doc,"教师","对。这把“尺子”我们一起磨出来——各角之和恰好360°、还要无缝不重叠。拿它去量表1里的每一个。")
dlg(doc,"学生","那我刚才猜正五边形能，是错的，得改。")
dlg(doc,"教师","好——猜想要经数学检验才算数。这把尺子，等下就用来量AI画的图。")

p=para(doc,ind=2,after=1); r=p.add_run("片段3｜环节⑤　审AI作品：正五边形能密铺吗（核心）"); setf(r,HEI,10.5,bold=True)
dlg(doc,"教师","我让AI“画一幅用正五边形密铺的图案”，它交卷了。")
dlg(doc,"AI","（给出一张正五边形铺展图）这是用正五边形密铺的图案，它们无缝拼接。",ai=True)
dlg(doc,"教师","真的无缝吗?先别信，拿数学量。正五边形一个内角多少?")
dlg(doc,"学生","(5−2)×180°÷5 = 108°。")
dlg(doc,"教师","拼接点要凑满多少?")
dlg(doc,"学生2","360°。可是108×3=324不到，108×4=432又超了，没有整数个刚好360!")
dlg(doc,"教师","所以?")
dlg(doc,"学生（齐）","正五边形不能单独密铺!AI那张图，拼接处一定有缝或者重叠。")
dlg(doc,"教师","放大看——（投屏放大）每个拼接点果然留着一道缝。AI会画，但画的是“假密铺”。")
stat(doc,"“这幅AI图合法吗”：开始判“不合法”19人；经角度计算＋放大验证后判“不合法”44人，3人待助。")
dlg(doc,"学生3","正五边形不行，换正六边形，120×3=360，能铺。")
dlg(doc,"教师","这就是数学当裁判——AI造图，我们判它合不合法。")

p=para(doc,ind=2,after=1); r=p.add_run("片段4｜环节⑥　共创：AI变形＋数学验证"); setf(r,HEI,10.5,bold=True)
dlg(doc,"学生组","我们让AI把正六边形“变”成一条小鱼当单位图案。")
dlg(doc,"教师","变形后还能密铺吗?凭什么?")
dlg(doc,"学生","变形只是把边“凹进去一块、凸出去同样一块”，拼接角度没变，拼接点还是三个120°=360°，照样无缝密铺。")
dlg(doc,"教师","漂亮——AI负责把它变好看，你们负责证明它还合法。作品和这段说明传到作业辅导助手，互相评一评。")

p=para(doc,ind=2,after=1); r=p.add_run("片段5｜收尾提炼"); setf(r,HEI,10.5,bold=True)
dlg(doc,"教师","今天AI靠得住的是什么?靠不住的是什么?")
dlg(doc,"学生","靠得住画图；靠不住“判对错”——密铺合不合法，得我们用360°去量。")
dlg(doc,"教师","那今天我们真正学到的，只是镶嵌吗?")
dlg(doc,"学生","是学会了用数学去识破AI——它画得再像，凑不满360°就是假的。")
dlg(doc,"教师","对。以后AI给你任何“看着对”的东西，别急着信，用数学这把尺子先量一量。")

p=para(doc,ind=2,after=1); r=p.add_run("预演反思（据模拟打磨设计）"); setf(r,HEI,10.5,bold=True)
para(doc); runs(doc.paragraphs[-1],[
    ("①环节②中多数学生和AI一样只看到“平移”，说明“审辩AI的识别”确有认知价值，值得做；"
     "②正五边形是绝佳锚点——计算简单、结论反直觉、可当场放大验证；"
     "③环节⑤易超时，AI待审作品宜控制在2~3幅；"
     "④教师课前须挑“缝隙／重叠看得见”的AI图，否则审辩落空；"
     "⑤全程务必守住“判定权在数学与学生”，AI不参与判对错；"
     "⑥“先猜后证”环节里正五、正六边形分歧最大，是最好的探究切入点，但约需8~10分钟，"
     "要与环节⑤的审辩时间合理分配，防止整节超时。",''),
])

H(doc,"参考文献")
for ref in [
    "[1]　中华人民共和国教育部.义务教育数学课程标准(2022年版)[M].北京:北京师范大学出版社,2022.",
    "[2]　义务教育教科书（五·四学制）数学·七年级上册及其教学参考资料[M].上海:上海教育出版社.（“综合与实践·从传统连续纹样到现代镶嵌图案”）",
]:
    p=doc.add_paragraph(); p.paragraph_format.line_spacing=1.35; p.paragraph_format.space_after=Pt(1)
    p.paragraph_format.left_indent=Pt(21); p.paragraph_format.first_line_indent=Pt(-21)
    r=p.add_run(ref); setf(r,SONG,9)
p=para(doc,ind=0,after=0,line=1.3); p.paragraph_format.space_before=Pt(6)
r=p.add_run("　说明：本设计依托教材“综合与实践”活动与其空中课堂课件，数据、截图与文献细节待实施后据实补充，不夸大AI作用；AI仅用于生成与初判，数学判定与结论由师生把关。")
setf(r,KAI,8.5,color=RGBColor(0x88,0x88,0x88))

doc.save(OUT); print("已保存:",OUT)
