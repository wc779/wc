# -*- coding: utf-8 -*-
"""生成可编辑的PowerPoint文件。
运行：pip install python-pptx && python generate_autophagy_ppt.py
输出：细胞自噬原理与研究进展.pptx
"""
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

OUT = "细胞自噬原理与研究进展.pptx"
BG = RGBColor(16,24,39); TEAL = RGBColor(125,226,209); GOLD = RGBColor(255,204,102); WHITE = RGBColor(238,244,255); MUTED = RGBColor(198,214,231)
prs = Presentation(); prs.slide_width= Inches(13.333); prs.slide_height= Inches(7.5)
blank=prs.slide_layouts[6]

def box(slide, x,y,w,h, text, size=24, color=WHITE, bold=False, align=PP_ALIGN.LEFT):
    shape=slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); tf=shape.text_frame; tf.word_wrap=True
    p=tf.paragraphs[0]; p.text=text; p.alignment=align
    for r in p.runs: r.font.name='Microsoft YaHei'; r.font.size=Pt(size); r.font.bold=bold; r.font.color.rgb=color
    return shape

def slide(title, bullets=None, kicker="CELL BIOLOGY SEMINAR", cards=None):
    s=prs.slides.add_slide(blank); bg=s.background.fill; bg.solid(); bg.fore_color.rgb=BG
    box(s,.8,.45,11.8,.35,kicker,11,GOLD,True); box(s,.8,.9,11.8,.75,title,30,TEAL,True)
    if cards:
        for idx,(head,body) in enumerate(cards):
            x=.9+idx*6.15; sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(2.2), Inches(5.5), Inches(3.5)); sh.fill.solid(); sh.fill.fore_color.rgb=RGBColor(32,62,87); sh.line.color.rgb=TEAL
            box(s,x+.25,2.48,5,0.45,head,20,TEAL,True); box(s,x+.25,3.1,5,2.2,body,19,WHITE)
    elif bullets:
        text='\n'.join('• '+b for b in bullets); box(s,1.0,2.1,11.2,4.7,text,22,WHITE)
    box(s,.8,7.1,8,.2,'细胞自噬的原理与研究进展',9,MUTED); box(s,12.1,7.1,.5,.2,str(len(prs.slides)),9,MUTED,False,PP_ALIGN.RIGHT)

# title
s=prs.slides.add_slide(blank); s.background.fill.solid(); s.background.fill.fore_color.rgb=BG
box(s,.9,1.2,11.5,.4,'CELL BIOLOGY SEMINAR',14,GOLD,True); box(s,.9,2.0,11.8,1.0,'细胞自噬的原理与研究进展',38,TEAL,True); box(s,.9,3.3,11.3,.8,'从细胞内清除与再循环，到疾病机制和治疗靶点',24,WHITE); box(s,.9,5.4,10,.4,'汇报人：________    日期：________',16,MUTED)
slide('研究背景与意义',['自噬维持细胞稳态、物质循环和应激适应。','营养缺乏、缺氧、氧化应激、感染和细胞器损伤均可调节自噬。','异常自噬与肿瘤、神经退行性疾病、感染、代谢紊乱和衰老相关。','关键问题：自噬增强是否代表通量真正顺畅？'])
slide('细胞自噬的概念与类型',cards=[('按过程分类','宏自噬：形成双层膜自噬体\n微自噬：溶酶体直接内陷吞入\nCMA：伴侣蛋白介导底物进入溶酶体'),('按选择性分类','非选择性自噬\n选择性自噬：线粒体自噬、病原体自噬、内质网自噬等')])
slide('自噬的核心分子机制',['营养缺乏/能量应激 → mTOR受抑、ULK1激活 → Beclin1–VPS34生成PI3P → 隔离膜形成与延伸。','核心网络：mTORC1、AMPK、ULK1/2、ATG13、FIP200、Beclin 1、VPS34、WIPI和DFCP1。'])
slide('自噬体形成、成熟与降解',['膜成核 → 膜延伸与LC3脂化 → 自噬体封闭 → 与溶酶体融合 → 底物降解和物质再循环。','ATG12–ATG5–ATG16L1、ATG7、ATG3促进膜延伸和LC3脂化；SNARE、Rab7、HOPS和LAMP参与融合。'])
slide('选择性自噬与线粒体自噬',cards=[('货物识别','p62/SQSTM1、NBR1、OPTN、NDP52等受体结合LC3并识别泛素化底物。'),('PINK1–Parkin通路','受损线粒体上PINK1积累，激活Parkin介导的泛素化，招募受体并促进线粒体清除。')])
slide('常用实验方法与指标',['Western blot：LC3-I/II、p62、ATG蛋白。','免疫荧光/共聚焦：LC3 puncta和共定位。','mRFP–GFP–LC3：区分自噬体和自噬溶酶体。','TEM：观察双层膜自噬体。','原则：单一指标不能证明自噬流增强。'])
slide('自噬流：判断功能状态的关键',['LC3-II增加可能来自自噬体形成增加，也可能来自溶酶体降解受阻。','结合BafA1或CQ、LC3-II、p62、双荧光报告和细胞活性进行交叉验证。','自噬流反映自噬体形成、融合和降解的动态总过程。'])
slide('自噬与肿瘤：双重作用',cards=[('肿瘤早期','清除受损线粒体和蛋白质，降低氧化应激与基因组不稳定性，可能抑癌。'),('肿瘤进展期','适应缺氧、营养限制和治疗压力，促进肿瘤细胞存活、转移和耐药。')])
slide('自噬与神经退行性疾病',['神经元依赖高效的蛋白质质量控制和线粒体质量控制。','自噬流受阻可导致蛋白聚集体、受损线粒体和炎症信号积累。','帕金森病、阿尔茨海默病等疾病涉及自噬体运输、溶酶体酸化和底物降解异常。'])
slide('自噬与感染、免疫',cards=[('宿主防御','病原体自噬清除细菌、病毒或寄生虫；自噬还影响抗原加工、炎症和免疫细胞功能。'),('病原体利用','部分病原体可劫持或抑制自噬体形成、融合和溶酶体降解，逃避免疫清除。')])
slide('自噬、衰老与代谢',['衰老伴随自噬功能下降、受损蛋白和异常线粒体积累。','自噬参与脂滴分解、氨基酸回收、能量稳态和代谢重编程。','雷帕霉素、热量限制、间歇性禁食和亚精胺等方向受到关注，但长期安全性仍需评估。'])
slide('研究进展：从基础机制到临床转化',['研究对象扩展至选择性自噬、线粒体自噬和溶酶体稳态。','研究手段转向动态成像、单细胞组学、空间组学和功能性通量检测。','药物开发涵盖ULK、VPS34、ATG4、PPT1和溶酶体功能等靶点。','联合治疗和生物标志物分层是提高转化成功率的重要方向。'])
slide('挑战与展望',['不同细胞类型、疾病阶段和亚细胞区室中的自噬作用可能相反。','必须区分“自噬体数量增加”和“自噬流增强”。','药物的靶点选择性、组织分布、毒性和耐药问题仍待解决。','未来重点：时空精准调控、患者分层、可重复生物标志物与组合疗法。'])
slide('总结',['自噬是连接细胞稳态、代谢、免疫和疾病的重要枢纽。','mTOR–ULK1–Beclin1–VPS34–LC3轴构成宏自噬核心机制。','自噬研究必须重视自噬流、选择性和细胞情境。','靶向自噬具有治疗潜力，但精准调控和临床分层是关键。'])
slide('主要参考文献',['Klionsky DJ, et al. Autophagy. 2021;17:1–382. doi:10.1080/15548627.2020.1797280','Mizushima N, Komatsu M. Cell. 2011;147:728–741. doi:10.1016/j.cell.2011.10.026','Levine B, Kroemer G. Cell. 2019;176:11–42. doi:10.1016/j.cell.2018.09.048','Galluzzi L, et al. Nat Rev Drug Discov. 2017;16:487–511. doi:10.1038/nrd.2017.22','Vargas JNS, et al. Nat Rev Mol Cell Biol. 2022;23:167–185. doi:10.1038/s41580-022-00462-4','Nakatogawa H. Nat Rev Mol Cell Biol. 2020;21:439–458. doi:10.1038/s41580-020-0241-2'])
prs.save(OUT); print(OUT)
