#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
DAY 1, DAY 2, DAY 3 강의 교안 PPTX (16:9 와이드스크린, 강사 발표 노트 포함) 및 PDF 슬라이드 생성 스크립트
"""
import os
import subprocess
import tempfile
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

OUT_DIR = "/Users/jsoftware/Documents/강의/AI_vibe_coding/02_강의교안_PPTX"
os.makedirs(OUT_DIR, exist_ok=True)

# 컬러 팔레트
C_NAVY = RGBColor(15, 23, 42)       # #0F172A
C_SLATE = RGBColor(30, 41, 59)      # #1E293B
C_BLUE = RGBColor(37, 99, 235)      # #2563EB
C_LIGHT_BG = RGBColor(248, 250, 252)# #F8FAFC
C_CARD_BG = RGBColor(255, 255, 255) # #FFFFFF
C_BORDER = RGBColor(203, 213, 225)  # #CBD5E1
C_TEXT = RGBColor(30, 41, 59)       # #1E293B
C_MUTED = RGBColor(100, 116, 139)   # #64748B
C_ACCENT_BG = RGBColor(239, 246, 255) # #EFF6FF
C_WHITE = RGBColor(255, 255, 255)
C_CODE_BG = RGBColor(15, 23, 42)
C_GREEN = RGBColor(5, 150, 105)
C_ROSE = RGBColor(225, 29, 72)

FONT_KR = "맑은 고딕"
FONT_CODE = "Consolas"

def init_prs():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    return prs

def add_bg(slide, color=C_LIGHT_BG):
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = color
    bg.line.fill.background()
    return bg

def add_header(slide, badge_text, title_text, subtitle_text=""):
    # 상단 네이비 바
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(1.25))
    bar.fill.solid()
    bar.fill.fore_color.rgb = C_NAVY
    bar.line.fill.background()

    # 배지
    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(0.22), Inches(2.4), Inches(0.32))
    badge.fill.solid()
    badge.fill.fore_color.rgb = C_BLUE
    badge.line.fill.background()
    tf_b = badge.text_frame
    tf_b.word_wrap = True
    p_b = tf_b.paragraphs[0]
    p_b.text = badge_text
    p_b.font.name = FONT_KR
    p_b.font.size = Pt(11)
    p_b.font.bold = True
    p_b.font.color.rgb = C_WHITE
    p_b.alignment = PP_ALIGN.CENTER

    # 타이틀
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(0.55), Inches(12.0), Inches(0.6))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = title_text
    p.font.name = FONT_KR
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = C_WHITE

    if subtitle_text:
        tb2 = slide.shapes.add_textbox(Inches(3.15), Inches(0.22), Inches(9.5), Inches(0.32))
        tf2 = tb2.text_frame
        p2 = tf2.paragraphs[0]
        p2.text = subtitle_text
        p2.font.name = FONT_KR
        p2.font.size = Pt(11.5)
        p2.font.color.rgb = RGBColor(148, 163, 184)

def add_footer(slide, day_label, page_num, total_pages):
    tb = slide.shapes.add_textbox(Inches(0.6), Inches(7.05), Inches(12.1), Inches(0.35))
    tf = tb.text_frame
    p = tf.paragraphs[0]
    p.text = f"대기업 제조사 AX 핵심인재 양성 교육 (Claude Code 바이브코딩 PBL)  |  {day_label}  |  Slide {page_num} / {total_pages}"
    p.font.name = FONT_KR
    p.font.size = Pt(9.5)
    p.font.color.rgb = C_MUTED
    p.alignment = PP_ALIGN.RIGHT

def add_notes(slide, note_text):
    if note_text:
        notes_slide = slide.notes_slide
        notes_slide.notes_text_frame.text = note_text

def build_deck(deck_meta, slides_data, pptx_path, pdf_path):
    prs = init_prs()
    blank_layout = prs.slide_layouts[6]
    total = len(slides_data)

    for idx, s in enumerate(slides_data, 1):
        slide = prs.slides.add_slide(blank_layout)
        stype = s.get("type", "cards")

        if stype == "cover":
            add_bg(slide, C_NAVY)
            # 장식용 블루 라인
            accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.6), Inches(1.2), Inches(0.08))
            accent.fill.solid()
            accent.fill.fore_color.rgb = C_BLUE
            accent.line.fill.background()

            tb = slide.shapes.add_textbox(Inches(0.8), Inches(1.9), Inches(11.5), Inches(3.2))
            tf = tb.text_frame
            tf.word_wrap = True
            p0 = tf.paragraphs[0]
            p0.text = s["badge"]
            p0.font.name = FONT_KR
            p0.font.size = Pt(16)
            p0.font.bold = True
            p0.font.color.rgb = RGBColor(96, 165, 250)

            p1 = tf.add_paragraph()
            p1.text = s["title"]
            p1.font.name = FONT_KR
            p1.font.size = Pt(36)
            p1.font.bold = True
            p1.font.color.rgb = C_WHITE
            p1.space_before = Pt(12)

            p2 = tf.add_paragraph()
            p2.text = s["subtitle"]
            p2.font.name = FONT_KR
            p2.font.size = Pt(18)
            p2.font.color.rgb = RGBColor(203, 213, 225)
            p2.space_before = Pt(16)

            # 하단 요약 박스들
            boxes = s.get("meta_boxes", [])
            for b_i, b_info in enumerate(boxes):
                left = Inches(0.8 + b_i * 3.0)
                card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(5.1), Inches(2.8), Inches(1.45))
                card.fill.solid()
                card.fill.fore_color.rgb = C_SLATE
                card.line.color.rgb = RGBColor(71, 85, 105)
                ctf = card.text_frame
                ctf.word_wrap = True
                cp0 = ctf.paragraphs[0]
                cp0.text = b_info[0]
                cp0.font.name = FONT_KR
                cp0.font.size = Pt(11)
                cp0.font.bold = True
                cp0.font.color.rgb = RGBColor(96, 165, 250)
                cp1 = ctf.add_paragraph()
                cp1.text = b_info[1]
                cp1.font.name = FONT_KR
                cp1.font.size = Pt(13)
                cp1.font.bold = True
                cp1.font.color.rgb = C_WHITE
                cp1.space_before = Pt(6)

        elif stype == "cards":
            add_bg(slide, C_LIGHT_BG)
            add_header(slide, s["badge"], s["title"], s.get("subtitle", ""))
            cards = s.get("cards", [])
            n = len(cards)
            cols = 3 if n == 3 else (2 if n in (2, 4) else 1)
            rows = (n + cols - 1) // cols
            w = (12.1 - (cols - 1) * 0.3) / cols
            h = (4.85 - (rows - 1) * 0.25) / rows

            for c_i, c_data in enumerate(cards):
                r_idx = c_i // cols
                c_idx = c_i % cols
                left = Inches(0.6 + c_idx * (w + 0.3))
                top = Inches(1.5 + r_idx * (h + 0.25))

                card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, Inches(w), Inches(h))
                card.fill.solid()
                card.fill.fore_color.rgb = C_CARD_BG
                card.line.color.rgb = C_BORDER

                tb = slide.shapes.add_textbox(left + Inches(0.22), top + Inches(0.18), Inches(w - 0.44), Inches(h - 0.32))
                tf = tb.text_frame
                tf.word_wrap = True
                p0 = tf.paragraphs[0]
                p0.text = c_data["head"]
                p0.font.name = FONT_KR
                p0.font.size = Pt(15)
                p0.font.bold = True
                p0.font.color.rgb = C_BLUE

                for b_line in c_data.get("bullets", []):
                    p_b = tf.add_paragraph()
                    p_b.text = "• " + b_line
                    p_b.font.name = FONT_KR
                    p_b.font.size = Pt(12)
                    p_b.font.color.rgb = C_TEXT
                    p_b.space_before = Pt(6)

            if s.get("bottom_takeaway"):
                bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(6.48), Inches(12.1), Inches(0.48))
                bar.fill.solid()
                bar.fill.fore_color.rgb = C_ACCENT_BG
                bar.line.color.rgb = C_BLUE
                btf = bar.text_frame
                bp = btf.paragraphs[0]
                bp.text = "💡 핵심 포인트: " + s["bottom_takeaway"]
                bp.font.name = FONT_KR
                bp.font.size = Pt(11.5)
                bp.font.bold = True
                bp.font.color.rgb = C_NAVY
            add_footer(slide, deck_meta["day_label"], idx, total)

        elif stype == "prompt":
            add_bg(slide, C_LIGHT_BG)
            add_header(slide, s["badge"], s["title"], s.get("subtitle", ""))

            # 좌측 설명 카드
            left_card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(1.5), Inches(4.5), Inches(5.3))
            left_card.fill.solid()
            left_card.fill.fore_color.rgb = C_CARD_BG
            left_card.line.color.rgb = C_BORDER
            ltb = slide.shapes.add_textbox(Inches(0.8), Inches(1.7), Inches(4.1), Inches(4.9))
            ltf = ltb.text_frame
            ltf.word_wrap = True
            lp0 = ltf.paragraphs[0]
            lp0.text = s.get("left_title", "프롬프트 작성 핵심 포인트")
            lp0.font.name = FONT_KR
            lp0.font.size = Pt(15)
            lp0.font.bold = True
            lp0.font.color.rgb = C_BLUE
            for b_line in s.get("left_bullets", []):
                lp = ltf.add_paragraph()
                lp.text = "✔ " + b_line
                lp.font.name = FONT_KR
                lp.font.size = Pt(12)
                lp.font.color.rgb = C_TEXT
                lp.space_before = Pt(8)

            # 우측 코드/프롬프트 박스
            code_card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.35), Inches(1.5), Inches(7.35), Inches(5.3))
            code_card.fill.solid()
            code_card.fill.fore_color.rgb = C_CODE_BG
            code_card.line.color.rgb = C_SLATE
            ctb = slide.shapes.add_textbox(Inches(5.55), Inches(1.65), Inches(6.95), Inches(5.0))
            ctf = ctb.text_frame
            ctf.word_wrap = True
            cp0 = ctf.paragraphs[0]
            cp0.text = s.get("code_title", "📋 Claude Code 입력 프롬프트 예시")
            cp0.font.name = FONT_KR
            cp0.font.size = Pt(12.5)
            cp0.font.bold = True
            cp0.font.color.rgb = RGBColor(96, 165, 250)

            for line in s.get("code_lines", []):
                cp = ctf.add_paragraph()
                cp.text = line
                cp.font.name = FONT_KR
                cp.font.size = Pt(11)
                cp.font.color.rgb = RGBColor(241, 245, 249)
                cp.space_before = Pt(3)

            add_footer(slide, deck_meta["day_label"], idx, total)

        elif stype == "table":
            add_bg(slide, C_LIGHT_BG)
            add_header(slide, s["badge"], s["title"], s.get("subtitle", ""))
            headers = s["headers"]
            rows = s["rows"]
            t_shape = slide.shapes.add_table(len(rows) + 1, len(headers), Inches(0.6), Inches(1.55), Inches(12.1), Inches(5.1))
            table = t_shape.table
            for c_i, h_text in enumerate(headers):
                cell = table.cell(0, c_i)
                cell.fill.solid()
                cell.fill.fore_color.rgb = C_NAVY
                p = cell.text_frame.paragraphs[0]
                p.text = h_text
                p.font.name = FONT_KR
                p.font.size = Pt(11.5)
                p.font.bold = True
                p.font.color.rgb = C_WHITE

            for r_i, r_data in enumerate(rows, 1):
                for c_i, val in enumerate(r_data):
                    cell = table.cell(r_i, c_i)
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = C_CARD_BG if r_i % 2 == 1 else RGBColor(241, 245, 249)
                    p = cell.text_frame.paragraphs[0]
                    p.text = str(val)
                    p.font.name = FONT_KR
                    p.font.size = Pt(11)
                    p.font.bold = (c_i == 0)
                    p.font.color.rgb = C_TEXT

            add_footer(slide, deck_meta["day_label"], idx, total)

        add_notes(slide, s.get("notes", ""))

    prs.save(pptx_path)
    print("Saved PPTX:", pptx_path)

    # 동시에 동일한 내용의 16:9 와이드 PDF 교안 생성
    html_slides = []
    for idx, s in enumerate(slides_data, 1):
        stype = s.get("type", "cards")
        if stype == "cover":
            boxes_html = "".join(
                f'<div class="meta-box"><div class="meta-label">{b[0]}</div><div class="meta-val">{b[1]}</div></div>'
                for b in s.get("meta_boxes", [])
            )
            html_slides.append(f"""
            <div class="slide cover">
              <div class="badge">{s['badge']}</div>
              <h1>{s['title']}</h1>
              <p class="sub">{s['subtitle']}</p>
              <div class="meta-row">{boxes_html}</div>
            </div>
            """)
        elif stype == "cards":
            cards = s.get("cards", [])
            cols = 3 if len(cards) == 3 else 2
            cards_html = "".join(
                f'<div class="card"><h3>{c["head"]}</h3><ul>' +
                "".join(f'<li>{b}</li>' for b in c.get("bullets", [])) +
                '</ul></div>'
                for c in cards
            )
            takeaway_html = f'<div class="takeaway">💡 <strong>핵심 포인트:</strong> {s["bottom_takeaway"]}</div>' if s.get("bottom_takeaway") else ''
            html_slides.append(f"""
            <div class="slide">
              <div class="header">
                <span class="h-badge">{s['badge']}</span>
                <span class="h-sub">{s.get('subtitle','')}</span>
                <h2>{s['title']}</h2>
              </div>
              <div class="grid cols-{cols}">{cards_html}</div>
              {takeaway_html}
              <div class="footer">{deck_meta['day_label']} | Slide {idx} / {total}</div>
            </div>
            """)
        elif stype == "prompt":
            left_b = "".join(f'<li>✔ {b}</li>' for b in s.get("left_bullets", []))
            code_l = "\n".join(s.get("code_lines", []))
            html_slides.append(f"""
            <div class="slide">
              <div class="header">
                <span class="h-badge">{s['badge']}</span>
                <span class="h-sub">{s.get('subtitle','')}</span>
                <h2>{s['title']}</h2>
              </div>
              <div class="prompt-layout">
                <div class="card">
                  <h3>{s.get('left_title','핵심 가이드')}</h3>
                  <ul class="plain">{left_b}</ul>
                </div>
                <div class="code-panel">
                  <div class="code-head">{s.get('code_title','프롬프트 예시')}</div>
                  <pre>{code_l}</pre>
                </div>
              </div>
              <div class="footer">{deck_meta['day_label']} | Slide {idx} / {total}</div>
            </div>
            """)
        elif stype == "table":
            th_html = "".join(f'<th>{h}</th>' for h in s["headers"])
            tr_html = "".join(
                '<tr>' + "".join(f'<td>{v}</td>' for v in r) + '</tr>'
                for r in s["rows"]
            )
            html_slides.append(f"""
            <div class="slide">
              <div class="header">
                <span class="h-badge">{s['badge']}</span>
                <span class="h-sub">{s.get('subtitle','')}</span>
                <h2>{s['title']}</h2>
              </div>
              <div class="table-wrap">
                <table><thead><tr>{th_html}</tr></thead><tbody>{tr_html}</tbody></table>
              </div>
              <div class="footer">{deck_meta['day_label']} | Slide {idx} / {total}</div>
            </div>
            """)

    full_html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<style>
  @page {{ size: A4 landscape; margin: 0; }}
  * {{ box-sizing: border-box; }}
  body {{ margin: 0; padding: 0; font-family: 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif; background: #f8fafc; color: #1e293b; }}
  .slide {{
    width: 297mm; height: 209mm;
    page-break-after: always;
    position: relative;
    background: #f8fafc;
    overflow: hidden;
  }}
  .slide.cover {{
    background: #0f172a;
    color: #fff;
    padding: 28mm 22mm;
  }}
  .slide.cover .badge {{
    display: inline-block; background: #2563eb; color: #fff; padding: 6px 14px; border-radius: 20px; font-weight: 700; font-size: 13pt;
  }}
  .slide.cover h1 {{ font-size: 32pt; margin: 18px 0 12px 0; line-height: 1.25; }}
  .slide.cover .sub {{ font-size: 15pt; color: #cbd5e1; margin-bottom: 32px; }}
  .meta-row {{ display: table; width: 100%; table-layout: fixed; margin-top: 25mm; }}
  .meta-box {{ display: table-cell; background: #1e293b; border: 1px solid #334155; border-radius: 10px; padding: 14px 16px; margin-right: 12px; }}
  .meta-label {{ color: #60a5fa; font-size: 10pt; font-weight: 700; }}
  .meta-val {{ color: #fff; font-size: 12pt; font-weight: 700; margin-top: 6px; }}
  .header {{ background: #0f172a; color: #fff; padding: 14px 24px; height: 34mm; }}
  .h-badge {{ background: #2563eb; color: #fff; font-size: 9.5pt; font-weight: 700; padding: 3px 10px; border-radius: 12px; margin-right: 10px; }}
  .h-sub {{ color: #94a3b8; font-size: 10pt; }}
  .header h2 {{ margin: 8px 0 0 0; font-size: 19pt; }}
  .grid {{ padding: 16px 24px; }}
  .card {{ background: #fff; border: 1px solid #cbd5e1; border-radius: 10px; padding: 14px 18px; margin-bottom: 12px; }}
  .card h3 {{ margin: 0 0 8px 0; color: #1d4ed8; font-size: 13pt; }}
  .card ul {{ margin: 0; padding-left: 18px; }}
  .card li {{ font-size: 10.5pt; margin-bottom: 5px; line-height: 1.45; }}
  .card ul.plain {{ list-style: none; padding-left: 0; }}
  .prompt-layout {{ display: table; width: 100%; padding: 16px 24px; table-layout: fixed; }}
  .prompt-layout > div {{ display: table-cell; vertical-align: top; }}
  .prompt-layout .card {{ width: 38%; }}
  .code-panel {{ width: 60%; background: #0f172a; color: #f8fafc; border-radius: 10px; padding: 14px 18px; }}
  .code-head {{ color: #60a5fa; font-weight: 700; font-size: 11pt; margin-bottom: 8px; }}
  .code-panel pre {{ font-family: Menlo, Consolas, monospace; font-size: 9.5pt; white-space: pre-wrap; margin: 0; line-height: 1.42; color: #f1f5f9; }}
  .takeaway {{ position: absolute; bottom: 14mm; left: 24px; right: 24px; background: #eff6ff; border: 1px solid #3b82f6; border-radius: 8px; padding: 9px 14px; font-size: 10.5pt; color: #1e3a8a; }}
  .table-wrap {{ padding: 18px 24px; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 10pt; background: #fff; }}
  th {{ background: #0f172a; color: #fff; padding: 9px 10px; text-align: left; border: 1px solid #334155; }}
  td {{ padding: 8px 10px; border: 1px solid #cbd5e1; }}
  tr:nth-child(even) td {{ background: #f1f5f9; }}
  .footer {{ position: absolute; bottom: 4mm; right: 24px; font-size: 8.5pt; color: #64748b; }}
</style>
</head>
<body>
{''.join(html_slides)}
</body>
</html>
"""
    with tempfile.NamedTemporaryFile('w', suffix='.html', encoding='utf-8', delete=False) as tf:
        tf.write(full_html)
        t_html = tf.name
    try:
        subprocess.run([
            "/usr/local/bin/wkhtmltopdf",
            "--quiet",
            "--page-size", "A4",
            "--orientation", "Landscape",
            "--margin-top", "0mm",
            "--margin-bottom", "0mm",
            "--margin-left", "0mm",
            "--margin-right", "0mm",
            t_html,
            pdf_path
        ], check=True)
        print("Saved Slide PDF:", pdf_path)
    finally:
        if os.path.exists(t_html):
            os.remove(t_html)

# ==============================================================================
# DAY 1 슬라이드 데이터 (11/2 월, 13:00~18:00, 5H)
# ==============================================================================
day1_slides = [
    {
        "type": "cover",
        "badge": "DAY 1 · 2026. 11. 02 (월) 13:00–18:00 (5H)",
        "title": "바이브코딩 기본기 및 현업 MVP 설계",
        "subtitle": "현업 과제 구체화 → Claude Code 실전 워크플로우 → 첫 번째 동작하는 MVP(v1) 구현",
        "meta_boxes": [
            ("교육 대상", "전략기획·재무·마케팅·세일즈\n과장급 핵심인재 20명"),
            ("핵심 도구", "Claude Code · 로컬 웹앱\nGoogle Colab · CSV"),
            ("오늘의 목표", "핵심 기능 1~2개 확정 및\n첫 동작 MVP 화면 완성"),
            ("멘토링 방식", "5개 조 × 10분 순회 코칭\n과제 범위(Scope) 확정")
        ],
        "notes": "강사 오프닝: 1일차 사전 교육에서 완료한 환경 세팅과 기획서 초안을 바탕으로, 오늘부터 실제 작동하는 나만의 업무용 소프트웨어를 만듭니다."
    },
    {
        "type": "table",
        "badge": "DAY 1 타임테이블",
        "title": "오늘의 학습 일정 및 핵심 산출물 (13:00 ~ 18:00)",
        "subtitle": "이론은 압축하고 실습과 코칭으로 내 화면에 첫 MVP를 띄웁니다",
        "headers": ["시간", "세션명", "주요 학습 및 실습 활동", "진행 방식"],
        "rows": [
            ["13:00–13:30", "1. AX와 바이브코딩의 이해", "질문하는 AI vs 도구를 만드는 AI, 제조기업 직무별 MVP 사례", "강의"],
            ["13:30–14:10", "2. Claude Code 실전 활용", "CLAUDE.md 헌장, Plan 모드(Shift+Tab), 파일 생성·실행·복구 데모", "시연·실습"],
            ["14:10–15:00", "3. 과제 정의서 & MVP 범위 설정", "기획서 초안 다이어트, Input-Process-Output 정의, 핵심 기능 1~2개 확정", "워크숍"],
            ["15:00–15:10", "휴식 (Break)", "환경 점검 및 개별 질의응답", "휴식"],
            ["15:10–16:00", "4. 개발 프롬프트 엔지니어링", "역할·맥락·입출력·제약조건 구조화, 3단계 분할 지시 실습", "실습"],
            ["16:00–17:10", "5. 1차 MVP 프로토타입 구현", "화면 레이아웃 + 가명 샘플 CSV + 1번 핵심 기능 연결 및 실행", "PBL"],
            ["17:10–18:00", "6. 조별 순회 코칭 (조당 10분)", "5개 조 순회 점검, 초기 오류 해결, 최종 MVP 범위 확정", "코칭"]
        ]
    },
    {
        "type": "cards",
        "badge": "Session 1 · AX 패러다임",
        "title": "챗봇 활용(Chat)과 바이브코딩(Tool Building)의 결정적 차이",
        "subtitle": "매번 프롬프트를 다시 치는 대신, 클릭 한 번에 끝나는 '전용 업무 도구'를 만듭니다",
        "cards": [
            {
                "head": "기존 방식: 대화형 AI (ChatGPT/Claude 웹)",
                "bullets": [
                    "매월 마감 때마다 같은 프롬프트를 다시 입력하고 설명해야 함",
                    "대용량 엑셀/CSV 계산 시 할루시네이션(계산 실수) 발생 우려",
                    "나만 쓸 수 있고 팀원이나 후임자에게 업무 표준으로 인계하기 어려움"
                ]
            },
            {
                "head": "바이브코딩: AI와 함께 만드는 업무용 웹 도구 (MVP)",
                "bullets": [
                    "계산 수식과 규칙을 코드로 고정해 언제나 100% 동일한 정확도 보장",
                    "CSV 파일만 올리면 1초 만에 표·차트·요약 보고서가 동시에 생성됨",
                    "`index.html` 파일 하나로 부서 동료 누구나 설명 없이 즉시 사용 가능"
                ]
            }
        ],
        "bottom_takeaway": "비전공자인 우리가 가장 잘 아는 '현업의 업무 규칙과 페인포인트'가 최고의 설계도입니다."
    },
    {
        "type": "cards",
        "badge": "Session 1 · 직무별 사례",
        "title": "제조 대기업 4대 직무별 바이브코딩 MVP 성공 패턴",
        "subtitle": "백엔드 서버나 ERP 직접 연동 없이도 현업 시간을 80% 줄이는 로컬 웹 도구",
        "cards": [
            {
                "head": "① 전략기획: 전사 경영 KPI 브리핑 보드",
                "bullets": [
                    "입력: 사업부별 월간 실적 CSV (목표, 실적, 전월비)",
                    "핵심 기능: 달성률 자동 계산 + 신호등(정상/주의/위험) 판정",
                    "출력: 미달성 지표 중심 임원 보고용 3줄 요약 자동 생성"
                ]
            },
            {
                "head": "② 재무/회계: 예산 대비 실적(BVA) 이상 탐지기",
                "bullets": [
                    "입력: 부서별·계정별 예산 및 실집행 내역 CSV",
                    "핵심 기능: 편차 금액·소진율(%) 산출 및 100% 초과 항목 적색 강조",
                    "출력: 초과 집행 부서 필터링 및 소명 요청 대상 CSV 다운로드"
                ]
            },
            {
                "head": "③ 마케팅: B2B 캠페인 ROI & ROAS 비교기",
                "bullets": [
                    "입력: 채널별 광고비, 클릭수, 발굴 리드수, 수주전환 매출",
                    "핵심 기능: CTR, CPL(리드당 비용), ROAS(%) 자동 계산 및 순위 차트",
                    "출력: 저효율 캠페인 경보 및 차월 예산 재배분 시뮬레이션"
                ]
            },
            {
                "head": "④ 세일즈: 스마트 견적 산출 & 마진 검증기",
                "bullets": [
                    "입력: 제품 단가표 CSV + 수량·적용 할인율 슬라이더",
                    "핵심 기능: 할인 적용 공급가 및 예상 마진율(%) 실시간 계산",
                    "출력: 최소 승인 마진 미달 경고 + 고객 제출용 표준 견적서 인쇄"
                ]
            }
        ],
        "bottom_takeaway": "4가지 사례 모두 [CSV/폼 입력 → 자동 계산 및 이상 탐지 → 차트/보고서 출력]의 깔끔한 구조를 따릅니다."
    },
    {
        "type": "cards",
        "badge": "Session 2 · Claude Code",
        "title": "Claude Code 실전 활용 3대 핵심 무기",
        "subtitle": "비전공자가 길을 잃지 않고 안전하게 개발하는 작업 루틴",
        "cards": [
            {
                "head": "1. 프로젝트 헌장: CLAUDE.md",
                "bullets": [
                    "폴더 최상단에 위치하는 'AI 전용 업무 매뉴얼'",
                    "기술 제약(단일 HTML, 한글 인코딩, 서버 미사용)을 미리 못 박음",
                    "대화가 길어져도 AI가 프로젝트 규칙과 컬럼 구조를 잊지 않음"
                ]
            },
            {
                "head": "2. 설계부터 하는 습관: Plan 모드 (Shift+Tab)",
                "bullets": [
                    "곧바로 코드를 짜게 하면 엉뚱한 화면이 나올 확률이 높음",
                    "Plan 모드에서 먼저 '화면 구성과 3단계 개발 순서'를 제안받음",
                    "내가 이해할 수 있는 계획인지 확인 후 승인(Approve)하여 구현 시작"
                ]
            },
            {
                "head": "3. 작은 단위 반복: 단계별 구현 & 즉시 확인",
                "bullets": [
                    "1단계: 기본 화면 레이아웃 + 내장 샘플 데이터 표 출력",
                    "2단계: 핵심 계산 수식 + 필터/차트 연결",
                    "3단계: CSV 업로드/다운로드 및 요약 브리핑 다듬기"
                ]
            }
        ],
        "bottom_takeaway": "한 번에 완벽한 프로그램을 만들려 하지 말고, 매 단계 브라우저(F5)로 눈으로 확인하며 전진하세요."
    },
    {
        "type": "prompt",
        "badge": "Session 2 · 라이브 데모",
        "title": "[시연] 영업 실적 분석 대시보드 15분 만에 만들기",
        "subtitle": "강사 시연을 통해 Plan 모드 → CLAUDE.md 생성 → 1차 화면 실행 흐름을 함께 봅니다",
        "left_title": "시연 관전 포인트 4가지",
        "left_bullets": [
            "사내 실데이터 대신 '가상의 샘플 CSV 12행'을 먼저 만들게 지시",
            "로컬 더블클릭 시 빈 화면이 뜨지 않도록 기본 데이터를 내장 요청",
            "목표 대비 달성률 95% 미만 항목을 빨간색으로 자동 강조",
            "우측 하단에 임원 보고용 3줄 요약 문구 자동 생성 기능 탑재"
        ],
        "code_title": "📋 데모 프롬프트 (00_공통예제_영업실적_대시보드)",
        "code_lines": [
            "당신은 제조기업 영업관리 자동화 전문가이자 웹 개발자입니다.",
            "",
            "[목표] 서버 없이 index.html만 열면 작동하는 영업 실적 대시보드 제작",
            "[데이터] 연월, 지역본부, 제품군, 목표매출, 실적매출, 전년동월",
            "[핵심 기능]",
            "1. 페이지 오픈 시 기본 샘플 데이터로 KPI 카드 4개와 표·차트 즉시 표시",
            "2. 새로운 CSV 업로드(UTF-8/EUC-KR 자동 호환) 및 필터링 기능",
            "3. 달성률 95% 미만 항목 적색 경고 및 3줄 요약 브리핑 복사 버튼",
            "",
            "먼저 CLAUDE.md와 sales_performance_sample.csv를 만들고,",
            "단일 파일로 완결된 index.html을 작성해 주세요."
        ]
    },
    {
        "type": "cards",
        "badge": "Session 3 · MVP 범위 설정",
        "title": "14시간 안에 무조건 완주하는 'MVP 다이어트' 법칙",
        "subtitle": "1일차 기획서 초안에서 '거품'을 빼고 '핵심 가치 1~2개'만 남깁니다",
        "cards": [
            {
                "head": "✅ 반드시 포함할 것 (In-Scope · Must Have)",
                "bullets": [
                    "가명 처리된 CSV/Excel 파일 업로드 및 샘플 데이터 자동 로드",
                    "내 직무의 핵심 계산 로직 1~2개 (편차, 달성률, ROAS, 마진율 등)",
                    "이상치/점검 대상 자동 하이라이트 및 조건별 필터링",
                    "결과 요약 리포트 문구 생성, CSV 저장 또는 인쇄용 깔끔한 뷰"
                ]
            },
            {
                "head": "❌ 과감히 덜어낼 것 (Out-of-Scope · 현업 복귀 후 검토)",
                "bullets": [
                    "사내 ERP / SAP / MES 시스템과 실시간 API 직접 연결 (보안 승인 필요)",
                    "임직원 로그인/회원가입 및 부서별 권한 관리 시스템",
                    "복잡한 백엔드 DB 서버 구축 및 실시간 다중 사용자 동시 편집",
                    "5개 이상의 복잡한 메뉴 탭 (한 화면 대시보드에 집중!)"
                ]
            }
        ],
        "bottom_takeaway": "좋은 MVP는 '모든 기능이 있는 미완성품'이 아니라, '핵심 기능 1~2개가 완벽히 돌아가는 도구'입니다."
    },
    {
        "type": "prompt",
        "badge": "Session 4 · 프롬프트 설계",
        "title": "비전공자를 위한 5요소 구조화 프롬프트 공식",
        "subtitle": "역할 · 업무맥락 · 입출력 데이터 · 핵심기능 · 제약조건을 명확히 전달합니다",
        "left_title": "프롬프트 5대 필수 구성요소",
        "left_bullets": [
            "1. 역할(Role): 제조기업 OO직무 자동화 전문가 & 웹 개발자",
            "2. 맥락(Context): 누가, 어떤 반복 업무의 비효율을 줄이려 하는가",
            "3. 입출력(I/O): CSV의 정확한 컬럼명과 화면에 보여줄 요약 지표",
            "4. 핵심기능(Core): 계산 공식(예: 실집행/예산*100)과 색상 강조 기준",
            "5. 제약조건(Constraints): 서버 없이 로컬 실행, 한글 인코딩 호환"
        ],
        "code_title": "📋 실습 ① Plan 모드용 요구사항 정의 프롬프트",
        "code_lines": [
            "당신은 기업 업무 자동화 전문가이자 웹 개발자입니다.",
            "",
            "[업무 배경] 나는 제조기업의 재무 담당자입니다.",
            "매월 부서별 예산과 실제 집행 내역을 비교하는 데 5시간이 걸립니다.",
            "[입력 데이터] 연월, 부서명, 계정과목, 예산액, 실집행액, 전월집행액",
            "[핵심 기능]",
            "1. CSV 업로드 시 예산 대비 편차와 소진율(%) 자동 계산",
            "2. 소진율 100% 초과 항목 적색 표시 및 초과 항목만 필터링",
            "[제약조건] 서버/DB 없이 로컬 브라우저 실행, 가명 샘플 데이터 포함",
            "",
            "아직 코드는 작성하지 마세요. 개발 계획과 테스트 방법을 먼저 제안하세요."
        ]
    },
    {
        "type": "cards",
        "badge": "Session 5~6 · PBL & 코칭",
        "title": "1차 MVP 구현 실습 및 조별 순회 코칭 가이드 (16:00~18:00)",
        "subtitle": "오늘 저녁 6시 전까지 내 PC 브라우저에 첫 번째 동작 화면(v1)을 띄웁니다",
        "cards": [
            {
                "head": "Step 1. 과제 정의서 v1 & CLAUDE.md 작성 (20분)",
                "bullets": [
                    "배포된 [실습 템플릿 1]에 내 과제의 입력 컬럼과 핵심 기능 2개 기입",
                    "내 프로젝트 폴더 생성 후 `CLAUDE.md` 파일 저장"
                ]
            },
            {
                "head": "Step 2. 샘플 CSV 및 1차 화면 생성 (50분)",
                "bullets": [
                    "Claude Code에 가명 샘플 데이터 30행(`sample.csv`) 생성 지시",
                    "상단 요약 카드 4개 + 데이터 표가 나오는 `index.html` 1차 구현 및 실행"
                ]
            },
            {
                "head": "Step 3. 조별 순회 코칭 (5개 조 × 10분)",
                "bullets": [
                    "강사와 함께 현재 화면 점검 및 과도한 기능 범위 다이어트",
                    "초기 실행 오류 즉석 해결 및 다음 주(DAY 2)까지 해올 과제 확인"
                ]
            }
        ],
        "bottom_takeaway": "DAY 1 종료 체크: ① CLAUDE.md ② 가명 샘플 CSV ③ 브라우저에서 열리는 index.html v1 화면!"
    }
]

# ==============================================================================
# DAY 2 슬라이드 데이터 (11/9 월, 13:00~18:00, 5H)
# ==============================================================================
day2_slides = [
    {
        "type": "cover",
        "badge": "DAY 2 · 2026. 11. 09 (월) 13:00–18:00 (5H)",
        "title": "MVP 핵심 기능 완성 · AI 디버깅 및 실전 멘토링",
        "subtitle": "업무 데이터 연동 → 계산/시각화 고도화 → 에러 잡는 AI 디버깅 → 발표 시나리오 설계",
        "meta_boxes": [
            ("핵심 목표", "실제 업무 구조와 동일한\n가명 데이터로 핵심 기능 완성"),
            ("미니 강의 1", "비전공자를 위한 AI 디버깅 &\n테스트 체크리스트 작성법"),
            ("미니 강의 2", "5분 라이브 시연 발표 구성법\n(1분 문제 - 3분 시연 - 1분 효과)"),
            ("집중 코칭", "조별/개별 고난도 기술 이슈 해결\n및 업무 적용성 고도화")
        ]
    },
    {
        "type": "table",
        "badge": "DAY 2 타임테이블",
        "title": "오늘의 학습 일정 및 마일스톤 (13:00 ~ 18:00)",
        "subtitle": "오류 없이 매끄럽게 동작하는 MVP v2를 완성하고 금요일 발표 준비를 마칩니다",
        "headers": ["시간", "세션명", "주요 학습 및 실습 활동", "진행 방식"],
        "rows": [
            ["13:00–13:30", "1. 진척 점검 & 공통 이슈 리뷰", "주간 개발 현황 점검, 한글 깨짐·CORS·수식 NaN 등 공통 장애 해결 공유", "리뷰"],
            ["13:30–14:20", "2. 업무 데이터 연동 및 정제 실습", "현업 엑셀 헤더 기반 가명 CSV 고도화, 파일 업로드 및 결과 다운로드 구현", "실습"],
            ["14:20–15:10", "3. 직무별 핵심 비즈니스 로직 구현", "계산·검색·다중 필터링·Chart.js 시각화·보고서 요약 자동 생성", "PBL"],
            ["15:10–15:20", "휴식 (Break)", "휴식 및 동료 작품 둘러보기", "휴식"],
            ["15:20–16:10", "4. [미니강의] AI 디버깅·테스트 & 발표법", "콘솔 에러 해석, 최소 수정 지시법, 테스트 시나리오, 5분 시연 가이드", "미니강의·실습"],
            ["16:10–17:30", "5. 집중 개발 및 고난도 이슈 코칭", "개별 MVP 완성도 향상 및 조별 순회 기술 멘토링", "멘토링"],
            ["17:30–18:00", "6. 기능 검수 및 시연 시나리오 확정", "테스트 체크리스트 상호 검증, 발표용 데모 데이터 확정", "코칭"]
        ]
    },
    {
        "type": "cards",
        "badge": "Session 1~2 · 데이터 연동",
        "title": "현업 Excel/CSV 데이터 연동 시 가장 많이 겪는 3대 함정과 해법",
        "subtitle": "사내 보안은 100% 지키면서 현업 엑셀 파일이 막힘없이 읽히도록 다듬습니다",
        "cards": [
            {
                "head": "함정 1. Windows 엑셀 CSV 한글 깨짐 (CP949 vs UTF-8)",
                "bullets": [
                    "원인: 사내 Windows 엑셀에서 저장한 CSV는 주로 EUC-KR(CP949) 인코딩",
                    "해법: 업로드 시 UTF-8로 먼저 읽고 깨짐 문자 발견 시 EUC-KR로 자동 재독출",
                    "다운로드 시에는 `\\uFEFF`(UTF-8 BOM)를 붙여 엑셀에서 바로 열리게 처리"
                ]
            },
            {
                "head": "함정 2. 천단위 콤마(,)·빈칸으로 인한 NaN 계산 오류",
                "bullets": [
                    "원인: `'1,250,000'`이나 `' - '`, 빈 셀을 그대로 더하면 결과가 `NaN`으로 깨짐",
                    "해법: 모든 숫자 연산 전 콤마·통화기호·공백을 제거하는 `parseNum()` 함수 적용",
                    "분모가 0일 때 `Infinity` 대신 `0%`로 처리하는 방어 로직 추가"
                ]
            },
            {
                "head": "함정 3. 로컬 더블클릭 실행 시 외부 파일 로드 차단 (CORS)",
                "bullets": [
                    "원인: 브라우저 보안상 `file://` 경로에서는 외부 CSV 자동 fetch가 차단됨",
                    "해법: 기본 샘플 데이터 30행은 `index.html` 내부 변수로 내장해 즉시 표시",
                    "추가 데이터는 `<input type='file'>` 버튼으로 업로드하도록 이중 구성"
                ]
            }
        ],
        "bottom_takeaway": "배포해 드린 [공통오류 대응 FAQ]의 해결 프롬프트를 복사해 넣으면 Claude Code가 30초 만에 방어 코드를 추가합니다."
    },
    {
        "type": "cards",
        "badge": "Session 3 · 업무 가치 고도화",
        "title": "그냥 '표만 보여주는 화면'에서 '의사결정 도구'로 업그레이드하는 3가지 장치",
        "subtitle": "평가 기준 35점(핵심 기능) + 20점(사용성)을 동시에 끌어올리는 실전 팁",
        "cards": [
            {
                "head": "① 신호등 배지 & 조건부 하이라이트",
                "bullets": [
                    "모든 숫자를 검정 글씨로 나열하면 실무자가 다시 눈으로 찾아야 함",
                    "임계치(예: 예산 100% 초과, 달성률 90% 미만, 마진 20% 미달) 설정",
                    "문제 행은 연한 적색 배경 + `🚨 점검필요` 배지로 1초 만에 식별"
                ]
            },
            {
                "head": "② 다중 필터 & 차트 실시간 연동",
                "bullets": [
                    "부서/지역 필터 + '경고 항목만 보기' 필터를 조합",
                    "필터 변경 시 상단 KPI 카드 4개와 막대/선 차트, 하단 표가 동시에 변경",
                    "Chart.js 재렌더링 시 기존 차트 `.destroy()` 호출로 화면 깨짐 방지"
                ]
            },
            {
                "head": "③ 원클릭 결과물 출력 (Actionable Output)",
                "bullets": [
                    "분석에서 끝나지 않고 다음 업무 행동으로 바로 이어져야 진짜 도구",
                    "기능 A: '임원 보고용 3줄 요약 자동 생성 + 클립보드 복사' 버튼",
                    "기능 B: '필터링된 결과만 CSV 다운로드' 또는 '인쇄/PDF 클린 뷰'"
                ]
            }
        ],
        "bottom_takeaway": "수강생 여러분의 도메인 노하우는 '어떤 기준으로 경고를 띄우고 어떤 요약 문구를 만들지'에서 가장 빛납니다."
    },
    {
        "type": "prompt",
        "badge": "Session 4 · AI 디버깅 미니강의",
        "title": "비전공자를 위한 AI 디버깅 & 최소 범위 수정 지시법",
        "subtitle": "\"안 돼요, 고쳐주세요\" 대신 현상·원인·범위를 콕 집어 지시하면 기존 기능이 망가지지 않습니다",
        "left_title": "절대 실패하지 않는 디버깅 3원칙",
        "left_bullets": [
            "원칙 1. 브라우저에서 F12(개발자 도구) → Console 탭의 빨간 에러 문구 복사",
            "원칙 2. '내가 한 동작' vs '기대한 결과' vs '실제 현상'을 구분해 설명",
            "원칙 3. '기존에 잘 되는 OO 기능은 절대 건드리지 말고 최소 수정해줘' 명시",
            "원칙 4. 수정 직전에는 정상 동작하던 `index.html` 복사본(`index_backup.html`) 보관!"
        ],
        "code_title": "📋 실습 ② 오류 분석 및 최소 수정 프롬프트 템플릿",
        "code_lines": [
            "현재 프로젝트에서 발생한 오류를 분석하고 수정해 주세요.",
            "",
            "[발생 현상]",
            "- 내가 한 동작: 부서 필터를 변경한 뒤 '초과 항목만 보기'를 선택함",
            "- 기대 결과: 해당 부서의 초과 항목만 차트와 표에 표시되어야 함",
            "- 실제 현상: 차트가 이전 차트와 겹쳐서 깨지고 합계가 NaN으로 나옴",
            "",
            "[처리 순서]",
            "1. 오류 원인을 비전공자가 이해할 수 있게 쉽게 설명하세요.",
            "2. 다른 정상 기능에 영향이 없는 최소 수정안을 제시하고 적용하세요.",
            "3. 수정 내용과 내가 브라우저에서 테스트할 방법을 요약하세요."
        ]
    },
    {
        "type": "table",
        "badge": "Session 4 · 테스트 & 검증",
        "title": "MVP 완성을 위한 7대 필수 테스트 체크리스트",
        "subtitle": "오늘 퇴근 전 짝꿍과 서로의 화면을 바꿔가며 7개 항목을 모두 통과시키세요",
        "headers": ["번호", "점검 항목", "직접 수행할 테스트 시나리오", "통과 기준 (Pass Criteria)"],
        "rows": [
            ["1", "초기 로컬 실행", "index.html 파일을 더블클릭하여 브라우저 실행", "빈 화면 없이 기본 샘플 데이터와 대시보드 즉시 표시"],
            ["2", "CSV 파일 업로드", "sample.csv 업로드 (UTF-8 및 엑셀 저장 파일)", "한글 깨짐 없이 모든 행·열 데이터 정상 반영"],
            ["3", "핵심 수식 정확도", "요약 카드 합계·달성률·편차 숫자를 엑셀 값과 대조", "NaN/Infinity 없이 소수점·천단위 콤마 정확히 일치"],
            ["4", "필터·차트 연동", "드롭다운 필터 및 슬라이더 조건 변경", "KPI 카드, 차트, 테이블, 요약 문구가 동시에 업데이트"],
            ["5", "예외 데이터 방어", "콤마가 포함된 숫자, 빈 줄, 0값 데이터 입력", "화면이 멈추지 않고 안전하게 0 또는 정상 숫자로 처리"],
            ["6", "결과 출력 동작", "결과 CSV 저장, 브리핑 문구 복사, 인쇄 버튼 클릭", "파일 다운로드 및 클립보드 복사 정상 작동"],
            ["7", "사내 보안 준수", "화면 및 CSV 파일 내 실제 사내 기밀·실명 포함 여부", "100% 가명/합성 데이터로만 구성 확인"]
        ]
    },
    {
        "type": "cards",
        "badge": "Session 4 · 발표 가이드 안내",
        "title": "DAY 3 (11/13 금) 개인별 5분 시연 발표 준비 가이드",
        "subtitle": "인당 8분 배정 = 발표 및 라이브 시연 5분 + 주강사 전문가 피드백 3분",
        "cards": [
            {
                "head": "1분: [Why] 해결하려는 현업 문제 정의",
                "bullets": [
                    "우리 부서의 어떤 반복·비효율 업무를 해결하려 했는가?",
                    "기존 방식(수작업 엑셀 취합 등)은 매월 몇 시간이 걸렸고 어떤 실수가 있었는가?"
                ]
            },
            {
                "head": "3분: [How & Live Demo] 작동하는 MVP 라이브 시연",
                "bullets": [
                    "별도 PPT 대신 완성된 웹 MVP 화면을 바로 띄워놓고 시연!",
                    "샘플 CSV 업로드 → 핵심 기능 1(자동 계산/이상 탐지) → 핵심 기능 2(필터/차트/보고서 출력) 시연"
                ]
            },
            {
                "head": "1분: [Impact & Next] 업무 도입 효과 및 확장 계획",
                "bullets": [
                    "현업 도입 시 예상되는 시간 절감 효과 (예: 월 6시간 → 15분 단축)",
                    "교육 종료 후 부서 내 적용 계획 및 추가로 붙이고 싶은 2단계 기능"
                ]
            }
        ],
        "bottom_takeaway": "화려한 슬라이드 장표보다 '실제 내 업무 데이터를 넣었을 때 막힘없이 돌아가는 3분 시연'이 최고 점수를 받습니다."
    }
]

# ==============================================================================
# DAY 3 슬라이드 데이터 (11/13 금, 13:00~17:00, 4H)
# ==============================================================================
day3_slides = [
    {
        "type": "cover",
        "badge": "DAY 3 · 2026. 11. 13 (금) 13:00–17:00 (4H)",
        "title": "개인별 MVP 성과 발표 · 전문가 피드백 및 AX 확산 Wrap-up",
        "subtitle": "최종 실행 점검 → 20인 20색 라이브 MVP 시연 → 전문가 피드백 → 현업 적용 로드맵",
        "meta_boxes": [
            ("발표 인원", "AX 핵심인재 20명 전원\n1인 1과제 MVP 시연"),
            ("발표 시간", "인당 총 8분\n(시연 5분 + 피드백 3분)"),
            ("핵심 평가", "문제정의·핵심기능·사용성\n활용효과·확장성 (100점)"),
            ("최종 목표", "내 부서에 즉시 도입 가능한\n실전 AX 도구 확보")
        ]
    },
    {
        "type": "table",
        "badge": "DAY 3 타임테이블",
        "title": "오늘의 발표 및 Wrap-up 일정 (13:00 ~ 17:00)",
        "subtitle": "20명의 현업 과제가 실제 작동하는 도구로 탄생하는 시간입니다",
        "headers": ["시간", "세션명", "세부 진행 내용", "비고"],
        "rows": [
            ["13:00–13:30", "1. 최종 실행 점검 및 리허설", "브라우저 실행 상태, 데모용 CSV 파일, 화면 배율(100~125%) 최종 점검", "강사 순회 지원"],
            ["13:30–14:00", "2. 설득력 있는 MVP 시연법", "문제-해결-효과 중심 발표 핵심 팁 및 상호 학습 평가 안내", "미니 강의"],
            ["14:00–15:20", "3. 개인별 성과 발표 (Part 1)", "1번 ~ 10번 교육생 라이브 시연 발표 및 전문가 피드백 (10명 × 8분)", "1~3조"],
            ["15:20–15:30", "휴식 (Break)", "휴식 및 심사 점수 중간 집계", "10분"],
            ["15:30–16:40", "4. 개인별 성과 발표 (Part 2)", "11번 ~ 20번 교육생 라이브 시연 발표 및 전문가 피드백 (10명 × 7~8분)", "3~5조"],
            ["16:40–17:00", "5. 시상 · 총평 및 현업 적용 회고", "우수 프로젝트 선정, 보안·IT 협업 기반 현업 확산 로드맵 및 과정 회고", "Wrap-up"]
        ]
    },
    {
        "type": "table",
        "badge": "심사 및 피드백 기준",
        "title": "MVP 프로젝트 평가 기준표 (100점 만점)",
        "subtitle": "겉모습만 화려한 화면보다 '실제 현업 문제를 해결하는 견고한 동작'을 가장 높게 평가합니다",
        "headers": ["평가 영역", "배점", "핵심 평가 질문", "만점 포인트"],
        "rows": [
            ["1. 업무 문제 정의", "20점", "해결하려는 현업의 페인포인트가 명확하고 구체적인가?", "기존 소요 시간과 문제점이 선명하게 정의됨"],
            ["2. 핵심 기능 구현", "35점", "정의한 핵심 기능 1~2개가 실제로 오류 없이 동작하는가?", "샘플 데이터 업로드부터 계산·차트·출력까지 완벽 작동"],
            ["3. 사용성과 완성도", "20점", "비개발자 동료가 봐도 직관적이며 예외 처리가 되어 있는가?", "신호등 강조, 깔끔한 레이아웃, 에러 없는 안정성"],
            ["4. 업무 활용 효과", "15점", "현업 도입 시 시간 절감 및 휴먼 에러 감소 효과가 큰가?", "정량적 기대 효과(예: 월 5시간→10분)가 설득력 있음"],
            ["5. 발표 및 확장성", "10점", "문제-시연-효과가 명확하며 향후 확산 계획이 구체적인가?", "5분 타임박스 준수 및 현실적인 현업 정착 로드맵 제시"]
        ]
    },
    {
        "type": "cards",
        "badge": "Session 2 · 시연 노하우",
        "title": "발표 직전 최종 점검! 5분 라이브 시연을 빛내는 4가지 디테일",
        "subtitle": "발표자 여러분이 지난 2주간 만든 도구의 가치를 100% 전달하는 방법",
        "cards": [
            {
                "head": "1. 브라우저 확대 배율 110~125% 세팅",
                "bullets": [
                    "프로젝터 화면에서는 작은 표 글씨가 잘 보이지 않을 수 있음",
                    "`Ctrl` + `+` 키로 화면을 110~125%로 키워 숫자 가독성 확보"
                ]
            },
            {
                "head": "2. '비포(Before) 엑셀'과 '애프터(After) MVP' 대비",
                "bullets": [
                    "시연 시작 때 원본 샘플 CSV를 잠깐 보여준 뒤 업로드하면 임팩트 배가",
                    "\"이 복잡하던 50줄짜리 표가 업로드 1초 만에 이렇게 요약됩니다\""
                ]
            },
            {
                "head": "3. 코드 설명은 과감히 생략, 업무 언어에 집중",
                "bullets": [
                    "HTML이나 자바스크립트 문법을 설명할 필요는 전혀 없음",
                    "\"우리 팀이 매달 겪던 예산 초과 누락 문제를 이 빨간 배지가 막아줍니다\"로 설명"
                ]
            },
            {
                "head": "4. 동료의 발표에서 내 부서에 가져갈 아이디어 메모",
                "bullets": [
                    "타 직무(기획·재무·마케팅·영업) 동료가 만든 필터·요약 방식을 벤치마킹",
                    "교육 종료 후 전체 우수 MVP 템플릿을 서로 공유하여 활용 가능"
                ]
            }
        ],
        "bottom_takeaway": "자신감 있게 시연하세요! 비개발자가 직접 동작하는 업무 도구를 만들었다는 것 자체가 놀라운 AX 혁신입니다."
    },
    {
        "type": "cards",
        "badge": "Session 5 · 현업 적용 로드맵",
        "title": "교육장을 나서며: 내 MVP를 부서 표준 업무 도구로 정착시키는 3단계 로드맵",
        "subtitle": "오늘 완성한 프로토타입을 일회성 실습으로 끝내지 않고 실제 현업 성과로 연결하는 법",
        "cards": [
            {
                "head": "Stage 1. 개인 업무 즉시 투입 (1~2주 차)",
                "bullets": [
                    "사내 PC 로컬 폴더에 `index.html`을 두고 이번 달 마감/집계 업무에 실제 활용",
                    "실제 현업 엑셀의 컬럼 순서와 예외 케이스(결측치, 특수문자)를 반영해 미세 튜닝",
                    "기존 수작업 결과와 MVP 계산 결과를 교차 검산하여 신뢰도 100% 확보"
                ]
            },
            {
                "head": "Stage 2. 팀 내 공유 및 표준 양식화 (1개월 차)",
                "bullets": [
                    "팀 주간회의에서 3분 시연 후 팀원들에게 로컬 파일 패키지 배포",
                    "부서 내 파편화된 엑셀 취합 양식을 MVP 입력 표준 CSV 템플릿으로 통일",
                    "팀원들의 피드백을 받아 Claude Code로 편의 기능 1~2개 추가 업데이트"
                ]
            },
            {
                "head": "Stage 3. IT/보안 부서 협업 전사 시스템화 (3개월 차~)",
                "bullets": [
                    "팀 단위 사용 효과(월 절감 시간) 데이터를 근거로 IT/DX 부서에 정식 고도화 요청",
                    "화면과 비즈니스 로직이 이미 작동하는 MVP가 있으므로 개발 기간·비용 대폭 단축",
                    "사내 인트라넷 배포, SSO 인증 및 ERP DB 자동 연동으로 확장"
                ]
            }
        ],
        "bottom_takeaway": "AX 핵심인재의 진짜 역할은 '직접 모든 시스템을 구축하는 것'이 아니라 '작동하는 MVP로 문제를 정의하고 조직을 움직이는 것'입니다."
    }
]

if __name__ == "__main__":
    build_deck(
        {"day_label": "DAY 1 (11/2 월) 바이브코딩 기본기 및 MVP 설계"},
        day1_slides,
        os.path.join(OUT_DIR, "DAY1_바이브코딩_기본기_및_MVP_설계.pptx"),
        os.path.join(OUT_DIR, "DAY1_바이브코딩_기본기_및_MVP_설계.pdf")
    )
    build_deck(
        {"day_label": "DAY 2 (11/9 월) MVP 개발 및 실전 멘토링"},
        day2_slides,
        os.path.join(OUT_DIR, "DAY2_MVP_개발_디버깅_및_실전_멘토링.pptx"),
        os.path.join(OUT_DIR, "DAY2_MVP_개발_디버깅_및_실전_멘토링.pdf")
    )
    build_deck(
        {"day_label": "DAY 3 (11/13 금) 성과 발표 및 현업 적용 Wrap-up"},
        day3_slides,
        os.path.join(OUT_DIR, "DAY3_MVP_시연발표_및_현업적용_WrapUp.pptx"),
        os.path.join(OUT_DIR, "DAY3_MVP_시연발표_및_현업적용_WrapUp.pdf")
    )
