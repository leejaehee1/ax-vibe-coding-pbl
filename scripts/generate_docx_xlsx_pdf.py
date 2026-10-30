#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
마크다운 원본 문서들을 기반으로 고품질 DOCX, PDF 및 강사용 평가/멘토링 XLSX를 생성하는 스크립트
"""
import os
import re
import subprocess
import tempfile
import docx
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = "/Users/jsoftware/Documents/강의/AI_vibe_coding"

def set_korean_font(run, font_name="맑은 고딕", size_pt=10.5, bold=False, color_rgb=None):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.size = Pt(size_pt)
    run.bold = bold
    if color_rgb:
        run.font.color.rgb = RGBColor(*color_rgb)

def shade_cell(cell, hex_color):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading)

def clean_md_inline(text):
    text = re.sub(r'\*\*(.*?)\*\*', r'\1', text)
    text = re.sub(r'`(.*?)`', r'\1', text)
    text = text.replace('<br>', '\n')
    return text.strip()

def md_to_docx(md_path, docx_path):
    with open(md_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    style = doc.styles['Normal']
    style.font.name = '맑은 고딕'
    style._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
    style.font.size = Pt(10)

    in_code = False
    code_lines = []
    table_rows = []

    def flush_table():
        nonlocal table_rows
        if not table_rows:
            return
        # 구분선(|---|---|) 제거
        valid_rows = []
        for r in table_rows:
            cells = [c.strip() for c in r.strip().strip('|').split('|')]
            if all(re.match(r'^:?-{2,}:?$', c) for c in cells if c):
                continue
            valid_rows.append(cells)
        if not valid_rows:
            table_rows = []
            return
        num_cols = max(len(r) for r in valid_rows)
        table = doc.add_table(rows=len(valid_rows), cols=num_cols)
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for r_idx, row_data in enumerate(valid_rows):
            for c_idx in range(num_cols):
                val = clean_md_inline(row_data[c_idx]) if c_idx < len(row_data) else ""
                cell = table.cell(r_idx, c_idx)
                cell.text = ""
                p = cell.paragraphs[0]
                run = p.add_run(val)
                if r_idx == 0:
                    shade_cell(cell, "1E293B")
                    set_korean_font(run, size_pt=9.5, bold=True, color_rgb=(255, 255, 255))
                else:
                    if r_idx % 2 == 0:
                        shade_cell(cell, "F8FAFC")
                    set_korean_font(run, size_pt=9.0, bold=(c_idx == 0), color_rgb=(30, 41, 59))
        doc.add_paragraph()
        table_rows = []

    def flush_code():
        nonlocal code_lines
        if not code_lines:
            return
        tbl = doc.add_table(rows=1, cols=1)
        tbl.style = 'Table Grid'
        cell = tbl.cell(0, 0)
        shade_cell(cell, "F1F5F9")
        p = cell.paragraphs[0]
        run = p.add_run("\n".join(code_lines))
        set_korean_font(run, font_name="Consolas", size_pt=9.0, color_rgb=(15, 23, 42))
        doc.add_paragraph()
        code_lines = []

    for raw_line in lines:
        line = raw_line.rstrip('\n')
        if line.strip().startswith('```'):
            flush_table()
            if not in_code:
                in_code = True
                code_lines = []
            else:
                in_code = False
                flush_code()
            continue

        if in_code:
            code_lines.append(line)
            continue

        if line.strip().startswith('|') and line.strip().endswith('|'):
            table_rows.append(line)
            continue
        else:
            flush_table()

        stripped = line.strip()
        if not stripped or stripped == '---':
            continue

        if stripped.startswith('# '):
            p = doc.add_paragraph()
            run = p.add_run(clean_md_inline(stripped[2:]))
            set_korean_font(run, size_pt=18, bold=True, color_rgb=(15, 23, 42))
        elif stripped.startswith('## '):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(12)
            run = p.add_run("■ " + clean_md_inline(stripped[3:]))
            set_korean_font(run, size_pt=13.5, bold=True, color_rgb=(30, 64, 175))
        elif stripped.startswith('### '):
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(8)
            run = p.add_run("• " + clean_md_inline(stripped[4:]))
            set_korean_font(run, size_pt=11.5, bold=True, color_rgb=(30, 41, 59))
        elif stripped.startswith('#### '):
            p = doc.add_paragraph()
            run = p.add_run(clean_md_inline(stripped[5:]))
            set_korean_font(run, size_pt=10.5, bold=True, color_rgb=(51, 65, 85))
        elif stripped.startswith('> '):
            tbl = doc.add_table(rows=1, cols=1)
            cell = tbl.cell(0, 0)
            shade_cell(cell, "EFF6FF")
            p = cell.paragraphs[0]
            run = p.add_run(clean_md_inline(stripped[2:]))
            set_korean_font(run, size_pt=9.5, bold=False, color_rgb=(30, 58, 138))
        elif stripped.startswith('- ') or stripped.startswith('* '):
            p = doc.add_paragraph(style='List Bullet')
            run = p.add_run(clean_md_inline(stripped[2:]))
            set_korean_font(run, size_pt=10.0, color_rgb=(30, 41, 59))
        elif re.match(r'^\d+\.\s+', stripped):
            p = doc.add_paragraph()
            run = p.add_run(clean_md_inline(stripped))
            set_korean_font(run, size_pt=10.0, color_rgb=(30, 41, 59))
        else:
            p = doc.add_paragraph()
            run = p.add_run(clean_md_inline(stripped))
            set_korean_font(run, size_pt=10.0, color_rgb=(51, 65, 85))

    flush_table()
    flush_code()
    doc.save(docx_path)
    print("Saved DOCX:", docx_path)

def md_to_pdf(md_path, pdf_path, doc_title):
    """
    마크다운을 정교한 HTML로 변환 후 wkhtmltopdf로 고화질 PDF 생성
    """
    with open(md_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    html_parts = []
    in_code = False
    code_buf = []
    table_buf = []
    in_ul = False

    def format_inline(t):
        t = re.sub(r'\*\*(.*?)\*\*', r'<strong>\1</strong>', t)
        t = re.sub(r'`(.*?)`', r'<code>\1</code>', t)
        return t

    def flush_ul():
        nonlocal in_ul
        if in_ul:
            html_parts.append("</ul>")
            in_ul = False

    def flush_table():
        nonlocal table_buf
        if not table_buf:
            return
        rows = []
        for r in table_buf:
            cells = [c.strip() for c in r.strip().strip('|').split('|')]
            if all(re.match(r'^:?-{2,}:?$', c) for c in cells if c):
                continue
            rows.append(cells)
        if rows:
            html_parts.append('<table><thead><tr>')
            for c in rows[0]:
                html_parts.append(f'<th>{format_inline(c)}</th>')
            html_parts.append('</tr></thead><tbody>')
            for r in rows[1:]:
                html_parts.append('<tr>')
                for c in r:
                    html_parts.append(f'<td>{format_inline(c)}</td>')
                html_parts.append('</tr>')
            html_parts.append('</tbody></table>')
        table_buf = []

    for raw in lines:
        line = raw.rstrip('\n')
        if line.strip().startswith('```'):
            flush_ul()
            flush_table()
            if not in_code:
                in_code = True
                code_buf = []
            else:
                in_code = False
                escaped = "\n".join(code_buf).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                html_parts.append(f'<pre class="code-box">{escaped}</pre>')
            continue
        if in_code:
            code_buf.append(line)
            continue
        if line.strip().startswith('|') and line.strip().endswith('|'):
            flush_ul()
            table_buf.append(line)
            continue
        else:
            flush_table()

        s = line.strip()
        if not s:
            flush_ul()
            continue
        if s == '---':
            flush_ul()
            html_parts.append('<hr/>')
        elif s.startswith('# '):
            flush_ul()
            html_parts.append(f'<h1>{format_inline(s[2:])}</h1>')
        elif s.startswith('## '):
            flush_ul()
            html_parts.append(f'<h2>{format_inline(s[3:])}</h2>')
        elif s.startswith('### '):
            flush_ul()
            html_parts.append(f'<h3>{format_inline(s[4:])}</h3>')
        elif s.startswith('#### '):
            flush_ul()
            html_parts.append(f'<h4>{format_inline(s[5:])}</h4>')
        elif s.startswith('> '):
            flush_ul()
            html_parts.append(f'<div class="callout">{format_inline(s[2:])}</div>')
        elif s.startswith('- ') or s.startswith('* '):
            if not in_ul:
                html_parts.append('<ul>')
                in_ul = True
            html_parts.append(f'<li>{format_inline(s[2:])}</li>')
        else:
            flush_ul()
            html_parts.append(f'<p>{format_inline(s)}</p>')

    flush_ul()
    flush_table()

    full_html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<title>{doc_title}</title>
<style>
  body {{
    font-family: 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif;
    color: #1e293b;
    line-height: 1.55;
    font-size: 10.5pt;
    margin: 0;
    padding: 10px 18px;
  }}
  h1 {{
    font-size: 19pt;
    color: #0f172a;
    border-bottom: 3px solid #1d4ed8;
    padding-bottom: 8px;
    margin-top: 0;
  }}
  h2 {{
    font-size: 13.5pt;
    color: #1e40af;
    background: #eff6ff;
    border-left: 5px solid #2563eb;
    padding: 6px 10px;
    margin-top: 18px;
    margin-bottom: 8px;
  }}
  h3 {{
    font-size: 11.5pt;
    color: #0f172a;
    margin-top: 14px;
    margin-bottom: 6px;
  }}
  h4 {{
    font-size: 10.5pt;
    color: #334155;
    margin-top: 10px;
    margin-bottom: 4px;
  }}
  p, li {{
    font-size: 10pt;
    color: #334155;
  }}
  ul {{
    margin-top: 4px;
    margin-bottom: 8px;
    padding-left: 20px;
  }}
  table {{
    width: 100%;
    border-collapse: collapse;
    margin: 10px 0 14px 0;
    font-size: 9.2pt;
    page-break-inside: auto;
  }}
  tr {{ page-break-inside: avoid; page-break-after: auto; }}
  th {{
    background-color: #1e293b;
    color: #ffffff;
    font-weight: 700;
    padding: 7px 8px;
    border: 1px solid #334155;
    text-align: left;
  }}
  td {{
    padding: 6px 8px;
    border: 1px solid #cbd5e1;
    vertical-align: top;
  }}
  tr:nth-child(even) td {{
    background-color: #f8fafc;
  }}
  code {{
    background: #f1f5f9;
    padding: 1px 4px;
    border-radius: 3px;
    font-family: Menlo, Consolas, monospace;
    font-size: 9pt;
    color: #be123c;
  }}
  pre.code-box {{
    background: #0f172a;
    color: #f8fafc;
    padding: 12px;
    border-radius: 6px;
    font-family: Menlo, Consolas, monospace;
    font-size: 8.8pt;
    white-space: pre-wrap;
    word-break: break-all;
    margin: 8px 0 12px 0;
  }}
  .callout {{
    background: #f0f9ff;
    border-left: 4px solid #0284c7;
    padding: 8px 12px;
    margin: 8px 0;
    font-size: 9.5pt;
    color: #0c4a6e;
  }}
  hr {{
    border: none;
    border-top: 1px solid #e2e8f0;
    margin: 14px 0;
  }}
</style>
</head>
<body>
{''.join(html_parts)}
</body>
</html>
"""
    with tempfile.NamedTemporaryFile('w', suffix='.html', encoding='utf-8', delete=False) as tf:
        tf.write(full_html)
        temp_html = tf.name

    try:
        subprocess.run([
            "/usr/local/bin/wkhtmltopdf",
            "--quiet",
            "--page-size", "A4",
            "--margin-top", "14mm",
            "--margin-bottom", "14mm",
            "--margin-left", "14mm",
            "--margin-right", "14mm",
            "--encoding", "UTF-8",
            temp_html,
            pdf_path
        ], check=True)
        print("Saved PDF:", pdf_path)
    finally:
        if os.path.exists(temp_html):
            os.remove(temp_html)

def generate_instructor_xlsx(xlsx_path):
    wb = openpyxl.Workbook()

    # 공통 스타일
    header_fill = PatternFill("solid", fgColor="1E293B")
    sub_fill = PatternFill("solid", fgColor="EFF6FF")
    zebra_fill = PatternFill("solid", fgColor="F8FAFC")
    total_fill = PatternFill("solid", fgColor="FEF3C7")
    header_font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
    title_font = Font(name="맑은 고딕", size=14, bold=True, color="0F172A")
    body_font = Font(name="맑은 고딕", size=10, color="1E293B")
    bold_font = Font(name="맑은 고딕", size=10, bold=True, color="1E293B")
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    # 1. 개인별_최종발표_평가표(20인)
    ws1 = wb.active
    ws1.title = "1_최종발표_평가표(20인)"
    ws1["A1"] = "대기업 제조사 AX 핵심인재 양성 교육 — 개인별 MVP 최종 발표 평가표 (DAY 3: 11/13)"
    ws1["A1"].font = title_font
    ws1["A2"] = "※ 평가 기준: 업무 문제 정의(20점) + 핵심 기능 구현(35점) + 사용성과 완성도(20점) + 업무 활용 효과(15점) + 발표 및 확장성(10점) = 100점 만점"
    ws1["A2"].font = Font(name="맑은 고딕", size=9.5, italic=True, color="475569")

    headers1 = [
        "발표순번", "소속조", "직무분과", "성명(직급)", "MVP 프로젝트 과제명",
        "① 문제정의\n(20점)", "② 핵심기능\n(35점)", "③ 사용성/완성도\n(20점)",
        "④ 업무활용효과\n(15점)", "⑤ 발표/확장성\n(10점)", "총점\n(100점)", "순위", "주강사 전문가 피드백 메모 (2~3분 코멘트용)"
    ]
    for col_idx, h in enumerate(headers1, 1):
        c = ws1.cell(row=4, column=col_idx, value=h)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = thin_border
    ws1.row_dimensions[4].height = 32

    sample_depts = ["전략기획", "전략기획", "전략기획", "전략기획", "재무/회계", "재무/회계", "재무/회계", "재무/회계",
                    "마케팅", "마케팅", "마케팅", "마케팅", "세일즈", "세일즈", "세일즈", "세일즈",
                    "세일즈", "전략기획", "재무/회계", "마케팅"]
    sample_projects = [
        "전사 경영 KPI 달성률 신호등 대시보드", "경쟁사 분기 실적 비교 브리핑 생성기", "신규 투자사업 타당성(ROI) 시뮬레이터", "월간 사업부 회의 안건/실적 자동 요약기",
        "부서별 예산 대비 실적(BVA) 이상 탐지기", "법인카드 지출 내역 계정과목 자동 분류기", "환율 변동에 따른 원가 민감도 계산기", "거래처 미수금 연령분석 및 알림 대시보드",
        "B2B 채널별 캠페인 ROAS 비교 분석기", "글로벌 전시회 리드(Lead) 스코어링 도구", "제품 카탈로그 문의 고객 전환율 트래커", "차월 마케팅 예산 재배분 시뮬레이터",
        "산업용 모터 스마트 견적 및 마진 검증기", "고객사별 단가 이력 및 할인율 조회 도구", "주간 영업 파이프라인 수주확률 대시보드", "대리점 월간 판매 장려금(리베이트) 계산기",
        "납기 지연 위험 수주건 자동 경보 보드", "원자재 단가 연동 판가 인상 시뮬레이터", "공장별 제조경비 편차 분석 리포트", "신제품 런칭 초기 수주 성과 트래커"
    ]
    sample_scores = [
        (19, 33, 18, 14, 9), (18, 31, 17, 13, 8), (17, 30, 16, 14, 9), (18, 32, 18, 13, 8),
        (20, 34, 19, 15, 9), (17, 29, 17, 13, 8), (18, 31, 16, 14, 8), (19, 32, 18, 14, 9),
        (19, 33, 19, 14, 9), (17, 30, 17, 12, 8), (18, 31, 17, 13, 8), (18, 32, 18, 14, 9),
        (20, 34, 18, 15, 10), (17, 30, 16, 13, 8), (18, 32, 17, 14, 8), (18, 31, 18, 13, 8),
        (17, 29, 16, 13, 8), (18, 31, 17, 14, 9), (19, 33, 18, 14, 9), (18, 30, 17, 13, 8)
    ]

    for i in range(20):
        r = 5 + i
        grp = f"{(i // 4) + 1}조"
        dept = sample_depts[i]
        name = f"교육생{i+1:02d} 과장"
        proj = sample_projects[i]
        s1, s2, s3, s4, s5 = sample_scores[i]

        row_vals = [i+1, grp, dept, name, proj, s1, s2, s3, s4, s5, f"=SUM(F{r}:J{r})", f"=RANK(K{r},$K$5:$K$24)", "핵심 기능 작동 우수 / 현업 CSV 양식 표준화 후 팀 내 즉시 도입 권장"]
        for c_idx, val in enumerate(row_vals, 1):
            cell = ws1.cell(row=r, column=c_idx, value=val)
            cell.font = bold_font if c_idx in (11, 12) else body_font
            cell.border = thin_border
            if i % 2 == 1:
                cell.fill = zebra_fill
            if c_idx == 11:
                cell.fill = total_fill
            if c_idx in (1, 2, 3, 4, 6, 7, 8, 9, 10, 11, 12):
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        ws1.row_dimensions[r].height = 22

    # 평균 행
    avg_r = 25
    ws1.cell(row=avg_r, column=1, value="전체 평균").alignment = Alignment(horizontal="center")
    ws1.merge_cells(start_row=avg_r, start_column=1, end_row=avg_r, end_column=5)
    for c_idx in range(1, 14):
        cell = ws1.cell(row=avg_r, column=c_idx)
        cell.fill = sub_fill
        cell.font = bold_font
        cell.border = thin_border
    for c_idx, col_let in [(6, 'F'), (7, 'G'), (8, 'H'), (9, 'I'), (10, 'J'), (11, 'K')]:
        c = ws1.cell(row=avg_r, column=c_idx, value=f"=ROUND(AVERAGE({col_let}5:{col_let}24),1)")
        c.alignment = Alignment(horizontal="center", vertical="center")

    col_widths1 = [9, 9, 13, 15, 36, 13, 13, 15, 14, 13, 11, 8, 48]
    for idx, w in enumerate(col_widths1, 1):
        ws1.column_dimensions[get_column_letter(idx)].width = w

    # 2. 조별_순회코칭_트래커(5개조)
    ws2 = wb.create_sheet(title="2_조별순회코칭_기록지(5개조)")
    ws2["A1"] = "조별 순회 코칭 트래커 (5개 조 × 4명, 조당 10분 타임박싱 운영)"
    ws2["A1"].font = title_font

    headers2 = ["조 편성", "분과 구성", "배정 시간(DAY 1)", "배정 시간(DAY 2)", "핵심 기능 1~2개 확정 상태", "공통 기술 이슈 (인코딩/차트/수식)", "주강사 코칭 조치 및 다음 단계 지시사항"]
    for idx, h in enumerate(headers2, 1):
        c = ws2.cell(row=3, column=idx, value=h)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = thin_border
    ws2.row_dimensions[3].height = 28

    groups_data = [
        ["1조 (1~4번)", "전략기획 분과", "17:10~17:20 (10분)", "16:10~16:25 (15분)", "✅ 4명 전원 핵심 기능 2개 확정", "사업부별 가중평균 수식 NaN 이슈", "ERP 직접 연동 제외, 월간 KPI CSV 업로드 + 임원 브리핑 자동 생성으로 통일"],
        ["2조 (5~8번)", "재무/회계 분과", "17:20~17:30 (10분)", "16:25~16:40 (15분)", "✅ 4명 전원 핵심 기능 2개 확정", "엑셀 천단위 콤마(,) 문자열 인식", "parseNumber() 방어 함수 적용 완료, 소진율 100% 초과 적색 하이라이트 구현"],
        ["3조 (9~12번)", "마케팅 분과", "17:30~17:40 (10분)", "16:40~16:55 (15분)", "✅ 4명 전원 핵심 기능 2개 확정", "Chart.js 필터 변경 시 차트 겹침", "chartInstance.destroy() 호출 패턴 적용, 채널별 ROAS 비교 차트 완성"],
        ["4조 (13~16번)", "세일즈 A분과", "17:40~17:50 (10분)", "16:55~17:10 (15분)", "✅ 4명 전원 핵심 기능 2개 확정", "인쇄 시 상단 버튼이 같이 출력됨", "@media print {.no-print {display:none}} 스타일 추가로 깔끔한 견적서 출력 해결"],
        ["5조 (17~20번)", "세일즈B/혼합 분과", "17:50~18:00 (10분)", "17:10~17:25 (15분)", "✅ 4명 전원 핵심 기능 2개 확정", "Windows 엑셀 EUC-KR 한글 깨짐", "TextDecoder('utf-8') 실패 시 'euc-kr' 자동 전환 코드 배포 완료"]
    ]
    for r_i, g_row in enumerate(groups_data, 4):
        for c_i, val in enumerate(g_row, 1):
            c = ws2.cell(row=r_i, column=c_i, value=val)
            c.font = body_font
            c.border = thin_border
            c.alignment = Alignment(horizontal="center" if c_i <= 5 else "left", vertical="center", wrap_text=True)
        ws2.row_dimensions[r_i].height = 30

    col_widths2 = [15, 16, 18, 18, 24, 30, 48]
    for idx, w in enumerate(col_widths2, 1):
        ws2.column_dimensions[get_column_letter(idx)].width = w

    # 3. DAY 3 발표 타임테이블 (160분)
    ws3 = wb.create_sheet(title="3_DAY3_발표타임테이블")
    ws3["A1"] = "DAY 3 (11/13 금) 개인별 MVP 시연 발표 타임테이블 (20명 × 8분 = 160분)"
    ws3["A1"].font = title_font
    headers3 = ["구분", "시간대", "소요시간", "발표자", "직무", "발표 구성 (5분 시연 + 3분 피드백)", "비고"]
    for idx, h in enumerate(headers3, 1):
        c = ws3.cell(row=3, column=idx, value=h)
        c.fill = header_fill
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = thin_border

    schedule_rows = [
        ["사전점검", "13:00~13:30", "30분", "전원", "공통", "최종 실행 점검, 데모 CSV 확인 및 해상도 테스트", "강사 순회 점검"],
        ["미니강의", "13:30~14:00", "30분", "주강사", "공통", "효과적인 MVP 시연법 (문제 1분 - 라이브 시연 3분 - 효과 1분)", "발표 순서 공지"],
        ["Part 1", "14:00~14:08", "8분", "교육생01 과장 (1조)", "전략기획", "시연 발표 5분 + 전문가 피드백 3분", "타이머 운영"],
        ["Part 1", "14:08~14:16", "8분", "교육생02 과장 (1조)", "전략기획", "시연 발표 5분 + 전문가 피드백 3분", ""],
        ["Part 1", "14:16~14:24", "8분", "교육생03 과장 (1조)", "전략기획", "시연 발표 5분 + 전문가 피드백 3분", ""],
        ["Part 1", "14:24~14:32", "8분", "교육생04 과장 (1조)", "전략기획", "시연 발표 5분 + 전문가 피드백 3분", ""],
        ["Part 1", "14:32~14:40", "8분", "교육생05 과장 (2조)", "재무/회계", "시연 발표 5분 + 전문가 피드백 3분", ""],
        ["Part 1", "14:40~14:48", "8분", "교육생06 과장 (2조)", "재무/회계", "시연 발표 5분 + 전문가 피드백 3분", ""],
        ["Part 1", "14:48~14:56", "8분", "교육생07 과장 (2조)", "재무/회계", "시연 발표 5분 + 전문가 피드백 3분", ""],
        ["Part 1", "14:56~15:04", "8분", "교육생08 과장 (2조)", "재무/회계", "시연 발표 5분 + 전문가 피드백 3분", ""],
        ["Part 1", "15:04~15:12", "8분", "교육생09 과장 (3조)", "마케팅", "시연 발표 5분 + 전문가 피드백 3분", ""],
        ["Part 1", "15:12~15:20", "8분", "교육생10 과장 (3조)", "마케팅", "시연 발표 5분 + 전문가 피드백 3분", ""],
        ["휴식", "15:20~15:30", "10분", "전원", "-", "중간 휴식 및 Part 2 발표자 세팅", "점수 중간 집계"],
        ["Part 2", "15:30~15:37", "7분", "교육생11 과장 (3조)", "마케팅", "시연 발표 4.5분 + 전문가 피드백 2.5분", ""],
        ["Part 2", "15:37~15:44", "7분", "교육생12 과장 (3조)", "마케팅", "시연 발표 4.5분 + 전문가 피드백 2.5분", ""],
        ["Part 2", "15:44~15:51", "7분", "교육생13 과장 (4조)", "세일즈", "시연 발표 4.5분 + 전문가 피드백 2.5분", ""],
        ["Part 2", "15:51~15:58", "7분", "교육생14 과장 (4조)", "세일즈", "시연 발표 4.5분 + 전문가 피드백 2.5분", ""],
        ["Part 2", "15:58~16:05", "7분", "교육생15 과장 (4조)", "세일즈", "시연 발표 4.5분 + 전문가 피드백 2.5분", ""],
        ["Part 2", "16:05~16:12", "7분", "교육생16 과장 (4조)", "세일즈", "시연 발표 4.5분 + 전문가 피드백 2.5분", ""],
        ["Part 2", "16:12~16:19", "7분", "교육생17 과장 (5조)", "세일즈", "시연 발표 4.5분 + 전문가 피드백 2.5분", ""],
        ["Part 2", "16:19~16:26", "7분", "교육생18 과장 (5조)", "전략기획", "시연 발표 4.5분 + 전문가 피드백 2.5분", ""],
        ["Part 2", "16:26~16:33", "7분", "교육생19 과장 (5조)", "재무/회계", "시연 발표 4.5분 + 전문가 피드백 2.5분", ""],
        ["Part 2", "16:33~16:40", "7분", "교육생20 과장 (5조)", "마케팅", "시연 발표 4.5분 + 전문가 피드백 2.5분", ""],
        ["Wrap-up", "16:40~17:00", "20분", "주강사/전원", "공통", "우수 프로젝트 선정(Top 3), 현업 확산 로드맵 총평 및 회고", "수료"]
    ]
    for r_i, s_row in enumerate(schedule_rows, 4):
        for c_i, val in enumerate(s_row, 1):
            c = ws3.cell(row=r_i, column=c_i, value=val)
            c.font = body_font
            c.border = thin_border
            c.alignment = Alignment(horizontal="center" if c_i <= 5 else "left", vertical="center")
        ws3.row_dimensions[r_i].height = 20

    col_widths3 = [12, 15, 10, 22, 14, 46, 20]
    for idx, w in enumerate(col_widths3, 1):
        ws3.column_dimensions[get_column_letter(idx)].width = w

    wb.save(xlsx_path)
    print("Saved XLSX:", xlsx_path)

if __name__ == "__main__":
    docs = [
        (
            os.path.join(ROOT, "01_커리큘럼_및_운영계획/대기업_제조사_AX_핵심인재_바이브코딩_커리큘럼.md"),
            os.path.join(ROOT, "01_커리큘럼_및_운영계획/대기업_제조사_AX_핵심인재_바이브코딩_커리큘럼.docx"),
            os.path.join(ROOT, "01_커리큘럼_및_운영계획/대기업_제조사_AX_핵심인재_바이브코딩_커리큘럼.pdf"),
            "대기업 제조사 AX 핵심인재 양성 교육 커리큘럼"
        ),
        (
            os.path.join(ROOT, "03_수강생_배포용_템플릿/01_과제정의서_및_MVP설계_템플릿.md"),
            os.path.join(ROOT, "03_수강생_배포용_템플릿/01_과제정의서_및_MVP설계_템플릿.docx"),
            os.path.join(ROOT, "03_수강생_배포용_템플릿/01_과제정의서_및_MVP설계_템플릿.pdf"),
            "개인별 AX 과제 정의서 및 MVP 범위 설계서"
        ),
        (
            os.path.join(ROOT, "03_수강생_배포용_템플릿/02_바이브코딩_프롬프트_템플릿_5종_가이드.md"),
            os.path.join(ROOT, "03_수강생_배포용_템플릿/02_바이브코딩_프롬프트_템플릿_5종_가이드.docx"),
            os.path.join(ROOT, "03_수강생_배포용_템플릿/02_바이브코딩_프롬프트_템플릿_5종_가이드.pdf"),
            "Claude Code 바이브코딩 프롬프트 템플릿 5종 가이드"
        ),
        (
            os.path.join(ROOT, "03_수강생_배포용_템플릿/03_MVP_테스트_체크리스트_및_시연발표_가이드.md"),
            os.path.join(ROOT, "03_수강생_배포용_템플릿/03_MVP_테스트_체크리스트_및_시연발표_가이드.docx"),
            os.path.join(ROOT, "03_수강생_배포용_템플릿/03_MVP_테스트_체크리스트_및_시연발표_가이드.pdf"),
            "MVP 테스트 체크리스트 및 5분 시연 발표 가이드"
        ),
        (
            os.path.join(ROOT, "03_수강생_배포용_템플릿/05_공통오류_대응_FAQ_및_보안가이드.md"),
            os.path.join(ROOT, "03_수강생_배포용_템플릿/05_공통오류_대응_FAQ_및_보안가이드.docx"),
            os.path.join(ROOT, "03_수강생_배포용_템플릿/05_공통오류_대응_FAQ_및_보안가이드.pdf"),
            "바이브코딩 공통 오류 대응 FAQ 및 사내 보안 가이드"
        ),
    ]
    for md_p, docx_p, pdf_p, title in docs:
        md_to_docx(md_p, docx_p)
        md_to_pdf(md_p, pdf_p, title)

    generate_instructor_xlsx(
        os.path.join(ROOT, "03_수강생_배포용_템플릿/04_강사용_프로젝트_평가표_및_멘토링_시트.xlsx")
    )
