#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
5종 실습 예제 MVP용 standalone index.html 생성 스크립트
"""
import os
import json

BASE_DIR = "/Users/jsoftware/Documents/강의/AI_vibe_coding/04_직무별_실습예제_MVP_및_샘플데이터"

def make_dashboard_html(config):
    """
    범용적인 고품격 로컬 웹 대시보드 생성기 (내장 샘플데이터 + CSV 업로드(UTF-8/EUC-KR) + 필터 + Chart.js + 자동 브리핑 + CSV 다운로드)
    """
    return f"""<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>{config['title']}</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
  <style>
    @import url('https://cdn.jsdelivr.net/gh/orioncactus/pretendard/dist/web/static/pretendard.css');
    body {{ font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, 'Apple SD Gothic Neo', 'Malgun Gothic', sans-serif; }}
    @media print {{ .no-print {{ display: none !important; }} }}
  </style>
</head>
<body class="bg-slate-50 text-slate-800 min-h-screen">
  <!-- 상단 헤더 -->
  <header class="bg-slate-900 text-white shadow-md">
    <div class="max-w-7xl mx-auto px-6 py-4 flex flex-wrap items-center justify-between gap-4">
      <div>
        <span class="inline-block px-2.5 py-0.5 text-xs font-semibold bg-blue-600 rounded-full mb-1">{config['badge']}</span>
        <h1 class="text-xl font-bold tracking-tight">{config['title']}</h1>
        <p class="text-xs text-slate-300 mt-0.5">{config['subtitle']}</p>
      </div>
      <div class="flex items-center gap-2 no-print">
        <label class="cursor-pointer bg-blue-600 hover:bg-blue-500 text-white text-sm font-medium px-4 py-2 rounded-lg transition shadow-sm">
          📂 CSV 파일 업로드
          <input type="file" id="csvInput" accept=".csv" class="hidden" />
        </label>
        <button id="resetBtn" class="bg-slate-700 hover:bg-slate-600 text-slate-100 text-sm font-medium px-3.5 py-2 rounded-lg transition">
          🔄 기본 샘플 복원
        </button>
        <button id="exportBtn" class="bg-emerald-600 hover:bg-emerald-500 text-white text-sm font-medium px-3.5 py-2 rounded-lg transition">
          ⬇️ 분석결과 CSV 저장
        </button>
        <button onclick="window.print()" class="bg-slate-800 hover:bg-slate-700 border border-slate-600 text-slate-200 text-sm px-3 py-2 rounded-lg transition">
          🖨️ 인쇄/PDF
        </button>
      </div>
    </div>
  </header>

  <main class="max-w-7xl mx-auto px-6 py-6 space-y-6">
    <!-- KPI 요약 카드 4개 -->
    <section id="kpiCards" class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4"></section>

    <!-- 필터 바 & 임원 보고 자동 요약 -->
    <section class="grid grid-cols-1 lg:grid-cols-3 gap-6">
      <!-- 좌측: 필터 + 자동 브리핑 -->
      <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-5 flex flex-col justify-between space-y-4">
        <div>
          <h2 class="text-base font-bold text-slate-800 flex items-center gap-2">
            <span>🔍 조건 필터 및 이상 항목 탐지</span>
          </h2>
          <div class="mt-3 grid grid-cols-2 gap-3 no-print">
            <div>
              <label class="block text-xs font-semibold text-slate-500 mb-1">{config['group_label']} 필터</label>
              <select id="groupFilter" class="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500 outline-none">
                <option value="ALL">전체 보기</option>
              </select>
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-500 mb-1">상태 필터</label>
              <select id="statusFilter" class="w-full border border-slate-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-blue-500 outline-none">
                <option value="ALL">전체 상태</option>
                <option value="ALERT">🚨 점검/경고 항목만</option>
                <option value="OK">✅ 정상/우수 항목만</option>
              </select>
            </div>
          </div>
          {config.get('extra_control_html', '')}
        </div>

        <!-- 자동 생성 브리핑 박스 -->
        <div class="bg-slate-900 text-slate-100 rounded-lg p-4 space-y-2">
          <div class="flex items-center justify-between">
            <span class="text-xs font-bold text-blue-400">🤖 AI 자동 생성 3줄 보고 브리핑</span>
            <button id="copyBriefBtn" class="text-xs bg-slate-700 hover:bg-slate-600 px-2.5 py-1 rounded text-white no-print">📋 문구 복사</button>
          </div>
          <pre id="briefingText" class="text-xs leading-relaxed whitespace-pre-wrap font-sans text-slate-200"></pre>
        </div>
      </div>

      <!-- 우측: 시각화 차트 -->
      <div class="lg:col-span-2 bg-white rounded-xl shadow-sm border border-slate-200 p-5">
        <div class="flex items-center justify-between mb-3">
          <h2 class="text-base font-bold text-slate-800">📊 {config['chart_title']}</h2>
          <span class="text-xs text-slate-400">단위: {config['unit']}</span>
        </div>
        <div class="relative h-72 w-full">
          <canvas id="mainChart"></canvas>
        </div>
      </div>
    </section>

    <!-- 상세 데이터 테이블 -->
    <section class="bg-white rounded-xl shadow-sm border border-slate-200 overflow-hidden">
      <div class="px-5 py-4 border-b border-slate-200 flex items-center justify-between">
        <h2 class="text-base font-bold text-slate-800">📋 상세 데이터 및 자동 계산 결과</h2>
        <span id="rowCountText" class="text-xs font-medium text-slate-500"></span>
      </div>
      <div class="overflow-x-auto">
        <table class="w-full text-sm text-left">
          <thead id="tableHead" class="bg-slate-100 text-slate-600 text-xs uppercase border-b border-slate-200"></thead>
          <tbody id="tableBody" class="divide-y divide-slate-200"></tbody>
        </table>
      </div>
    </section>
  </main>

  <script>
    const CONFIG = {json.dumps(config, ensure_ascii=False)};
    let currentRows = JSON.parse(JSON.stringify(CONFIG.default_rows));
    let chartInstance = null;

    function parseNum(v) {{
      if (typeof v === 'number') return v;
      if (!v) return 0;
      const cleaned = String(v).replace(/[^0-9.-]/g, '');
      const n = parseFloat(cleaned);
      return isNaN(n) ? 0 : n;
    }}

    function fmtNum(n, digits = 0) {{
      return Number(n || 0).toLocaleString('ko-KR', {{
        minimumFractionDigits: digits,
        maximumFractionDigits: digits
      }});
    }}

    function enrichRow(row) {{
      const a = parseNum(row[CONFIG.col_a]);
      const b = parseNum(row[CONFIG.col_b]);
      const diff = b - a;
      const ratio = a !== 0 ? (b / a) * 100 : 0;
      let isAlert = false;
      let statusLabel = '정상';
      if (CONFIG.mode === 'BVA') {{
        // 예산(a) 대비 실집행(b): 100% 초과 시 경고
        isAlert = ratio > 100;
        statusLabel = isAlert ? '🚨 예산초과' : (ratio >= 90 ? '⚠️ 소진주의' : '✅ 적정');
      }} else if (CONFIG.mode === 'ROI') {{
        // 비용(a) 대비 매출(b): ROAS 250% 미만 시 경고
        isAlert = ratio < 250;
        statusLabel = isAlert ? '🚨 저효율점검' : '✅ 고효율';
      }} else if (CONFIG.mode === 'QUOTE') {{
        // 표준단가(a), 제조원가(b), 시뮬레이션 할인율 적용
        const discSlider = document.getElementById('simDiscSlider');
        const extraDisc = discSlider ? parseNum(discSlider.value) : 0;
        const qtyInput = document.getElementById('simQtyInput');
        const qty = qtyInput ? Math.max(1, parseNum(qtyInput.value)) : 10;
        const sellUnit = a * (1 - extraDisc / 100);
        const marginRate = sellUnit > 0 ? ((sellUnit - b) / sellUnit) * 100 : 0;
        const minMargin = parseNum(row['최소승인마진율(%)']);
        isAlert = marginRate < minMargin;
        statusLabel = isAlert ? '🚨 승인필요(마진미달)' : '✅ 즉시발행가능';
        return {{
          ...row,
          _a: a * qty,
          _b: sellUnit * qty,
          _diff: (sellUnit - b) * qty,
          _ratio: marginRate,
          _isAlert: isAlert,
          _statusLabel: statusLabel,
          _calc1: fmtNum(sellUnit, 1) + '만원',
          _calc2: fmtNum((sellUnit - b) * qty, 0) + '만원',
          _calc3: fmtNum(marginRate, 1) + '%'
        }};
      }} else {{
        // 목표(a) 대비 실적(b): 95% 미만 시 경고
        isAlert = ratio < 95;
        statusLabel = isAlert ? '🚨 미달성점검' : (ratio >= 100 ? '✅ 목표초과달성' : '⚠️ 근접(95%+)');
      }}
      return {{
        ...row,
        _a: a,
        _b: b,
        _diff: diff,
        _ratio: ratio,
        _isAlert: isAlert,
        _statusLabel: statusLabel,
        _calc1: (diff >= 0 ? '+' : '') + fmtNum(diff, 1),
        _calc2: fmtNum(ratio, 1) + '%'
      }};
    }}

    function populateGroupFilter(rows) {{
      const select = document.getElementById('groupFilter');
      const currentVal = select.value;
      const groups = Array.from(new Set(rows.map(r => r[CONFIG.group_col]).filter(Boolean)));
      select.innerHTML = '<option value="ALL">전체 보기</option>' +
        groups.map(g => `<option value="${{g}}">${{g}}</option>`).join('');
      if (groups.includes(currentVal)) select.value = currentVal;
    }}

    function renderAll() {{
      const groupVal = document.getElementById('groupFilter').value;
      const statusVal = document.getElementById('statusFilter').value;
      const enriched = currentRows.map(enrichRow);

      const filtered = enriched.filter(r => {{
        if (groupVal !== 'ALL' && r[CONFIG.group_col] !== groupVal) return false;
        if (statusVal === 'ALERT' && !r._isAlert) return false;
        if (statusVal === 'OK' && r._isAlert) return false;
        return true;
      }});

      // KPI 집계
      const sumA = filtered.reduce((s, r) => s + r._a, 0);
      const sumB = filtered.reduce((s, r) => s + r._b, 0);
      const avgRatio = sumA !== 0 ? (sumB / sumA) * 100 : 0;
      const alertCount = filtered.filter(r => r._isAlert).length;

      const kpiHtml = [
        {{ label: CONFIG.kpi1_label, val: fmtNum(sumA, 0) + ' ' + CONFIG.unit, sub: '현재 필터 기준 합계', color: 'border-slate-300' }},
        {{ label: CONFIG.kpi2_label, val: fmtNum(sumB, 0) + ' ' + CONFIG.unit, sub: '실제 집계 합계', color: 'border-blue-500' }},
        {{ label: CONFIG.kpi3_label, val: fmtNum(avgRatio, 1) + '%', sub: '종합 비율/효율', color: 'border-indigo-500' }},
        {{ label: '점검 필요(경고) 건수', val: alertCount + ' 건', sub: '전체 ' + filtered.length + '건 중', color: alertCount > 0 ? 'border-rose-500 bg-rose-50/40' : 'border-emerald-500' }}
      ].map(k => `
        <div class="bg-white rounded-xl shadow-sm border-l-4 ${{k.color}} border border-slate-200 p-4">
          <p class="text-xs font-semibold text-slate-500">${{k.label}}</p>
          <p class="text-2xl font-extrabold text-slate-900 mt-1">${{k.val}}</p>
          <p class="text-xs text-slate-400 mt-1">${{k.sub}}</p>
        </div>
      `).join('');
      document.getElementById('kpiCards').innerHTML = kpiHtml;

      // 브리핑 생성
      const alerts = filtered.filter(r => r._isAlert);
      const topAlertNames = alerts.slice(0, 3).map(r => `${{r[CONFIG.group_col]}}(${{r[CONFIG.item_col]}})`).join(', ') || '없음';
      const briefLines = [
        `1. [종합 현황] 조회 대상 총 ${{filtered.length}}건의 ${{CONFIG.kpi1_label}} 합계는 ${{fmtNum(sumA)}}${{CONFIG.unit}}, ${{CONFIG.kpi2_label}} 합계는 ${{fmtNum(sumB)}}${{CONFIG.unit}} (종합 ${{fmtNum(avgRatio, 1)}}%) 입니다.`,
        `2. [핵심 이상항목] 기준 임계치를 벗어난 중점 점검 항목은 총 ${{alertCount}}건이며, 주요 대상은 [${{topAlertNames}}] 입니다.`,
        `3. [권고 조치] 점검 대상 항목에 대한 원인 분석 및 차월 계획 조정을 즉시 권고합니다.`
      ];
      document.getElementById('briefingText').textContent = briefLines.join('\\n');

      // 테이블 렌더링
      const cols = CONFIG.headers;
      document.getElementById('tableHead').innerHTML = '<tr>' +
        cols.map(c => `<th class="px-4 py-3 font-semibold">${{c}}</th>`).join('') +
        `<th class="px-4 py-3 font-semibold text-right">${{CONFIG.calc1_label}}</th>` +
        `<th class="px-4 py-3 font-semibold text-right">${{CONFIG.calc2_label}}</th>` +
        `<th class="px-4 py-3 font-semibold text-center">판정상태</th>` +
        '</tr>';

      document.getElementById('tableBody').innerHTML = filtered.map(r => `
        <tr class="${{r._isAlert ? 'bg-rose-50/70 hover:bg-rose-100/60' : 'hover:bg-slate-50'}} transition">
          ${{cols.map(c => `<td class="px-4 py-2.5 whitespace-nowrap">${{r[c] ?? ''}}</td>`).join('')}}
          <td class="px-4 py-2.5 text-right font-mono font-semibold">${{r._calc1}}</td>
          <td class="px-4 py-2.5 text-right font-mono font-bold ${{r._isAlert ? 'text-rose-600' : 'text-blue-700'}}">${{r._calc3 || r._calc2}}</td>
          <td class="px-4 py-2.5 text-center">
            <span class="inline-block px-2.5 py-0.5 rounded-full text-xs font-bold ${{r._isAlert ? 'bg-rose-100 text-rose-700' : 'bg-emerald-100 text-emerald-700'}}">
              ${{r._statusLabel}}
            </span>
          </td>
        </tr>
      `).join('');
      document.getElementById('rowCountText').textContent = `총 ${{filtered.length}}개 항목 표시 중`;

      // 차트 업데이트
      const ctx = document.getElementById('mainChart').getContext('2d');
      if (chartInstance) chartInstance.destroy();
      const labels = filtered.map(r => `${{r[CONFIG.group_col]}}-${{r[CONFIG.item_col]}}`);
      chartInstance = new Chart(ctx, {{
        type: 'bar',
        data: {{
          labels,
          datasets: [
            {{
              label: CONFIG.col_a,
              data: filtered.map(r => r._a),
              backgroundColor: 'rgba(148, 163, 184, 0.7)',
              borderRadius: 5
            }},
            {{
              label: CONFIG.col_b,
              data: filtered.map(r => r._b),
              backgroundColor: filtered.map(r => r._isAlert ? 'rgba(225, 29, 72, 0.85)' : 'rgba(37, 99, 235, 0.85)'),
              borderRadius: 5
            }}
          ]
        }},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          plugins: {{ legend: {{ position: 'top' }} }},
          scales: {{ y: {{ beginAtZero: true }} }}
        }}
      }});
    }}

    // CSV 파서 (UTF-8 및 EUC-KR 자동 감지)
    function parseCSVText(text) {{
      const lines = text.replace(/\\r\\n/g, '\\n').replace(/\\r/g, '\\n').split('\\n').filter(l => l.trim() !== '');
      if (lines.length < 2) return [];
      const headers = lines[0].replace(/^\\uFEFF/, '').split(',').map(h => h.trim());
      return lines.slice(1).map(line => {{
        const vals = line.split(',').map(v => v.trim());
        const obj = {{}};
        headers.forEach((h, i) => obj[h] = vals[i] ?? '');
        return obj;
      }});
    }}

    document.getElementById('csvInput').addEventListener('change', (e) => {{
      const file = e.target.files[0];
      if (!file) return;
      const reader = new FileReader();
      reader.onload = (ev) => {{
        const buffer = ev.target.result;
        let text = new TextDecoder('utf-8').decode(buffer);
        if (text.includes('\\uFFFD')) {{
          text = new TextDecoder('euc-kr').decode(buffer);
        }}
        const parsed = parseCSVText(text);
        if (parsed.length > 0) {{
          currentRows = parsed;
          populateGroupFilter(currentRows);
          renderAll();
        }}
      }};
      reader.readAsArrayBuffer(file);
    }});

    document.getElementById('resetBtn').addEventListener('click', () => {{
      currentRows = JSON.parse(JSON.stringify(CONFIG.default_rows));
      populateGroupFilter(currentRows);
      document.getElementById('groupFilter').value = 'ALL';
      document.getElementById('statusFilter').value = 'ALL';
      renderAll();
    }});

    document.getElementById('groupFilter').addEventListener('change', renderAll);
    document.getElementById('statusFilter').addEventListener('change', renderAll);

    document.getElementById('copyBriefBtn').addEventListener('click', () => {{
      navigator.clipboard.writeText(document.getElementById('briefingText').textContent);
      const btn = document.getElementById('copyBriefBtn');
      btn.textContent = '✅ 복사 완료!';
      setTimeout(() => btn.textContent = '📋 문구 복사', 1500);
    }});

    document.getElementById('exportBtn').addEventListener('click', () => {{
      const enriched = currentRows.map(enrichRow);
      const headers = [...CONFIG.headers, CONFIG.calc1_label, CONFIG.calc2_label, '판정상태'];
      const csvRows = [
        headers.join(','),
        ...enriched.map(r => [
          ...CONFIG.headers.map(h => r[h]),
          String(r._calc1).replace(/,/g, ''),
          String(r._calc3 || r._calc2).replace(/,/g, ''),
          r._statusLabel
        ].join(','))
      ];
      const blob = new Blob(['\\uFEFF' + csvRows.join('\\r\\n')], {{ type: 'text/csv;charset=utf-8;' }});
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = 'analysis_result.csv';
      a.click();
    }});

    // 초기화
    populateGroupFilter(currentRows);
    renderAll();
  </script>
</body>
</html>
"""

configs = [
    {
        "folder": "00_공통예제_영업실적_대시보드",
        "badge": "DAY 1 공통 실습 데모",
        "title": "영업 실적 분석 및 달성률 대시보드",
        "subtitle": "CSV 업로드 기반 지역본부·제품군별 매출 달성률 자동 계산 및 임원 브리핑 생성 도구",
        "mode": "KPI",
        "unit": "백만원",
        "group_label": "지역본부",
        "group_col": "지역본부",
        "item_col": "제품군",
        "col_a": "목표매출(백만원)",
        "col_b": "실적매출(백만원)",
        "kpi1_label": "목표 매출 합계",
        "kpi2_label": "실적 매출 합계",
        "kpi3_label": "평균 목표 달성률",
        "calc1_label": "목표대비 증감",
        "calc2_label": "달성률(%)",
        "chart_title": "지역본부·제품군별 목표 매출 vs 실적 매출 비교",
        "headers": ["연월", "지역본부", "제품군", "담당팀", "목표매출(백만원)", "실적매출(백만원)", "전년동월(백만원)", "신규고객수"],
        "default_rows": [
            {"연월": "2026-08", "지역본부": "수도권본부", "제품군": "산업용모터", "담당팀": "영업1팀", "목표매출(백만원)": 1200, "실적매출(백만원)": 1320, "전년동월(백만원)": 1150, "신규고객수": 8},
            {"연월": "2026-08", "지역본부": "수도권본부", "제품군": "인버터제어기", "담당팀": "영업2팀", "목표매출(백만원)": 950, "실적매출(백만원)": 880, "전년동월(백만원)": 910, "신규고객수": 5},
            {"연월": "2026-09", "지역본부": "수도권본부", "제품군": "산업용모터", "담당팀": "영업1팀", "목표매출(백만원)": 1250, "실적매출(백만원)": 1390, "전년동월(백만원)": 1180, "신규고객수": 9},
            {"연월": "2026-09", "지역본부": "중부권본부", "제품군": "산업용모터", "담당팀": "중부팀", "목표매출(백만원)": 850, "실적매출(백만원)": 790, "전년동월(백만원)": 810, "신규고객수": 5},
            {"연월": "2026-10", "지역본부": "수도권본부", "제품군": "산업용모터", "담당팀": "영업1팀", "목표매출(백만원)": 1300, "실적매출(백만원)": 1480, "전년동월(백만원)": 1220, "신규고객수": 11},
            {"연월": "2026-10", "지역본부": "수도권본부", "제품군": "인버터제어기", "담당팀": "영업2팀", "목표매출(백만원)": 1050, "실적매출(백만원)": 940, "전년동월(백만원)": 980, "신규고객수": 6},
            {"연월": "2026-10", "지역본부": "중부권본부", "제품군": "산업용모터", "담당팀": "중부팀", "목표매출(백만원)": 900, "실적매출(백만원)": 960, "전년동월(백만원)": 830, "신규고객수": 7},
            {"연월": "2026-10", "지역본부": "영남권본부", "제품군": "스마트센서", "담당팀": "영남1팀", "목표매출(백만원)": 1200, "실적매출(백만원)": 1020, "전년동월(백만원)": 1100, "신규고객수": 5}
        ]
    },
    {
        "folder": "01_전략기획_경영KPI_대시보드",
        "badge": "전략기획 권장 MVP",
        "title": "전사 경영 KPI 달성률 & 임원 보고 자동 요약 대시보드",
        "subtitle": "사업부별 핵심 성과지표(KPI) 달성률 자동 판정 및 미달성 지표 중심 경영진 브리핑 생성기",
        "mode": "KPI",
        "unit": "포인트/억원",
        "group_label": "사업부",
        "group_col": "사업부",
        "item_col": "KPI지표명",
        "col_a": "목표치",
        "col_b": "실적치",
        "kpi1_label": "목표 수치 합산",
        "kpi2_label": "실적 수치 합산",
        "kpi3_label": "종합 KPI 달성률",
        "calc1_label": "목표대비 차이",
        "calc2_label": "KPI 달성률(%)",
        "chart_title": "사업부별 주요 KPI 목표치 vs 실적치 비교",
        "headers": ["연월", "사업부", "KPI지표명", "단위", "목표치", "실적치", "전월실적", "가중치(%)"],
        "default_rows": [
            {"연월": "2026-10", "사업부": "전장부품사업부", "KPI지표명": "매출액", "단위": "억원", "목표치": 450, "실적치": 472, "전월실적": 438, "가중치(%)": 30},
            {"연월": "2026-10", "사업부": "전장부품사업부", "KPI지표명": "영업이익률", "단위": "%", "목표치": 8.5, "실적치": 9.1, "전월실적": 8.2, "가중치(%)": 25},
            {"연월": "2026-10", "사업부": "전장부품사업부", "KPI지표명": "납기준수율", "단위": "%", "목표치": 98.0, "실적치": 91.2, "전월실적": 97.5, "가중치(%)": 20},
            {"연월": "2026-10", "사업부": "에너지솔루션사업부", "KPI지표명": "매출액", "단위": "억원", "목표치": 380, "실적치": 342, "전월실적": 360, "가중치(%)": 30},
            {"연월": "2026-10", "사업부": "에너지솔루션사업부", "KPI지표명": "영업이익률", "단위": "%", "목표치": 7.5, "실적치": 6.4, "전월실적": 7.1, "가중치(%)": 25},
            {"연월": "2026-10", "사업부": "에너지솔루션사업부", "KPI지표명": "신규수주액", "단위": "억원", "목표치": 520, "실적치": 585, "전월실적": 490, "가중치(%)": 25},
            {"연월": "2026-10", "사업부": "스마트팩토리사업부", "KPI지표명": "매출액", "단위": "억원", "목표치": 290, "실적치": 305, "전월실적": 275, "가중치(%)": 30},
            {"연월": "2026-10", "사업부": "스마트팩토리사업부", "KPI지표명": "설비불량률감소", "단위": "%", "목표치": 95.0, "실적치": 88.5, "전월실적": 91.0, "가중치(%)": 20}
        ]
    },
    {
        "folder": "02_재무_예산대비_실적분석기",
        "badge": "재무/회계 권장 MVP",
        "title": "부서별 예산 대비 실적(BVA) 편차 분석 및 이상 지출 탐지기",
        "subtitle": "월간 예산·실집행 내역 대조, 소진율 100% 초과 계정 자동 색출 및 소명 요청 리포트 도구",
        "mode": "BVA",
        "unit": "백만원",
        "group_label": "부서명",
        "group_col": "부서명",
        "item_col": "계정과목",
        "col_a": "예산액(백만원)",
        "col_b": "실집행액(백만원)",
        "kpi1_label": "편성 예산 합계",
        "kpi2_label": "실제 집행 합계",
        "kpi3_label": "평균 예산 소진율",
        "calc1_label": "편차(실적-예산)",
        "calc2_label": "예산소진율(%)",
        "chart_title": "부서·계정별 편성 예산 vs 실제 집행액 비교 (적색: 예산 초과)",
        "headers": ["연월", "부서명", "계정과목", "예산액(백만원)", "실집행액(백만원)", "전월집행액(백만원)", "담당자"],
        "default_rows": [
            {"연월": "2026-10", "부서명": "생산기술팀", "계정과목": "설비유지보수비", "예산액(백만원)": 180, "실집행액(백만원)": 215, "전월집행액(백만원)": 165, "담당자": "김철수"},
            {"연월": "2026-10", "부서명": "생산기술팀", "계정과목": "소모품비", "예산액(백만원)": 60, "실집행액(백만원)": 54, "전월집행액(백만원)": 58, "담당자": "김철수"},
            {"연월": "2026-10", "부서명": "글로벌마케팅팀", "계정과목": "해외전시회참가비", "예산액(백만원)": 240, "실집행액(백만원)": 278, "전월집행액(백만원)": 120, "담당자": "박영희"},
            {"연월": "2026-10", "부서명": "글로벌마케팅팀", "계정과목": "디지털광고선전비", "예산액(백만원)": 150, "실집행액(백만원)": 142, "전월집행액(백만원)": 148, "담당자": "박영희"},
            {"연월": "2026-10", "부서명": "R&D선행연구팀", "계정과목": "시제품제작비", "예산액(백만원)": 320, "실집행액(백만원)": 368, "전월집행액(백만원)": 290, "담당자": "이민호"},
            {"연월": "2026-10", "부서명": "R&D선행연구팀", "계정과목": "외주용역비", "예산액(백만원)": 200, "실집행액(백만원)": 185, "전월집행액(백만원)": 195, "담당자": "이민호"},
            {"연월": "2026-10", "부서명": "국내영업1팀", "계정과목": "물류운반비", "예산액(백만원)": 130, "실집행액(백만원)": 149, "전월집행액(백만원)": 128, "담당자": "최지훈"}
        ]
    },
    {
        "folder": "03_마케팅_캠페인_ROI_분석기",
        "badge": "마케팅 권장 MVP",
        "title": "B2B 마케팅 채널별 캠페인 ROI & ROAS 성과 분석기",
        "subtitle": "매체별 집행 비용 대비 수주 전환 매출(ROAS) 자동 계산 및 저효율 캠페인 자동 진단 대시보드",
        "mode": "ROI",
        "unit": "만원",
        "group_label": "채널구분",
        "group_col": "채널구분",
        "item_col": "캠페인명",
        "col_a": "집행비용(만원)",
        "col_b": "수주전환매출(만원)",
        "kpi1_label": "총 마케팅 집행비",
        "kpi2_label": "총 수주전환 매출",
        "kpi3_label": "평균 ROAS(매출/비용)",
        "calc1_label": "순기여매출(매출-비용)",
        "calc2_label": "ROAS(%)",
        "chart_title": "캠페인별 마케팅 집행비용 vs 수주전환 매출 비교 (기준 ROAS 250%)",
        "headers": ["캠페인명", "채널구분", "집행비용(만원)", "노출수", "클릭수", "리드발굴수", "수주전환매출(만원)"],
        "default_rows": [
            {"캠페인명": "스마트센서_검색키워드_10월", "채널구분": "검색광고(SA)", "집행비용(만원)": 1800, "노출수": 240000, "클릭수": 8400, "리드발굴수": 310, "수주전환매출(만원)": 7920},
            {"캠페인명": "산업용모터_링크드인_B2B", "채널구분": "글로벌SNS", "집행비용(만원)": 2500, "노출수": 180000, "클릭수": 5200, "리드발굴수": 245, "수주전환매출(만원)": 12750},
            {"캠페인명": "에너지솔루션_유튜브브랜딩", "채널구분": "동영상광고", "집행비용(만원)": 3200, "노출수": 620000, "클릭수": 6800, "리드발굴수": 95, "수주전환매출(만원)": 3840},
            {"캠페인명": "자동화박람회_리타겟팅배너", "채널구분": "디스플레이(DA)", "집행비용(만원)": 1400, "노출수": 410000, "클릭수": 4900, "리드발굴수": 120, "수주전환매출(만원)": 3120},
            {"캠페인명": "기술백서_다운로드_뉴스레터", "채널구분": "이메일/CRM", "집행비용(만원)": 600, "노출수": 45000, "클릭수": 3600, "리드발굴수": 280, "수주전환매출(만원)": 4680},
            {"캠페인명": "신제품_인버터_웨비나프로모션", "채널구분": "웨비나/이벤트", "집행비용(만원)": 2100, "노출수": 85000, "클릭수": 4100, "리드발굴수": 340, "수주전환매출(만원)": 11340}
        ]
    },
    {
        "folder": "04_세일즈_스마트견적_및_마진시뮬레이터",
        "badge": "세일즈/영업관리 권장 MVP",
        "title": "B2B 제조 제품 스마트 견적 산출 및 최소 마진율 검증 시뮬레이터",
        "subtitle": "수량 및 할인율 적용 시 예상 이익·마진율 실시간 계산 및 승인 기준 미달 자동 경고 도구",
        "mode": "QUOTE",
        "unit": "만원",
        "group_label": "제품카테고리",
        "group_col": "제품카테고리",
        "item_col": "제품명",
        "col_a": "표준단가(만원)",
        "col_b": "제조원가(만원)",
        "kpi1_label": "표준 견적총액",
        "kpi2_label": "할인적용 공급총액",
        "kpi3_label": "시뮬레이션 비율",
        "calc1_label": "할인적용 단가",
        "calc2_label": "예상마진율(%)",
        "chart_title": "품목별 표준 견적액 vs 할인 적용 공급액 비교",
        "extra_control_html": """
          <div class="mt-3 pt-3 border-t border-slate-200 grid grid-cols-2 gap-3 no-print">
            <div>
              <label class="block text-xs font-semibold text-blue-700 mb-1">적용 할인율 시뮬레이션 (<span id="discValLabel">8</span>%)</label>
              <input type="range" id="simDiscSlider" min="0" max="25" value="8" class="w-full accent-blue-600" oninput="document.getElementById('discValLabel').textContent=this.value; renderAll();" />
            </div>
            <div>
              <label class="block text-xs font-semibold text-blue-700 mb-1">견적 수량 (대/EA)</label>
              <input type="number" id="simQtyInput" min="1" max="1000" value="10" class="w-full border border-slate-300 rounded-lg px-3 py-1.5 text-sm" oninput="renderAll();" />
            </div>
          </div>
        """,
        "headers": ["품목코드", "제품명", "제품카테고리", "표준단가(만원)", "제조원가(만원)", "권장할인한도(%)", "최소승인마진율(%)"],
        "default_rows": [
            {"품목코드": "MOT-500A", "제품명": "고효율 서보모터 5kW", "제품카테고리": "구동모터", "표준단가(만원)": 185, "제조원가(만원)": 122, "권장할인한도(%)": 12, "최소승인마진율(%)": 22},
            {"품목코드": "MOT-750B", "제품명": "방폭형 산업모터 7.5kW", "제품카테고리": "구동모터", "표준단가(만원)": 260, "제조원가(만원)": 178, "권장할인한도(%)": 10, "최소승인마진율(%)": 20},
            {"품목코드": "INV-300X", "제품명": "스마트 인버터 제어반", "제품카테고리": "제어시스템", "표준단가(만원)": 420, "제조원가(만원)": 265, "권장할인한도(%)": 15, "최소승인마진율(%)": 25},
            {"품목코드": "SEN-100P", "제품명": "초정밀 진동·온도 센서모듈", "제품카테고리": "스마트센서", "표준단가(만원)": 68, "제조원가(만원)": 39, "권장할인한도(%)": 18, "최소승인마진율(%)": 28},
            {"품목코드": "PLC-900M", "제품명": "통합 공정 컨트롤러 패키지", "제품카테고리": "제어시스템", "표준단가(만원)": 680, "제조원가(만원)": 442, "권장할인한도(%)": 12, "최소승인마진율(%)": 24}
        ]
    }
]

for cfg in configs:
    out_path = os.path.join(BASE_DIR, cfg["folder"], "index.html")
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(make_dashboard_html(cfg))
    print("Generated:", out_path)
