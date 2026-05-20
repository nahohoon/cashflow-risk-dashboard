const FOLDER_ID = '1Z07c3mrAsx6QLNVSk_Dow6ERnsk-uVNN';
const PROCESSED_FOLDER_NAME = '처리완료';
const ERROR_LOG_SHEET_NAME = '오류로그';
const RAW_SHEET_NAME = 'RAW_거래처원장';

/** xlsx·CSV 공통 RAW 컬럼 순서 */
const RAW_LEDGER_HEADERS_ = [
  '거래처명',
  '원본파일',
  '행구분',
  '일자-No.',
  '품목명[규격]',
  '수량',
  '단가',
  '금액',
  '부가세',
  '구매',
  '판매',
  '수금',
  '지급',
  '잔액',
  '적요'
];

/** 거래처요약 시트 (ERP출처 + 8열). RAW 파서 출력과 동일 순서로 유지 */
const SUMMARY_SHEET_HEADERS_ = [[
  'ERP출처',
  '거래처명',
  '총판매',
  '총수금',
  '현재잔액',
  '최근수금일',
  '거래건수',
  '위험등급',
  '관리메모'
]];

/** 거래처요약 데이터 행 컬럼 인덱스 (SUMMARY_SHEET_HEADERS_ 순서와 일치) */
const SUMMARY_SHEET_COL_ = {
  erp: 0,
  name: 1,
  sales: 2,
  collect: 3,
  balance: 4,
  lastPay: 5,
  tradeCount: 6,
  risk: 7,
  memo: 8
};

/** RAW 행 인덱스 (parseKyungyoungDoctorCsv_ / parseLedgerXlsxFile_ 출력과 동일) */
const RAW_LEDGER_COL_ = {
  name: 0,
  file: 1,
  type: 2,
  date: 3,
  sales: 10,
  coll: 11,
  bal: 13
};

/**
 * 일자-No. 문자열에서 Date 해석 (요약·정렬용). RAW 값은 그대로 두고 이 함수만 보강.
 */
function parseLedgerDateNo_(dateNo, refDate) {
  refDate = refDate || new Date();
  if (Object.prototype.toString.call(dateNo) === '[object Date]' && !isNaN(dateNo.getTime())) {
    return dateNo;
  }
  const raw = String(dateNo == null ? '' : dateNo).trim();
  if (!raw) return null;

  const ymd = raw.match(/^(\d{4})[-.\/](\d{1,2})[-.\/](\d{1,2})/);
  if (ymd) {
    const y = Number(ymd[1]);
    const mo = Number(ymd[2]);
    const d = Number(ymd[3]);
    if (y && mo && d) return new Date(y, mo - 1, d);
  }

  const yyMd = raw.match(/^(\d{2})[.\/](\d{1,2})[.\/](\d{1,2})/);
  if (yyMd) {
    const yy = Number(yyMd[1]);
    const mo = Number(yyMd[2]);
    const d = Number(yyMd[3]);
    if (yy >= 0 && yy < 100 && mo >= 1 && mo <= 12 && d >= 1 && d <= 31) {
      const y = 2000 + yy;
      return new Date(y, mo - 1, d);
    }
  }

  const parts = raw.split('/');
  if (parts.length >= 2) {
    const month = Number(parts[0]);
    const rest = String(parts[1] || '');
    const dayMatch = rest.match(/^(\d+)/);
    const day = dayMatch ? Number(dayMatch[1]) : NaN;
    if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
      return new Date(refDate.getFullYear(), month - 1, day);
    }
  }

  return null;
}

function isCarriedForwardLedgerLabel_(itemName, memo) {
  const i = String(itemName == null ? '' : itemName);
  const m = String(memo == null ? '' : memo);
  if (i.indexOf('이월잔액') !== -1 || m.indexOf('이월잔액') !== -1) return true;
  if (/<\s*전\s*기\s*이\s*월\s*>/.test(i) || /<\s*전\s*기\s*이\s*월\s*>/.test(m)) return true;

  function fieldNormCf_(s) {
    return String(s || '')
      .replace(/\s+/g, '')
      .replace(/[<>《》]/g, '');
  }

  const fi = fieldNormCf_(i);
  const fm = fieldNormCf_(m);
  const both = fi + fm;

  if (both.indexOf('전기이월') !== -1) return true;
  if (fi === '이월' || fm === '이월') return true;
  if (both.indexOf('이월') !== -1 && (both.indexOf('전') !== -1 || both.indexOf('<') !== -1)) {
    return true;
  }
  return false;
}

function classifyLedgerRowType_(dateNo, itemName, memo) {
  const d = String(dateNo == null ? '' : dateNo);
  const i = String(itemName == null ? '' : itemName);
  const m = String(memo == null ? '' : memo);

  if (isCarriedForwardLedgerLabel_(i, m)) return '이월잔액';

  if (d.includes('합계') || i.includes('합계') || m.includes('합계')) return '합계';

  if (d.includes('계')) return '월계';

  return '거래';
}

function splitCsvStringToRows_(text) {
  const rows = [];
  let row = [];
  let field = '';
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const c = text[i];
    if (inQuotes) {
      if (c === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += c;
      }
    } else if (c === '"') {
      inQuotes = true;
    } else if (c === ',') {
      row.push(field);
      field = '';
    } else if (c === '\r') {
      continue;
    } else if (c === '\n') {
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
    } else {
      field += c;
    }
  }
  if (field.length || row.length) {
    row.push(field);
    rows.push(row);
  }
  return rows;
}

/** RAW 행구분(C열) 허용값만. classify 결과가 어긋나면 '거래'로 고정 */
var RAW_LEDGER_ROW_TYPES_ = { 거래: true, 월계: true, 합계: true, 이월잔액: true };

function sanitizeRowTypeForRaw_(rowType) {
  const t = String(rowType == null ? '' : rowType).trim();
  if (RAW_LEDGER_ROW_TYPES_[t]) return t;
  return '거래';
}

/**
 * A=거래처명, B=원본파일(fileName만), C=행구분(허용값만) — B·C에 값이 섞이지 않도록 단일 진입점.
 * tail12: 일자~적요 12칸
 */
function pushKyLedgerRawRow_(out, customerName, fileName, rowType, tail12) {
  const rt = sanitizeRowTypeForRaw_(rowType);
  const fn = String(fileName == null ? '' : fileName);
  if (/[\s]+(거래|월계|합계|이월잔액)\s*$/.test(fn)) {
    Logger.log('pushKyLedgerRawRow_: 원본파일명에 행구분 패턴이 포함됨(비정상) — 확인 필요: ' + fn);
  }
  out.push([String(customerName == null ? '' : customerName), fn, rt].concat(tail12));
}

/**
 * Utilities.parseCsv 우선(따옴표·쉼표 RFC 처리), 단열로만 나오면 커스텀 split 재시도.
 * 탭 구분은 parseCsv(텍스트, '\\t')로 보조.
 */
function parseLedgerCsvToRows_(text) {
  const probe = 'A,"B,C",D';
  try {
    const customProbe = splitCsvStringToRows_(probe);
    const parseProbe = Utilities.parseCsv(probe);
    Logger.log(
      'CSV split probe (quoted comma): custom=' +
        JSON.stringify(customProbe) +
        ' parseCsv=' +
        JSON.stringify(parseProbe)
    );
  } catch (probeErr) {
    Logger.log('CSV split probe error: ' + (probeErr && probeErr.message ? probeErr.message : String(probeErr)));
  }

  let rows = [];
  let method = '';
  if (!text) return rows;

  try {
    rows = Utilities.parseCsv(text);
    method = 'Utilities.parseCsv';
    if (!rows || !rows.length) throw new Error('empty parseCsv');
    let maxW = 0;
    const sampleN = Math.min(40, rows.length);
    for (let s = 0; s < sampleN; s++) {
      const w = rows[s] ? rows[s].length : 0;
      if (w > maxW) maxW = w;
    }
    if (maxW <= 1 && text.indexOf(',') !== -1) {
      Logger.log('parseLedgerCsvToRows_: parseCsv 단열·다중쉼표 → splitCsvStringToRows_ 재시도');
      rows = splitCsvStringToRows_(text);
      method = 'splitCsvStringToRows_(comma-rich)';
    }
  } catch (e) {
    rows = splitCsvStringToRows_(text);
    method = 'splitCsvStringToRows_(parseCsv threw: ' + (e && e.message ? e.message : String(e)) + ')';
  }

  if ((!rows || !rows.length) && String(text).trim()) {
    rows = splitCsvStringToRows_(text);
    method = 'splitCsvStringToRows_(fallback)';
  }

  if (rows && rows.length && rows[0] && rows[0].length <= 2 && text.indexOf('\t') !== -1) {
    try {
      const tabbed = Utilities.parseCsv(text, '\t');
      if (tabbed && tabbed[0] && tabbed[0].length > (rows[0] ? rows[0].length : 0)) {
        rows = tabbed;
        method = 'Utilities.parseCsv(tab)';
      }
    } catch (e2) {
      Logger.log('parseLedgerCsvToRows_: tab parse failed: ' + (e2 && e2.message ? e2.message : String(e2)));
    }
  }

  Logger.log('parseLedgerCsvToRows_: method=' + method + ' rowCount=' + (rows ? rows.length : 0));
  return rows || [];
}

/** 헤더 셀 normalize 후 후보 문자열과 완전 일치하는 첫 열 */
function findCsvColExactNormalized_(headerRow, candidates) {
  const list = typeof candidates === 'string' ? [candidates] : candidates.slice();
  const normHeaders = headerRow.map(function (h) {
    return normalizeHeaderText_(String(h == null ? '' : h).replace(/^\uFEFF/, ''));
  });
  for (let li = 0; li < list.length; li++) {
    const target = normalizeHeaderText_(list[li]);
    if (!target) continue;
    for (let c = 0; c < normHeaders.length; c++) {
      if (normHeaders[c] === target) return c;
    }
  }
  return -1;
}

function findCsvColPartialBothNormalized_(headerRow, a, b) {
  const na = normalizeHeaderText_(String(a == null ? '' : a));
  const nb = normalizeHeaderText_(String(b == null ? '' : b));
  for (let c = 0; c < headerRow.length; c++) {
    const h = normalizeHeaderText_(String(headerRow[c] == null ? '' : headerRow[c]).replace(/^\uFEFF/, ''));
    if (h.indexOf(na) !== -1 && h.indexOf(nb) !== -1) return c;
  }
  return -1;
}

/** 레거시 원장 CSV: 헤더 정규화 정확 일치 우선, 없으면 findCsvColByHeader_ 보조 */
function findLedgerLegacyColumnMap_(headerRow) {
  function exactThenLoose(candidates) {
    const idx = findCsvColExactNormalized_(headerRow, candidates);
    if (idx !== -1) return idx;
    return findCsvColByHeader_(headerRow, typeof candidates === 'string' ? [candidates] : candidates);
  }
  return {
    colDate: exactThenLoose(['일자-No.', '일자-No', '일자']),
    colItem: exactThenLoose(['품목명[규격]', '품목명']),
    colQty: exactThenLoose(['수량']),
    colPrice: exactThenLoose(['단가']),
    colAmt: exactThenLoose(['금액']),
    colVat: exactThenLoose(['부가세']),
    colBuy: exactThenLoose(['구매']),
    colSale: exactThenLoose(['판매']),
    colColl: exactThenLoose(['수금']),
    colPay: exactThenLoose(['지급']),
    colBal: exactThenLoose(['잔액']),
    colMemo: exactThenLoose(['적요', '비고'])
  };
}

function csvTextLooksCorrupted_(text) {
  if (!text || text.indexOf('\uFFFD') !== -1) return true;
  const sample = text.slice(0, 4000);
  const korean = (sample.match(/[가-힣]/g) || []).length;
  const broken = (sample.match(/[íìÎÏðñòóôõöøùúûüýþÿÄÅÆÇÈÉÊËÌ]/gi) || []).length;
  return korean < 2 && broken > 8;
}

function readLedgerCsvBlobAsString_(blob) {
  const encodings = ['MS949', 'EUC-KR', 'UTF-8'];
  for (let i = 0; i < encodings.length; i++) {
    try {
      const text = blob.getDataAsString(encodings[i]);
      if (!text || !String(text).trim()) {
        continue;
      }
      if (csvTextLooksCorrupted_(text)) {
        Logger.log('CSV encoding skipped (looks corrupted): ' + encodings[i]);
        continue;
      }
      Logger.log('CSV encoding selected: ' + encodings[i]);
      Logger.log(String(text).slice(0, 200));
      return text;
    } catch (e) {
      Logger.log('CSV encoding failed: ' + encodings[i] + ' / ' + (e && e.message ? e.message : String(e)));
    }
  }
  const fallback = blob.getDataAsString();
  Logger.log('CSV encoding selected: default');
  Logger.log(String(fallback).slice(0, 200));
  return fallback;
}

/** 공백 제거 후 보고서 제목 비교용 */
function normalizeKyReportTitle_(s) {
  return String(s == null ? '' : s).replace(/\s+/g, '');
}

/** 잔액명세 CSV 헤더 셀: BOM·앞뒤 공백 제거 후 내부 공백까지 제거해 '상 호 명' → '상호명' 등으로 비교 */
function normalizeHeaderText_(s) {
  return String(s == null ? '' : s)
    .replace(/^\uFEFF/, '')
    .trim()
    .replace(/\s+/g, '');
}

/**
 * 경영박사 '거래처별 잔액 명세' CSV 여부 (파일명 또는 본문 앞부분).
 */
function isKyungyoungDoctorBalanceCsvFile_(file) {
  const fn = normalizeKyReportTitle_(file.getName());
  if (fn.indexOf('거래처별잔액명세서') !== -1 || fn.indexOf('거래처별잔액명세') !== -1) {
    return true;
  }
  const text = readLedgerCsvBlobAsString_(file.getBlob()).slice(0, 6000);
  const rows = splitCsvStringToRows_(text).slice(0, 25);
  for (let i = 0; i < rows.length; i++) {
    const joined = normalizeKyReportTitle_(rows[i].join(','));
    if (joined.indexOf('거래처별잔액명세서') !== -1 || joined.indexOf('거래처별잔액명세') !== -1) {
      return true;
    }
  }
  return false;
}

/**
 * 잔액명세 헤더 행: 아래 토큰 그룹 중 실제로 나타난 그룹이 3개 이상이면 헤더로 인정.
 * (경영박사 파일마다 열 구성이 달라 상호·판매·수금·잔액 네 개를 모두 요구하지 않음)
 */
var BALANCE_CSV_HEADER_SCORE_GROUPS_ = [
  ['상호명'],
  ['거래처명'],
  ['거래처'],
  ['전기이월'],
  ['판매액', '판매'],
  ['수금액', '수금'],
  ['구매액'],
  ['지급액'],
  ['잔액'],
  ['전화', '전화번호']
];

function balanceCsvHeaderScoreFromNormalizedCells_(cellsNorm) {
  const present = {};
  for (let i = 0; i < cellsNorm.length; i++) {
    const t = cellsNorm[i];
    if (t) present[t] = true;
  }
  let score = 0;
  for (let g = 0; g < BALANCE_CSV_HEADER_SCORE_GROUPS_.length; g++) {
    const keys = BALANCE_CSV_HEADER_SCORE_GROUPS_[g];
    let hit = false;
    for (let k = 0; k < keys.length; k++) {
      if (present[keys[k]]) {
        hit = true;
        break;
      }
    }
    if (hit) score++;
  }
  return score;
}

/** 헤더 행 중 normalizeHeaderText_ 기준으로 점수 3 이상인 첫 행. 실패 시 첫 15행 로그. */
function findBalanceCsvHeaderRowIndex_(rows) {
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (!row || !row.length) continue;
    const cellsNorm = row.map(function (c) {
      return normalizeHeaderText_(c);
    });
    if (balanceCsvHeaderScoreFromNormalizedCells_(cellsNorm) >= 3) {
      return i;
    }
  }
  Logger.log('잔액명세 헤더 탐색 실패. 첫 15행:');
  const slice = rows.slice(0, 15);
  for (let j = 0; j < slice.length; j++) {
    const r = slice[j] || [];
    Logger.log(j + ': ' + r.join(' | '));
  }
  return -1;
}

function findKyBalanceCsvHeaderRowIndex_(rows) {
  return findBalanceCsvHeaderRowIndex_(rows);
}

/** 잔액명세 헤더: 정규화된 셀 문자열이 candidates 중 하나와 일치하는 첫 열 인덱스 */
function findBalanceCsvColumnIndex_(headerRow, normalizedNameCandidates) {
  for (let c = 0; c < headerRow.length; c++) {
    const h = normalizeHeaderText_(headerRow[c]);
    for (let k = 0; k < normalizedNameCandidates.length; k++) {
      if (h === normalizedNameCandidates[k]) return c;
    }
  }
  return -1;
}

function hasUsableLastPayForAging_(lastPay) {
  return parseLedgerDateNo_(lastPay, new Date()) != null;
}

/** 미수 관리 KPI용: 현재잔액 > 0 인 거래처요약 행만 */
function receivableSummaryRows_(rows, idxBalance) {
  return rows.filter(function (r) {
    return toNumber_(r[idxBalance]) > 0;
  });
}

function mergeBalanceSummaryRowsByName_(rowArrays) {
  const map = {};
  for (let i = 0; i < rowArrays.length; i++) {
    const row = rowArrays[i];
    const nm = String(row[1] == null ? '' : row[1]).trim();
    if (!nm) continue;
    map[nm] = row;
  }
  const out = [];
  const keys = Object.keys(map);
  for (let k = 0; k < keys.length; k++) {
    out.push(map[keys[k]]);
  }
  return out;
}

/** ERP출처 + 거래처명 기준 키 (9열 요약 행) — 추후 중복 병합 시 사용 예정 */
function summaryMergeKey_(row) {
  return String(row[0] == null ? '' : row[0]).trim() + '\t' + String(row[1] == null ? '' : row[1]).trim();
}

/**
 * 잔액명세 요약(balance) + 원장 요약(raw) 병합.
 * (고급 dedupe는 추후) — 현재는 단순 concat 만 필요 시 importLedgerFiles 에서 직접 호출.
 */
function mergeSummaryRows_(balanceRows, rawRows) {
  return (balanceRows || []).concat(rawRows || []);
}

/** 관리메모에 ERP 표기 (API·구 시트 호환). 기존에 ERP: 가 있으면 유지 */
function combineSummaryMemoErp_(erpLabel, existingMemo) {
  const e = 'ERP: ' + String(erpLabel == null ? '' : erpLabel).trim();
  const m = String(existingMemo == null ? '' : existingMemo).trim();
  if (!m) return e;
  if (m.indexOf('ERP:') !== -1) return m;
  return e + ' / ' + m;
}

/** 원본파일명으로 원장 ERP 출처 추정 (xlsx→이카운트, csv→경영박사) */
function inferLedgerErpSourceFromRawRow_(row) {
  const fn = String(row[RAW_LEDGER_COL_.file] == null ? '' : row[RAW_LEDGER_COL_.file])
    .trim()
    .toLowerCase();
  if (fn.endsWith('.xlsx')) return '이카운트';
  if (fn.endsWith('.csv')) return '경영박사';
  return '원장';
}

/** RAW 원장 행에서 열 인덱스 접근 (가변 길이·빈 꼬리 열 대응) */
function rawLedgerCell_(row, idx) {
  if (!row || idx < 0) return '';
  return idx < row.length ? row[idx] : '';
}

/**
 * RAW 원장 행 배열(헤더 없음) → 거래처요약 데이터 행만 반환 (9열, 시트 미기록).
 * 거래처명+ERP(원본파일 기준)별로 그룹화해 합계·마지막잔액 로직은 buildLedgerSummary 와 동일.
 */
function buildLedgerSummaryRowsFromRawOutput_(rawOutput) {
  if (!rawOutput || !rawOutput.length) {
    Logger.log('[buildRawSummary] rawOutput 비어 있음');
    return [];
  }

  const IN = RAW_LEDGER_COL_;
  Logger.log(
    '[buildRawSummary] rawOutput.length=' +
      rawOutput.length +
      ' firstRowLen=' +
      (rawOutput[0] ? rawOutput[0].length : 0) +
      ' firstRowSample=' +
      JSON.stringify(rawOutput[0])
  );

  const byKey = {};
  const order = [];
  let skippedNoName = 0;

  for (let r = 0; r < rawOutput.length; r++) {
    const row = rawOutput[r];
    if (!row) continue;
    const name = String(rawLedgerCell_(row, IN.name)).trim();
    if (!name) {
      skippedNoName++;
      continue;
    }
    const erp = inferLedgerErpSourceFromRawRow_(row);
    const key = erp + '\t' + name;
    if (!byKey[key]) {
      byKey[key] = [];
      order.push(key);
    }
    byKey[key].push(row);
  }

  if (skippedNoName) {
    Logger.log('[buildRawSummary] skipped rows (거래처명 없음)=' + skippedNoName);
  }

  const out = [];
  for (let i = 0; i < order.length; i++) {
    const key = order[i];
    const list = byKey[key];
    const erp = inferLedgerErpSourceFromRawRow_(list[0]);
    const name = String(rawLedgerCell_(list[0], IN.name)).trim();

    let sumRow = null;
    for (let j = 0; j < list.length; j++) {
      if (String(rawLedgerCell_(list[j], IN.type)).trim() === '합계') {
        sumRow = list[j];
        break;
      }
    }

    let totalSales = 0;
    let totalColl = 0;
    // totalSales/totalColl: 합계 행 우선(기간 합계), 없으면 마지막 행
    if (sumRow) {
      totalSales = toNumber_(rawLedgerCell_(sumRow, IN.sales));
      totalColl = toNumber_(rawLedgerCell_(sumRow, IN.coll));
    } else {
      const last = list[list.length - 1];
      totalSales = toNumber_(rawLedgerCell_(last, IN.sales));
      totalColl = toNumber_(rawLedgerCell_(last, IN.coll));
    }
    // lastBal: 실제 거래 행 중 balance != 0 인 마지막 값을 사용.
    // 합계/월계 행, '현재잔액/잔액/이월' 성격 품목명, balance=0 행은 후보에서 제외.
    const _isTaeyang = (name === '태양산업사');
    let lastBal = 0;
    for (let j = list.length - 1; j >= 0; j--) {
      const t = String(rawLedgerCell_(list[j], IN.type)).trim();
      const item = String(rawLedgerCell_(list[j], 4)).replace(/\s+/g, '');
      const bal = toNumber_(rawLedgerCell_(list[j], IN.bal));
      if (t === '합계' || t === '월계') {
        if (_isTaeyang) Logger.log('[태양산업사] 제외(합계/월계) idx=' + j + ' item=' + item + ' bal=' + bal);
        continue;
      }
      if (/현재잔액|잔액|이월/.test(item)) {
        if (_isTaeyang) Logger.log('[태양산업사] 제외(잔액성격품목) idx=' + j + ' item=' + item + ' bal=' + bal);
        continue;
      }
      if (bal === 0) {
        if (_isTaeyang) Logger.log('[태양산업사] 제외(bal=0) idx=' + j + ' item=' + item);
        continue;
      }
      lastBal = bal;
      if (_isTaeyang) Logger.log('[태양산업사] lastBal 확정 idx=' + j + ' item=' + item + ' bal=' + bal);
      break;
    }
    // 1차 탐색 실패 시: balance != 0 인 행 아무거나 (합계 행 포함)
    if (lastBal === 0) {
      for (let j = list.length - 1; j >= 0; j--) {
        const bal = toNumber_(rawLedgerCell_(list[j], IN.bal));
        if (bal !== 0) {
          lastBal = bal;
          if (_isTaeyang) Logger.log('[태양산업사] lastBal 폴백 idx=' + j + ' bal=' + bal);
          break;
        }
      }
    }
    if (_isTaeyang) Logger.log('[태양산업사] 최종 lastBal=' + lastBal + ' listLen=' + list.length);

    let lastPay = '';
    let tradeCount = 0;
    for (let j = 0; j < list.length; j++) {
      const row = list[j];
      if (String(rawLedgerCell_(row, IN.type)).trim() === '거래') {
        tradeCount++;
        if (toNumber_(rawLedgerCell_(row, IN.coll)) > 0) {
          lastPay = rawLedgerCell_(row, IN.date);
        }
      }
    }

    const memo = combineSummaryMemoErp_(erp, '');
    out.push([
      erp,
      name,
      totalSales,
      totalColl,
      lastBal,
      lastPay,
      tradeCount,
      getReceivableRiskGrade_(lastBal, lastPay),
      memo
    ]);
  }

  Logger.log('[buildRawSummary] summary row count=' + out.length);
  for (let s = 0; s < Math.min(3, out.length); s++) {
    Logger.log('[buildRawSummary] out[' + s + ']=' + JSON.stringify(out[s]));
  }

  return out;
}

/**
 * 경영박사 거래처별 잔액 명세 CSV — 앱 메인 데이터. 거래처요약 행 배열로 직접 변환(판매액·수금액·잔액 원값).
 */
function parseKyungyoungDoctorBalanceCsv_(file) {
  const fileName = file.getName();
  const text = readLedgerCsvBlobAsString_(file.getBlob());
  const rows = splitCsvStringToRows_(text);

  if (!rows.length) {
    Logger.log(fileName + ' : 잔액명세 CSV 행이 없습니다.');
    return [];
  }

  const headerIdx = findBalanceCsvHeaderRowIndex_(rows);
  if (headerIdx === -1) {
    Logger.log(fileName + ' : 잔액명세 헤더를 찾지 못했습니다.');
    return [];
  }

  const headerRow = rows[headerIdx];
  const colName = findBalanceCsvColumnIndex_(headerRow, ['상호명', '거래처명', '거래처']);
  const colSale = findBalanceCsvColumnIndex_(headerRow, ['판매액', '판매']);
  const colColl = findBalanceCsvColumnIndex_(headerRow, ['수금액', '수금']);
  const colBal = findBalanceCsvColumnIndex_(headerRow, ['잔액']);
  const colPhone = findBalanceCsvColumnIndex_(headerRow, ['전화', '전화번호']);

  /*
   * 전기이월·구매액·지급액: 미수 대시보드에는 넣지 않음.
   * 향후 미지급·매입처 관리 시 아래처럼 열을 매핑해 별도 시트에 보존할 수 있음.
   * const colOpen = findCsvColExact_(headerRow, '전기이월');
   * const colBuy = findCsvColExact_(headerRow, '구매액');
   * const colPay = findCsvColExact_(headerRow, '지급액');
   */

  if (colName === -1 || colSale === -1 || colColl === -1 || colBal === -1) {
    Logger.log(fileName + ' : 잔액명세 필수 컬럼 매핑 실패');
    return [];
  }

  const out = [];

  for (let r = headerIdx + 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row || !row.join('').trim()) continue;

    const name = colName < row.length ? String(row[colName] == null ? '' : row[colName]).trim() : '';
    const normalizedName = String(name || '').replace(/\s+/g, '').trim();
    if (
      !normalizedName ||
      /^총계/i.test(normalizedName) ||
      /^합계/i.test(normalizedName) ||
      /^소계/i.test(normalizedName) ||
      /^누계/i.test(normalizedName)
    ) {
      continue;
    }

    const sales = colSale >= row.length ? 0 : parseNumericForLedgerCell_(row[colSale]);
    const collect = colColl >= row.length ? 0 : parseNumericForLedgerCell_(row[colColl]);
    const balance = colBal >= row.length ? 0 : parseNumericForLedgerCell_(row[colBal]);
    const phone =
      colPhone === -1 || colPhone >= row.length
        ? ''
        : String(row[colPhone] == null ? '' : row[colPhone]).trim();
    const memoBase = phone ? '연락처: ' + phone : '';
    const memo = combineSummaryMemoErp_('경영박사', memoBase);

    const risk = getReceivableRiskGrade_(balance, '');

    out.push(['경영박사', name, sales, collect, balance, '', 0, risk, memo]);
  }

  Logger.log(
    '[BalanceCSV] ' +
      fileName +
      ' headerRow(1-based)=' +
      (headerIdx + 1) +
      ' rows=' +
      out.length
  );
  if (out.length) {
    Logger.log('[BalanceCSV] firstRowSample: ' + JSON.stringify(out[0]));
  }

  return out;
}

/**
 * 거래처요약 정렬: ①미수(잔액>0) → 0잔액 → 음수 ②잔액 내림차순 ③ERP출처 ④거래처명(ko)
 */
function compareLedgerSummaryRowsForSheet_(a, b) {
  const S = SUMMARY_SHEET_COL_;
  const balA = toNumber_(a[S.balance]);
  const balB = toNumber_(b[S.balance]);
  const groupA = balA > 0 ? 0 : balA === 0 ? 1 : 2;
  const groupB = balB > 0 ? 0 : balB === 0 ? 1 : 2;
  if (groupA !== groupB) return groupA - groupB;
  if (balB !== balA) return balB - balA;
  const erpA = String(a[S.erp] == null ? '' : a[S.erp]);
  const erpB = String(b[S.erp] == null ? '' : b[S.erp]);
  const erpCmp = erpA.localeCompare(erpB, 'ko');
  if (erpCmp !== 0) return erpCmp;
  const nmA = String(a[S.name] == null ? '' : a[S.name]);
  const nmB = String(b[S.name] == null ? '' : b[S.name]);
  return nmA.localeCompare(nmB, 'ko');
}

/**
 * 거래처요약 시트 일괄 기록 (setValues 1회, 잔여 행만 clear — clearContents/autoResize 없음)
 */
function applySummarySheetOutput_(summary, output) {
  const numRows = output.length;
  const numCols = output[0].length;
  const prevLastRow = summary.getLastRow();

  summary.getRange(1, 1, numRows, numCols).setValues(output);

  if (prevLastRow > numRows) {
    summary.getRange(numRows + 1, 1, prevLastRow, numCols).clearContent();
  }

  summary.getRange(1, 1, 1, numCols)
    .setBackground('#0b5394')
    .setFontColor('white')
    .setFontWeight('bold');

  if (numRows > 1) {
    summary.getRange(2, 3, numRows, 5).setNumberFormat('#,##0');
  }
}

function writeLedgerSummaryFromBalanceRows_(ss, dataRows) {
  const summary = ss.getSheetByName('거래처요약') || ss.insertSheet('거래처요약');

  const rows = (dataRows || []).slice();
  Logger.log('[writeSummary] 정렬 전 건수=' + rows.length);
  const sorted = rows.sort(compareLedgerSummaryRowsForSheet_);
  for (let di = 0; di < Math.min(5, sorted.length); di++) {
    Logger.log('[writeSummary] sorted[' + di + '] name=' + sorted[di][1] + ' bal=' + sorted[di][4]);
  }
  const output = SUMMARY_SHEET_HEADERS_.concat(sorted);

  applySummarySheetOutput_(summary, output);

  Logger.log('[writeSummary] 완료 rows=' + sorted.length);
}

function findKyCsvHeaderRowIndex_(rows) {
  const legacyKeywords = [
    '일자-No.',
    '일자-No',
    '품목명',
    '수량',
    '단가',
    '금액',
    '부가세',
    '판매',
    '수금',
    '잔액'
  ];

  function cellsTrimmed_(row) {
    return row.map(function (c) {
      return String(c == null ? '' : c)
        .replace(/^\uFEFF/, '')
        .trim();
    });
  }

  function rowHasExactCell_(cells, exact) {
    for (let j = 0; j < cells.length; j++) {
      if (cells[j] === exact) return true;
    }
    return false;
  }

  function rowCellMatchesSalesOut_(cells) {
    for (let j = 0; j < cells.length; j++) {
      const c = cells[j].replace(/\s/g, '');
      if (c === '매출/출금' || (c.indexOf('매출') !== -1 && c.indexOf('출금') !== -1)) return true;
    }
    return false;
  }

  function rowCellMatchesPurchaseIn_(cells) {
    for (let j = 0; j < cells.length; j++) {
      const c = cells[j].replace(/\s/g, '');
      if (c === '매입/입금' || (c.indexOf('매입') !== -1 && c.indexOf('입금') !== -1)) return true;
    }
    return false;
  }

  for (let i = 0; i < rows.length; i++) {
    const cells = cellsTrimmed_(rows[i]);
    const joined = cells.join(' ');

    const kyCore =
      rowHasExactCell_(cells, '날짜') &&
      rowHasExactCell_(cells, '품명') &&
      rowHasExactCell_(cells, '수량') &&
      rowHasExactCell_(cells, '단가') &&
      rowHasExactCell_(cells, '부가세') &&
      rowHasExactCell_(cells, '잔액') &&
      rowCellMatchesSalesOut_(cells) &&
      rowCellMatchesPurchaseIn_(cells);

    if (kyCore) {
      return i;
    }

    let hit = 0;
    for (let k = 0; k < legacyKeywords.length; k++) {
      if (joined.indexOf(legacyKeywords[k]) !== -1) hit++;
    }
    if (joined.indexOf('일자') !== -1 && joined.indexOf('품목명') !== -1 && hit >= 4) {
      return i;
    }
  }
  return -1;
}

function findCsvColByHeader_(headerRow, candidates) {
  for (let c = 0; c < headerRow.length; c++) {
    const h = String(headerRow[c] == null ? '' : headerRow[c])
      .replace(/^\uFEFF/, '')
      .trim();
    for (let i = 0; i < candidates.length; i++) {
      const key = candidates[i];
      if (h === key || h.indexOf(key) !== -1) return c;
    }
  }
  return -1;
}

function findCsvColExact_(headerRow, exactName) {
  for (let c = 0; c < headerRow.length; c++) {
    const h = String(headerRow[c] == null ? '' : headerRow[c])
      .replace(/^\uFEFF/, '')
      .trim();
    if (h === exactName) return c;
  }
  return -1;
}

function findCsvColPartialBoth_(headerRow, a, b) {
  for (let c = 0; c < headerRow.length; c++) {
    const h = String(headerRow[c] == null ? '' : headerRow[c])
      .replace(/^\uFEFF/, '')
      .replace(/\s/g, '');
    if (h.indexOf(a) !== -1 && h.indexOf(b) !== -1) return c;
  }
  return -1;
}

function kyItemSpecLabel_(pum, spec) {
  const n = String(pum == null ? '' : pum).trim();
  const s = String(spec == null ? '' : spec).trim();
  if (n && s) return n + ' [' + s + ']';
  return n || s;
}

function extractCsvFallbackFileStem_(fileName) {
  const base = fileName.replace(/\.csv$/i, '').trim();
  if (/^거래처원장[_\s-]+\d{6}_\d{6}$/i.test(base)) return '';
  if (looksLikeAutoLedgerTimestampStem_(base)) return '';
  return base;
}

function looksLikeAutoLedgerTimestampStem_(name) {
  const n = String(name || '').trim();
  if (/^\d{6}_\d{6}$/.test(n)) return true;
  if (/^\d{6}-\d{6}$/.test(n)) return true;
  if (/^\d{14}$/.test(n)) return true;
  if (/^\d{8}_\d{6}$/.test(n)) return true;
  return false;
}

function extractCustomerNameFromLedgerFileName_(fileName, extPattern) {
  const base = String(fileName).replace(extPattern, '').trim();
  const m1 = base.match(/거래처원장[_\-\s]+(.+)/i);
  if (m1 && m1[1]) {
    const cap = m1[1].trim();
    if (!looksLikeAutoLedgerTimestampStem_(cap)) return cap;
  }
  const m2 = base.match(/^(.+)[_\-\s]+거래처원장$/i);
  if (m2 && m2[1]) {
    const cap = m2[1].trim();
    if (!looksLikeAutoLedgerTimestampStem_(cap)) return cap;
  }
  return '';
}

function extractCustomerNameFromCsvFirstRow_(firstRow) {
  if (!firstRow || firstRow.length < 2) return '';
  const second = String(firstRow[1] == null ? '' : firstRow[1]).trim();
  return second || '';
}

function extractCustomerNameFromCsvPreamble_(rowsBeforeHeader) {
  const lines = [];
  for (let r = 0; r < rowsBeforeHeader.length; r++) {
    const parts = rowsBeforeHeader[r].map(function (x) {
      return String(x == null ? '' : x).trim();
    });
    const nonEmpty = parts.filter(Boolean);
    if (nonEmpty.length) lines.push(nonEmpty.join(' '));
  }
  const blob = lines.join('\n');

  const patterns = [
    /거래처\s*[:：]\s*([^\n,]+)/,
    /거래처명\s*[:：]\s*([^\n,]+)/,
    /업체명\s*[:：]\s*([^\n,]+)/,
    /상호\s*[:：]\s*([^\n,]+)/
  ];
  for (let p = 0; p < patterns.length; p++) {
    const m = blob.match(patterns[p]);
    if (m && m[1]) return m[1].trim();
  }
  return '';
}

function sanitizeLedgerNumericString_(v) {
  let s = String(v == null ? '' : v)
    .replace(/,/g, '')
    .replace(/\s/g, '')
    .trim();
  s = s.replace(/^\*+/, '');
  while (s.length > 0 && /[^0-9().\-]/.test(s.charAt(0)) && s.charAt(0) !== '(') {
    s = s.slice(1);
  }
  return s;
}

function parseNumericForLedgerCell_(v) {
  if (v === null || v === undefined) return 0;
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  let s = sanitizeLedgerNumericString_(v);
  if (s === '' || s === '-') return 0;
  let neg = false;
  if (/^\(.*\)$/.test(s)) {
    neg = true;
    s = s.slice(1, -1).replace(/,/g, '').trim();
  }
  if (s.charAt(0) === '-') {
    neg = !neg;
    s = s.slice(1);
  }
  const n = Number(s);
  if (isNaN(n)) return 0;
  return neg ? -Math.abs(n) : n;
}

/**
 * Google Drive 파일(경영박사 ERP 등) CSV → RAW 행 배열
 */
function parseKyungyoungDoctorCsv_(file) {
  const fileName = file.getName();
  const blob = file.getBlob();
  const text = readLedgerCsvBlobAsString_(blob);
  const rows = parseLedgerCsvToRows_(text);

  if (!rows.length) {
    Logger.log(fileName + ' : CSV 행이 없습니다.');
    return [];
  }

  const headerIdx = findKyCsvHeaderRowIndex_(rows);
  if (headerIdx === -1) {
    Logger.log(fileName + ' : CSV 헤더를 찾지 못했습니다.');
    Logger.log('[KyCSV] headerRowIndex(1-based): (none)');
    Logger.log('[KyCSV] customerName: (n/a)');
    Logger.log('[KyCSV] outputRowCount: 0');
    Logger.log('[KyCSV] firstTradeRowSample: null');
    return [];
  }

  const headerRow = rows[headerIdx];
  const preamble = rows.slice(0, headerIdx);

  let customerName = extractCustomerNameFromLedgerFileName_(fileName, /\.csv$/i);
  if (!customerName) {
    customerName = extractCustomerNameFromCsvFirstRow_(rows[0] || []);
  }
  if (!customerName) {
    customerName = extractCustomerNameFromCsvPreamble_(preamble);
  }
  if (!customerName) {
    customerName = extractCsvFallbackFileStem_(fileName);
  }
  if (!customerName) {
    customerName = fileName.replace(/\.csv$/i, '').trim();
  }

  Logger.log('[KyCSV] headerRowIndex(1-based): ' + (headerIdx + 1));
  Logger.log('[KyCSV] header col count=' + headerRow.length);
  Logger.log('[KyCSV] header row: ' + headerRow.map(function (h) { return String(h).trim(); }).join(' | '));

  const colDateKy = findCsvColExactNormalized_(headerRow, ['날짜']);
  const colPumKy = findCsvColExactNormalized_(headerRow, ['품명']);
  const isKyDoctorLayout = colDateKy !== -1 && colPumKy !== -1;

  const out = [];

  if (isKyDoctorLayout) {
    const colGy = findCsvColExactNormalized_(headerRow, ['규격']);
    const colMemo = findCsvColExactNormalized_(headerRow, ['적요']);
    const colQty = findCsvColExactNormalized_(headerRow, ['수량']);
    const colPrice = findCsvColExactNormalized_(headerRow, ['단가']);
    const colVat = findCsvColExactNormalized_(headerRow, ['부가세']);
    const colBal = findCsvColExactNormalized_(headerRow, ['잔액']);
    let colSaleMapped = findCsvColExactNormalized_(headerRow, ['매출/출금']);
    if (colSaleMapped === -1) {
      colSaleMapped = findCsvColPartialBothNormalized_(headerRow, '매출', '출금');
    }
    let colCollMapped = findCsvColExactNormalized_(headerRow, ['매입/입금']);
    if (colCollMapped === -1) {
      colCollMapped = findCsvColPartialBothNormalized_(headerRow, '매입', '입금');
    }

    const kyIdxMap = {
      날짜: colDateKy,
      품명: colPumKy,
      규격: colGy,
      적요: colMemo,
      수량: colQty,
      단가: colPrice,
      부가세: colVat,
      잔액: colBal,
      매출출금: colSaleMapped,
      매입입금: colCollMapped
    };
    Logger.log('[KyCSV] KyDoctorLayout column indices: ' + JSON.stringify(kyIdxMap));

    if (colBal === -1 || colQty === -1 || colPrice === -1) {
      Logger.log(fileName + ' : 경영박사 CSV 필수 컬럼(수량·단가·잔액) 매핑 실패');
      Logger.log('[KyCSV] customerName: ' + customerName);
      Logger.log('[KyCSV] outputRowCount: 0');
      Logger.log('[KyCSV] firstTradeRowSample: null');
      return [];
    }

    const idxListKy = [colDateKy, colPumKy, colQty, colPrice, colVat, colBal, colSaleMapped, colCollMapped, colGy, colMemo].filter(
      function (x) {
        return x >= 0;
      }
    );
    const maxIdxKy = idxListKy.length ? Math.max.apply(null, idxListKy) : 0;

    let sampleLogged = 0;
    for (let r = headerIdx + 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || !row.join('').trim()) continue;

      if (row.length <= maxIdxKy) {
        Logger.log(
          '[KyCSV] skip short row r(1-based)=' + (r + 1) + ' len=' + row.length + ' minCols=' + (maxIdxKy + 1)
        );
        continue;
      }

      if (sampleLogged < 3) {
        Logger.log(
          '[KyCSV] sample data row r(1-based)=' +
            (r + 1) +
            ' len=' +
            row.length +
            ' cells0_7=' +
            JSON.stringify(row.slice(0, 8))
        );
        sampleLogged++;
      }

      const rawDate = colDateKy < row.length ? row[colDateKy] : '';
      const pum = colPumKy < row.length ? row[colPumKy] : '';
      const spec =
        colGy === -1 || colGy >= row.length ? '' : row[colGy];
      const itemDisp = kyItemSpecLabel_(pum, spec);
      const memo =
        colMemo === -1 || colMemo >= row.length ? '' : row[colMemo];

      const rowType = classifyLedgerRowType_(rawDate, itemDisp, memo);

      const qty = parseNumericForLedgerCell_(row[colQty]);
      const unitPrice = parseNumericForLedgerCell_(row[colPrice]);
      const vat =
        colVat === -1 || colVat >= row.length
          ? 0
          : parseNumericForLedgerCell_(row[colVat]);
      const balance = parseNumericForLedgerCell_(row[colBal]);

      /*
       * 경영박사 CSV: '매출/출금' → RAW '판매', '매입/입금' → RAW '수금' 으로 임시 매핑함.
       * 거래처·업종에 따라 매출/매입을 반대로 RAW 컬럼에 넣어야 할 수 있음 → 추후 옵션·설정으로 분기할 것.
       */
      const sales =
        colSaleMapped === -1 || colSaleMapped >= row.length
          ? 0
          : parseNumericForLedgerCell_(row[colSaleMapped]);
      const collection =
        colCollMapped === -1 || colCollMapped >= row.length
          ? 0
          : parseNumericForLedgerCell_(row[colCollMapped]);

      const amount = 0;
      const purchase = 0;
      const payment = 0;

      pushKyLedgerRawRow_(out, customerName, fileName, rowType, [
        rawDate,
        itemDisp,
        qty,
        unitPrice,
        amount,
        vat,
        purchase,
        sales,
        collection,
        payment,
        balance,
        memo
      ]);
    }
  } else {
    const m = findLedgerLegacyColumnMap_(headerRow);
    const colDate = m.colDate;
    const colItem = m.colItem;
    const colQty = m.colQty;
    const colPrice = m.colPrice;
    const colAmt = m.colAmt;
    const colVat = m.colVat;
    const colBuy = m.colBuy;
    const colSale = m.colSale;
    const colColl = m.colColl;
    const colPay = m.colPay;
    const colBal = m.colBal;
    const colMemo = m.colMemo;

    Logger.log('[KyCSV] Legacy layout column indices: ' + JSON.stringify(m));

    if (colDate === -1 || colItem === -1 || colBal === -1) {
      Logger.log(fileName + ' : CSV 필수 컬럼(일자·품목·잔액) 매핑 실패');
      Logger.log('[KyCSV] customerName: ' + customerName);
      Logger.log('[KyCSV] outputRowCount: 0');
      Logger.log('[KyCSV] firstTradeRowSample: null');
      return [];
    }

    const idxListLeg = [
      colDate,
      colItem,
      colQty,
      colPrice,
      colAmt,
      colVat,
      colBuy,
      colSale,
      colColl,
      colPay,
      colBal,
      colMemo
    ].filter(function (x) {
      return x >= 0;
    });
    const maxIdxLeg = idxListLeg.length ? Math.max.apply(null, idxListLeg) : 0;

    let sampleLoggedL = 0;
    for (let r = headerIdx + 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || !row.join('').trim()) continue;

      if (row.length <= maxIdxLeg) {
        Logger.log(
          '[KyCSV] [legacy] skip short row r(1-based)=' + (r + 1) + ' len=' + row.length + ' minCols=' + (maxIdxLeg + 1)
        );
        continue;
      }

      if (sampleLoggedL < 3) {
        Logger.log(
          '[KyCSV] [legacy] sample data row r(1-based)=' +
            (r + 1) +
            ' len=' +
            row.length +
            ' cells0_7=' +
            JSON.stringify(row.slice(0, 8))
        );
        sampleLoggedL++;
      }

      const dateNo = colDate < row.length ? row[colDate] : '';
      const itemName = colItem < row.length ? row[colItem] : '';
      const memo = colMemo === -1 ? '' : colMemo < row.length ? row[colMemo] : '';

      const rowType = classifyLedgerRowType_(dateNo, itemName, memo);

      const qty = colQty === -1 ? 0 : parseNumericForLedgerCell_(row[colQty]);
      const unitPrice = colPrice === -1 ? 0 : parseNumericForLedgerCell_(row[colPrice]);
      const amount = colAmt === -1 ? 0 : parseNumericForLedgerCell_(row[colAmt]);
      const vat = colVat === -1 ? 0 : parseNumericForLedgerCell_(row[colVat]);
      const purchase = colBuy === -1 ? 0 : parseNumericForLedgerCell_(row[colBuy]);
      const sales = colSale === -1 ? 0 : parseNumericForLedgerCell_(row[colSale]);
      const collection = colColl === -1 ? 0 : parseNumericForLedgerCell_(row[colColl]);
      const payment = colPay === -1 ? 0 : parseNumericForLedgerCell_(row[colPay]);
      const balance = parseNumericForLedgerCell_(row[colBal]);

      pushKyLedgerRawRow_(out, customerName, fileName, rowType, [
        dateNo,
        itemName,
        qty,
        unitPrice,
        amount,
        vat,
        purchase,
        sales,
        collection,
        payment,
        balance,
        memo
      ]);
    }
  }

  let firstTradeSample = null;
  for (let ti = 0; ti < out.length; ti++) {
    if (out[ti][2] === '거래') {
      firstTradeSample = out[ti];
      break;
    }
  }
  if (!firstTradeSample && out.length) {
    firstTradeSample = out[0];
  }

  Logger.log('[KyCSV] customerName: ' + customerName);
  Logger.log('[KyCSV] outputRowCount: ' + out.length);
  Logger.log('[KyCSV] firstTradeRowSample: ' + JSON.stringify(firstTradeSample));

  return out;
}

/**
 * .xls / .xlsx 잔액명세 파일 → balanceSummaryAccum 행 배열
 * (parseKyungyoungDoctorBalanceCsv_ 와 동일 출력 형식, XLS 2D 배열 기반)
 */
function parseBalanceXlsFile_(file, fileName) {
  const out = [];
  const blob = file.getBlob();
  let tempFile = null;

  try {
    tempFile = Drive.Files.create(
      { name: fileName, mimeType: MimeType.GOOGLE_SHEETS },
      blob
    );

    const tempSS = SpreadsheetApp.openById(tempFile.id);
    const tempSheet = tempSS.getSheets()[0];
    const values = tempSheet.getDataRange().getValues();

    Logger.log('[BalanceXLS] ' + fileName + ' totalRows=' + values.length);
    for (let di = 0; di < Math.min(15, values.length); di++) {
      Logger.log('[BalanceXLS] row' + di + '=' + values[di].join(' | '));
    }

    const headerIdx = findBalanceCsvHeaderRowIndex_(values);
    if (headerIdx === -1) {
      Logger.log('[BalanceXLS] ' + fileName + ' : 잔액명세 헤더를 찾지 못했습니다.');
      return out;
    }
    Logger.log('[BalanceXLS] header found at row=' + headerIdx);

    const headerRow = values[headerIdx];
    const colName = findBalanceCsvColumnIndex_(headerRow, ['상호명', '거래처명', '거래처']);
    const colSale = findBalanceCsvColumnIndex_(headerRow, ['판매액', '판매']);
    const colColl = findBalanceCsvColumnIndex_(headerRow, ['수금액', '수금']);
    const colBal  = findBalanceCsvColumnIndex_(headerRow, ['잔액']);
    const colPhone = findBalanceCsvColumnIndex_(headerRow, ['전화', '전화번호']);
    Logger.log('[BalanceXLS] colName=' + colName + ' colSale=' + colSale +
               ' colColl=' + colColl + ' colBal=' + colBal);

    if (colName === -1 || colSale === -1 || colColl === -1 || colBal === -1) {
      Logger.log('[BalanceXLS] ' + fileName + ' : 필수 컬럼 매핑 실패');
      return out;
    }

    for (let r = headerIdx + 1; r < values.length; r++) {
      const row = values[r];
      if (!row || !row.join('').trim()) continue;

      const name = colName < row.length
        ? String(row[colName] == null ? '' : row[colName]).trim() : '';
      const normalizedName = String(name || '').replace(/\s+/g, '').trim();
      if (
        !normalizedName ||
        /^총계/i.test(normalizedName) ||
        /^합계/i.test(normalizedName) ||
        /^소계/i.test(normalizedName) ||
        /^누계/i.test(normalizedName)
      ) continue;

      const sales   = colSale >= row.length ? 0 : parseNumericForLedgerCell_(row[colSale]);
      const collect = colColl >= row.length ? 0 : parseNumericForLedgerCell_(row[colColl]);
      const balance = colBal  >= row.length ? 0 : parseNumericForLedgerCell_(row[colBal]);
      const phone   = colPhone === -1 || colPhone >= row.length
        ? '' : String(row[colPhone] == null ? '' : row[colPhone]).trim();
      const memoBase = phone ? '연락처: ' + phone : '';
      const memo  = combineSummaryMemoErp_('경영박사', memoBase);
      const risk  = getReceivableRiskGrade_(balance, '');

      out.push(['경영박사', name, sales, collect, balance, '', 0, risk, memo]);
    }

    Logger.log('[BalanceXLS] ' + fileName + ' 추출 rows=' + out.length);
    if (out.length) Logger.log('[BalanceXLS] sample=' + JSON.stringify(out[0]));

  } finally {
    if (tempFile && tempFile.id) {
      try { DriveApp.getFileById(tempFile.id).setTrashed(true); }
      catch (trashErr) { Logger.log('temp xls trash: ' + trashErr); }
    }
  }

  return out;
}

/**
 * 단일 xlsx / xls 원장 파일 → RAW 행 배열 (임시 스프레드시트 변환 후 삭제)
 */
function parseLedgerXlsxFile_(file, fileName) {
  const out = [];
  const blob = file.getBlob();
  let tempFile = null;

  try {
    tempFile = Drive.Files.create(
      {
        name: fileName,
        mimeType: MimeType.GOOGLE_SHEETS
      },
      blob
    );

    const tempSS = SpreadsheetApp.openById(tempFile.id);
    const tempSheet = tempSS.getSheets()[0];
    const values = tempSheet.getDataRange().getValues();

    // 상위 15행 진단 로그
    Logger.log('[parse] ' + fileName + ' totalRows=' + values.length);
    for (let di = 0; di < Math.min(15, values.length); di++) {
      Logger.log('[parse] row' + di + '=' + values[di].join(' | '));
    }

    let startRow = -1;

    for (let i = 0; i < Math.min(20, values.length); i++) {
      const rowText = values[i].join(' ');
      Logger.log('[parse] checking header row=' + i + ' normalized=' + rowText.slice(0, 80));
      if (rowText.includes('일자-No.') && rowText.includes('품목명')) {
        startRow = i;
        Logger.log('[parse] header found at row=' + i);
        break;
      }
    }

    if (startRow === -1) {
      Logger.log('[parse] ' + fileName + ' : 헤더를 찾지 못했습니다. (원장 파서 — 잔액명세 파일은 parseBalanceXlsFile_ 로 처리)');
      return out;
    }

    let customerName = extractCustomerNameFromLedgerFileName_(fileName, /\.xlsx$/i);
    if (!customerName) customerName = fileName.replace(/\.xlsx$/i, '');

    for (let r = startRow + 1; r < values.length; r++) {
      const row = values[r];

      if (!row.join('').trim()) continue;

      const dateNo = row[0];
      const itemName = row[1];
      const qty = row[2];
      const unitPrice = row[3];
      const amount = row[4];
      const vat = row[5];
      const purchase = row[6];
      const sales = row[7];
      const collection = row[8];
      const payment = row[9];
      const balance = row[10];
      const memo = row[11];

      const rowType = classifyLedgerRowType_(dateNo, itemName, memo);

      out.push([
        customerName,
        fileName,
        rowType,
        dateNo,
        itemName,
        qty,
        unitPrice,
        amount,
        vat,
        purchase,
        sales,
        collection,
        payment,
        balance,
        memo
      ]);
    }
  } finally {
    if (tempFile && tempFile.id) {
      try {
        DriveApp.getFileById(tempFile.id).setTrashed(true);
      } catch (trashErr) {
        Logger.log('temp xlsx trash: ' + trashErr);
      }
    }
  }

  return out;
}

function importLedgerFiles() {
  const start = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = ss.getSheetByName(RAW_SHEET_NAME);

  if (!rawSheet) {
    throw new Error('RAW_거래처원장 시트가 없습니다.');
  }

  // ── 폴더 진단 로그 ──────────────────────────────────────────────────────────
  // 현재 구조: FOLDER_ID 로 지정된 폴더 바로 아래 파일만 처리 (하위 폴더 미검색)
  const folder = DriveApp.getFolderById(FOLDER_ID);
  Logger.log('[import] sourceFolderId=' + FOLDER_ID);
  Logger.log('[import] sourceFolderName=' + folder.getName());

  // 폴더 내 전체 파일 목록 (확장자 필터 적용 전)
  const fileList = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    const fName = f.getName();
    const mime = f.getMimeType();
    Logger.log('[import] found file: ' + fName + ' / mimeType=' + mime);

    const lower = fName.toLowerCase();
    const isLedgerFile =
      lower.endsWith('.csv') ||
      lower.endsWith('.xlsx') ||
      lower.endsWith('.xls');
    if (!isLedgerFile) {
      Logger.log('[import] skipped file: ' + fName);
    } else {
      Logger.log('[import] accepted file: ' + fName);
      fileList.push(f);
    }
  }
  Logger.log('[import] accepted total=' + fileList.length);
  // ────────────────────────────────────────────────────────────────────────────

  if (fileList.length === 0) {
    Logger.log('[import] 처리할 파일 없음 — 종료');
    ss.toast('폴더에 .csv / .xlsx / .xls 파일이 없습니다.\n폴더명: ' + folder.getName(), '파일 없음', 8);
    return;
  }

  rawSheet.clearContents();
  rawSheet.getRange(1, 1, 1, RAW_LEDGER_HEADERS_.length).setValues([RAW_LEDGER_HEADERS_]);

  fileList.sort(function (a, b) {
    const ab =
      a.getName().toLowerCase().endsWith('.csv') && isKyungyoungDoctorBalanceCsvFile_(a);
    const bb =
      b.getName().toLowerCase().endsWith('.csv') && isKyungyoungDoctorBalanceCsvFile_(b);
    if (ab && !bb) return -1;
    if (!ab && bb) return 1;
    return 0;
  });

  let rawOutput = [];
  let balanceSummaryAccum = [];
  let processedFiles = [];

  for (let fi = 0; fi < fileList.length; fi++) {
    const file = fileList[fi];
    const fileName = file.getName();
    const lower = fileName.toLowerCase();

    try {
      if (lower.endsWith('.csv') && isKyungyoungDoctorBalanceCsvFile_(file)) {
        // 잔액명세 CSV
        balanceSummaryAccum = balanceSummaryAccum.concat(parseKyungyoungDoctorBalanceCsv_(file));
        processedFiles.push(file);
      } else if ((lower.endsWith('.xlsx') || lower.endsWith('.xls')) && isKyungyoungDoctorBalanceCsvFile_(file)) {
        // 잔액명세 XLS/XLSX (파일명에 '거래처별잔액명세' 포함)
        Logger.log('[import] routing to BalanceXLS: ' + fileName);
        balanceSummaryAccum = balanceSummaryAccum.concat(parseBalanceXlsFile_(file, fileName));
        processedFiles.push(file);
      } else if (lower.endsWith('.xlsx') || lower.endsWith('.xls')) {
        // 원장 XLS/XLSX
        rawOutput = rawOutput.concat(parseLedgerXlsxFile_(file, fileName));
        processedFiles.push(file);
      } else if (lower.endsWith('.csv')) {
        // 원장 CSV
        rawOutput = rawOutput.concat(parseKyungyoungDoctorCsv_(file));
        processedFiles.push(file);
      }
    } catch (err) {
      Logger.log('파일 처리 실패: ' + fileName + ' / ' + (err && err.message ? err.message : String(err)));
      appendErrorLog_('파일 처리', fileName, err);
    }
  }

  Logger.log('[import] processedFiles.length=' + processedFiles.length);

  if (rawOutput.length > 0) {
    rawSheet.getRange(2, 1, rawOutput.length, rawOutput[0].length).setValues(rawOutput);
  }

  const balanceRows = mergeBalanceSummaryRowsByName_(balanceSummaryAccum);
  Logger.log('[import] rawOutput.length=' + rawOutput.length);
  const rawSummaryRows =
    rawOutput.length > 0 ? buildLedgerSummaryRowsFromRawOutput_(rawOutput) : [];
  Logger.log('[import] rawSummaryRows.length=' + rawSummaryRows.length);
  Logger.log('[import] balanceRows.length=' + balanceRows.length);
  // CSV(balanceRows) 거래처가 XLSX(rawSummaryRows)에도 있으면 XLSX 행을 제거 — CSV 잔액명세가 우선
  const balanceNameSet = {};
  balanceRows.forEach(function(r) { balanceNameSet[String(r[1]).trim()] = true; });
  const dedupedRawSummaryRows = rawSummaryRows.filter(function(r) {
    return !balanceNameSet[String(r[1]).trim()];
  });
  const mergedSummaryRows = balanceRows.concat(dedupedRawSummaryRows);
  Logger.log('[import] dedupedRaw.length=' + dedupedRawSummaryRows.length + ' mergedSummaryRows.length=' + mergedSummaryRows.length);
  for (let si = 0; si < Math.min(3, rawSummaryRows.length); si++) {
    Logger.log('[import] rawSummaryRows[' + si + ']=' + JSON.stringify(rawSummaryRows[si]));
  }

  writeLedgerSummaryFromBalanceRows_(ss, mergedSummaryRows);

  // 파일 이동은 moveProcessedFilesOnly() 로 분리 — 여기서는 로그만 기록
  Logger.log('[skip] 처리완료 폴더 이동 생략 files=' + processedFiles.length);
  processedFiles.forEach(function(f) {
    Logger.log('[skip] 이동 대기 파일: ' + f.getName());
  });
  // moveProcessedFiles_(processedFiles);

  SpreadsheetApp.flush();

  const end = Date.now();
  Logger.log('[main] 완료');
  Logger.log('[main] elapsed=' + ((end - start) / 1000).toFixed(1) + 's');

  // toast: 비블로킹 알림 (alert와 달리 사용자 클릭 대기 없이 즉시 반환)
  ss.toast(
    'RAW ' + rawOutput.length + '행 · 거래처요약 ' + mergedSummaryRows.length + '건 완료\n' +
    '파일 이동은 메뉴 → [파일 이동 실행]',
    '가져오기 완료',
    10
  );

  return;
}

/**
 * 거래처원장 RAW만 있을 때 거래처요약 생성.
 * 잔액명세+원장 병합은 importLedgerFiles()에서 writeLedgerSummaryFromBalanceRows_ 로 처리.
 * importLedgerFiles()는 요약만 갱신. 대시보드·관리대상은 runReportsFromSummary() 또는 메뉴에서 별도 실행.
 */
function buildLedgerSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName('RAW_거래처원장');
  const summary = ss.getSheetByName('거래처요약') || ss.insertSheet('거래처요약');

  if (!raw) throw new Error('RAW_거래처원장 시트가 없습니다.');

  const values = raw.getDataRange().getValues();
  if (values.length < 2) throw new Error('RAW_거래처원장 데이터가 없습니다.');

  const summaryRows = buildLedgerSummaryRowsFromRawOutput_(values.slice(1));
  const sorted = summaryRows.slice().sort(compareLedgerSummaryRowsForSheet_);
  const output = SUMMARY_SHEET_HEADERS_.concat(sorted);

  applySummarySheetOutput_(summary, output);

  SpreadsheetApp.getUi().alert('거래처요약 생성 완료(원장 기준, 합계·원값)');
}

function buildCEOdashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = ss.getSheetByName('거래처요약');
  const dash = ss.getSheetByName('대표대시보드') || ss.insertSheet('대표대시보드');

  if (!summary) throw new Error('거래처요약 시트가 없습니다.');

  const values = summary.getDataRange().getValues();
  if (values.length < 2) {
    SpreadsheetApp.getUi().alert('거래처요약에 데이터 행이 없습니다. 대표대시보드를 건너뜁니다.');
    return;
  }

  const header = values[0].map(h => String(h).trim());
  const rows = values.slice(1);

  const idxName = findColumnIndex_(header, ['거래처명']);
  const idxSales = findColumnIndex_(header, ['총판매']);
  const idxCollect = findColumnIndex_(header, ['총수금']);
  const idxBalance = findColumnIndex_(header, ['현재잔액']);
  const idxLastPay = findColumnIndex_(header, ['최근수금일']);
  const idxRisk = findColumnIndex_(header, ['위험등급']);

  if (
    idxName === -1 ||
    idxBalance === -1 ||
    idxLastPay === -1 ||
    idxRisk === -1
  ) {
    SpreadsheetApp.getUi().alert('거래처요약 시트 헤더명을 확인하세요.');
    return;
  }

  const rec = receivableSummaryRows_(rows, idxBalance);

  const totalBalance = rec.reduce(function (s, r) {
    return s + toNumber_(r[idxBalance]);
  }, 0);
  const totalSales =
    idxSales === -1 ? 0 : rec.reduce(function (s, r) { return s + toNumber_(r[idxSales]); }, 0);
  const totalCollect =
    idxCollect === -1 ? 0 : rec.reduce(function (s, r) { return s + toNumber_(r[idxCollect]); }, 0);
  const collectionRate = totalSales > 0 ? totalCollect / totalSales : 0;

  const riskTroubleCount = rec.filter(function (r) {
    return ['매우위험', '위험'].includes(String(r[idxRisk]).trim());
  }).length;

  const manageCount = rec.filter(function (r) {
    return ['매우위험', '위험', '주의'].includes(String(r[idxRisk]).trim());
  }).length;

  const today = new Date();
  let over30Count = 0;
  let over60Count = 0;
  rec.forEach(function (row) {
    const lastPay = String(row[idxLastPay]).trim();
    if (!hasUsableLastPayForAging_(lastPay)) return;
    const diffDays = getDaysFromLastPay_(lastPay, today);
    if (diffDays >= 30) over30Count++;
    if (diffDays >= 60) over60Count++;
  });

  const topRows = rec
    .slice()
    .sort(function (a, b) {
      return toNumber_(b[idxBalance]) - toNumber_(a[idxBalance]);
    })
    .slice(0, 5);

  const output = [
    ['지점 미수 및 현금흐름 위험관리 대시보드', '', '', ''],
    ['(거래처별 잔액 명세 기준 KPI)', '', '', ''],
    ['총 미수잔액(잔액>0)', totalBalance, '', ''],
    ['위험 거래처 수(매우위험+위험)', riskTroubleCount, '', ''],
    ['회수율(미수 거래처)', collectionRate, '', ''],
    ['30일 이상 수금 없음(수금일 있는 미수만)', over30Count, '', ''],
    ['60일 이상 수금 없음(수금일 있는 미수만)', over60Count, '', ''],
    ['관리대상(미수: 매우위험+위험+주의)', manageCount, '', ''],
    ['', '', '', ''],
    ['TOP 5 미수 거래처', '', '', ''],
    ['거래처명', '현재잔액', '최근수금일', '위험등급']
  ];

  topRows.forEach(function (r) {
    output.push([
      r[idxName],
      toNumber_(r[idxBalance]),
      r[idxLastPay],
      r[idxRisk]
    ]);
  });

  dash.clearContents();
  dash.clearFormats();
  dash.getDataRange().clearDataValidations();

  dash.getRange(1, 1, output.length, 4).setValues(output);

  dash.getRange('A1:D1')
    .merge()
    .setBackground('#0b5394')
    .setFontColor('white')
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('center');

  dash.getRange('A3:A9').setFontWeight('bold');
  const topTitle0 = output.findIndex(function (row) {
    return String(row[0]) === 'TOP 5 미수 거래처';
  });
  const topTitleSheetRow = topTitle0 >= 0 ? topTitle0 + 1 : 10;
  dash.getRange(topTitleSheetRow, 1, topTitleSheetRow, 4)
    .merge()
    .setBackground('#134f5c')
    .setFontColor('white')
    .setFontWeight('bold');
  dash.getRange(topTitleSheetRow + 1, 1, topTitleSheetRow + 1, 4)
    .setBackground('#0b5394')
    .setFontColor('white')
    .setFontWeight('bold');

  dash.getRange(3, 2, 3, 2).setNumberFormat('#,##0');
  dash.getRange(4, 2, 4, 2).setNumberFormat('#,##0');
  dash.getRange(5, 2, 5, 2).setNumberFormat('0.0%');
  dash.getRange(6, 2, 8, 2).setNumberFormat('#,##0');

  if (topRows.length > 0) {
    const dataStart = topTitleSheetRow + 2;
    dash
      .getRange(dataStart, 2, dataStart + topRows.length - 1, 2)
      .setNumberFormat('#,##0');
  }

  dash.autoResizeColumns(1, 4);

  SpreadsheetApp.getUi().alert('대표대시보드 생성 완료');
}

function applyRiskColorFormatting() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  applyRiskFormatToSheet_(ss, '거래처요약');
  applyRiskFormatToSheet_(ss, '대표대시보드');
  applyRiskFormatToSheet_(ss, '관리대상');

  SpreadsheetApp.getUi().alert('위험등급 색상 자동화가 적용되었습니다.');
}

function applyRiskFormatToSheet_(ss, sheetName) {
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return;

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return;

  const header = values[0];
  const riskColIndex = header.findIndex(h => String(h).trim() === '위험등급');

  if (riskColIndex === -1) return;

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();

  if (lastRow < 2) return;

  sh.clearConditionalFormatRules();

  const dataRange = sh.getRange(2, 1, lastRow - 1, lastCol);
  const riskColLetter = columnToLetter_(riskColIndex + 1);

  const rules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${riskColLetter}2="매우위험"`)
      .setBackground('#f4cccc')
      .setFontColor('#990000')
      .setRanges([dataRange])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${riskColLetter}2="위험"`)
      .setBackground('#fce5cd')
      .setFontColor('#b45f06')
      .setRanges([dataRange])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${riskColLetter}2="주의"`)
      .setBackground('#fff2cc')
      .setFontColor('#7f6000')
      .setRanges([dataRange])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${riskColLetter}2="관찰"`)
      .setBackground('#cfe2f3')
      .setFontColor('#073763')
      .setRanges([dataRange])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${riskColLetter}2="정상"`)
      .setBackground('#d9ead3')
      .setFontColor('#274e13')
      .setRanges([dataRange])
      .build(),

    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${riskColLetter}2="미수외(음수)"`)
      .setBackground('#efefef')
      .setFontColor('#434343')
      .setRanges([dataRange])
      .build()
  ];

  sh.setConditionalFormatRules(rules);
}

function generateManagementList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const summarySheet = ss.getSheetByName('거래처요약');
  const targetSheet = ss.getSheetByName('관리대상') || ss.insertSheet('관리대상');

  const data = summarySheet.getDataRange().getValues();

  if (data.length < 2) return;

  const header = data[0].map(h => String(h).trim());

  const idxName = findColumnIndex_(header, ['거래처명']);
  const idxBalance = findColumnIndex_(header, ['현재잔액', '현재 잔액']);
  const idxLastPay = findColumnIndex_(header, ['최근수금일', '최근 수금일']);
  const idxRisk = findColumnIndex_(header, ['위험등급', '위험 등급']);

  if (
    idxName === -1 ||
    idxBalance === -1 ||
    idxLastPay === -1 ||
    idxRisk === -1
  ) {
    SpreadsheetApp.getUi().alert('거래처요약 시트 헤더명을 확인하세요.');
    return;
  }

  const result = [];

  result.push([
    '거래처명',
    '위험등급',
    '현재잔액',
    '최근수금일',
    '담당자',
    '관리상태',
    '다음조치',
    '메모'
  ]);

  const dataRows = data.slice(1);
  const receivableRows = dataRows.filter(function (row) {
    return toNumber_(row[idxBalance]) > 0;
  });

  for (let i = 0; i < receivableRows.length; i++) {
    const row = receivableRows[i];

    const name = row[idxName];
    const balance = toNumber_(row[idxBalance]);
    const lastPay = row[idxLastPay];
    const risk = String(row[idxRisk]).trim();

    if (!name) continue;
    if (risk === '미수외(음수)') continue;
    if (risk === '정상' || risk === '관찰') continue;

    let manageStatus = '';
    let nextAction = '';

    if (risk === '매우위험') {
      manageStatus = '대표관리';
      nextAction = '즉시 연락';
    } else if (risk === '위험') {
      manageStatus = '집중관리';
      nextAction = '입금 확인';
    } else {
      manageStatus = '일반관리';
      nextAction = '경과 모니터링';
    }

    result.push([
      name,
      risk,
      balance,
      lastPay,
      '',
      manageStatus,
      nextAction,
      ''
    ]);
  }

  targetSheet.clearContents();
  targetSheet.clearFormats();
  targetSheet.getDataRange().clearDataValidations();

  targetSheet
    .getRange(1, 1, result.length, result[0].length)
    .setValues(result);

  targetSheet.getRange(1, 1, 1, 8)
    .setBackground('#0b5394')
    .setFontColor('white')
    .setFontWeight('bold');

  if (targetSheet.getLastRow() > 1) {
    targetSheet
      .getRange(2, 3, targetSheet.getLastRow() - 1, 1)
      .setNumberFormat('#,##0');
  }

  targetSheet.autoResizeColumns(1, 8);
}

function showCustomerDetail(customerName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const summarySheet = ss.getSheetByName('거래처요약');
  const rawSheet = ss.getSheetByName('RAW_거래처원장');
  const detailSheet = ss.getSheetByName('거래처상세') || ss.insertSheet('거래처상세');

  const summaryData = summarySheet.getDataRange().getValues();
  const rawData = rawSheet.getDataRange().getValues();

  const summaryHeader = summaryData[0].map(h => String(h).trim());

  const idxName = findColumnIndex_(summaryHeader, ['거래처명']);
  const idxRisk = findColumnIndex_(summaryHeader, ['위험등급']);
  const idxBalance = findColumnIndex_(summaryHeader, ['현재잔액']);
  const idxLastPay = findColumnIndex_(summaryHeader, ['최근수금일']);
  const idxSale = findColumnIndex_(summaryHeader, ['총판매']);
  const idxCollect = findColumnIndex_(summaryHeader, ['총수금']);

  let targetRow = null;

  for (let i = 1; i < summaryData.length; i++) {
    if (String(summaryData[i][idxName]).trim() === customerName) {
      targetRow = summaryData[i];
      break;
    }
  }

  if (!targetRow) {
    SpreadsheetApp.getUi().alert('거래처를 찾을 수 없습니다.');
    return;
  }

  detailSheet.clearContents();
  detailSheet.clearFormats();

  const lastPayRaw = String(targetRow[idxLastPay]).trim();
  const todayDetail = new Date();
  let riskAlert = '정상';
  let riskLevel = 'LOW';

  if (hasUsableLastPayForAging_(lastPayRaw)) {
    const diffDays = getDaysFromLastPay_(lastPayRaw, todayDetail);
    if (diffDays >= 60) {
      riskAlert = '60일 이상 수금 없음';
      riskLevel = 'HIGH';
    } else if (diffDays >= 30) {
      riskAlert = '30일 이상 수금 없음';
      riskLevel = 'MEDIUM';
    } else if (diffDays >= 14) {
      riskAlert = '14일 이상 수금 없음';
      riskLevel = 'WATCH';
    }
  } else if (lastPayRaw) {
    riskAlert = '최근수금일 해석 불가';
    riskLevel = 'LOW';
  } else {
    riskAlert = '최근수금일 없음(연체일수 미산정)';
    riskLevel = 'LOW';
  }

  const info = [
    ['거래처명', customerName],
    ['위험등급', targetRow[idxRisk]],
    ['위험알림', riskAlert],
    ['위험수준', riskLevel],
    ['현재잔액', targetRow[idxBalance]],
    ['최근수금일', targetRow[idxLastPay]],
    ['총판매', targetRow[idxSale]],
    ['총수금', targetRow[idxCollect]],
    ['관리상태', ''],
    ['다음조치', ''],
    ['메모', '']
  ];

  detailSheet
    .getRange(1, 1, info.length, 2)
    .setValues(info);

  detailSheet
    .getRange(1, 1, info.length, 1)
    .setFontWeight('bold')
    .setBackground('#d9eaf7');

  detailSheet.getRange(5, 2, 3, 1).setNumberFormat('#,##0');

  detailSheet.getRange(14, 1)
    .setValue('최근 거래내역')
    .setFontWeight('bold');

  const recentRows = [];

  for (let i = 1; i < rawData.length; i++) {
    const row = rawData[i];

    if (String(row[0]).trim() === customerName) {
      recentRows.push(row);
    }
  }

  if (recentRows.length > 0) {
    const recent = recentRows.slice(-20);
    const cleanedRows = [];

    cleanedRows.push([
      '거래일',
      '구분',
      '내용',
      '금액',
      '누적잔액'
    ]);

    for (let i = 0; i < recent.length; i++) {
      const r = recent[i];

      const type = String(r[2]).trim();
      const content = String(r[4]).trim();

      if (
        type === '월계' ||
        type === '합계' ||
        content === '' ||
        content === '합계'
      ) {
        continue;
      }

      cleanedRows.push([
        r[3],
        r[2],
        r[4],
        toNumber_(r[7]),
        toNumber_(r[10])
      ]);
    }

    detailSheet
      .getRange(15, 1, cleanedRows.length, cleanedRows[0].length)
      .setValues(cleanedRows);

    detailSheet
      .getRange(15, 1, 1, 5)
      .setBackground('#0b5394')
      .setFontColor('white')
      .setFontWeight('bold');

    if (cleanedRows.length > 1) {
      detailSheet
        .getRange(16, 4, cleanedRows.length - 1, 2)
        .setNumberFormat('#,##0');
    }
  }

  detailSheet.autoResizeColumns(1, 5);
}

function testShowCustomerDetail() {
  showCustomerDetail('태양산업사');
}

function testImportLedgerFiles() {
  importLedgerFiles();
}

function toNumber_(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  let s = sanitizeLedgerNumericString_(v);
  if (s === '' || s === '-') return 0;
  let neg = false;
  if (/^\(.*\)$/.test(s)) {
    neg = true;
    s = s.slice(1, -1).replace(/,/g, '').trim();
  }
  if (s.charAt(0) === '-') {
    neg = !neg;
    s = s.slice(1);
  }
  const n = Number(s);
  if (isNaN(n)) return 0;
  return neg ? -Math.abs(n) : n;
}

function getRiskLevel_(balance, lastCollectionDate) {
  if (balance >= 30000000) return '매우위험';
  if (balance >= 10000000) return '위험';
  if (balance >= 3000000) return '주의';
  if (balance > 0) return '관찰';
  return '정상';
}

/**
 * 미수(잔액>0) 위험등급. 잔액 0·음수는 미수 위험 KPI·목록에서 제외·별도 구분에 사용.
 */
function getReceivableRiskGrade_(balance, lastCollectionDate) {
  const b = toNumber_(balance);
  if (b < 0) return '미수외(음수)';
  if (b === 0) return '정상';
  return getRiskLevel_(b, lastCollectionDate);
}

function getDaysFromLastPay_(lastPay, today) {
  if (!lastPay) return 9999;

  try {
    const text = String(lastPay).trim();

    if (!text || text === '-') return 9999;

    const parsed = parseLedgerDateNo_(text, today);
    if (parsed && !isNaN(parsed.getTime())) {
      return Math.floor((today - parsed) / (1000 * 60 * 60 * 24));
    }

    return 9999;
  } catch (e) {
    return 9999;
  }
}

function findColumnIndex_(header, candidates) {
  for (const candidate of candidates) {
    const idx = header.indexOf(candidate);
    if (idx !== -1) return idx;
  }

  return -1;
}

function columnToLetter_(column) {
  let temp;
  let letter = '';

  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }

  return letter;
}
function showSelectedCustomerDetail() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const detailSheet = ss.getSheetByName('거래처상세');

  const customerName = String(detailSheet.getRange('B1').getValue()).trim();

  if (!customerName) {
    SpreadsheetApp.getUi().alert('거래처상세 시트 B1에 거래처명을 입력하세요.');
    return;
  }

  showCustomerDetail(customerName);
}

// ═══════════════════════════════════════════════════════════
// Web App JSON API — GitHub Pages 앱에서 GET으로 호출
//   ?action=dashboard | ?action=customers | ?action=detail&name=거래처명
// ═══════════════════════════════════════════════════════════

/** 회수 액션 로그 — Web API 전용 (ETL와 무관) */
var ACTION_LOG_SHEET_NAME_ = 'ACTION_LOG';
var ACTION_LOG_HEADERS_ = ['일시', '거래처명', '조치내용', '담당자', '다음예정일'];

function ensureActionLogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ACTION_LOG_SHEET_NAME_);
  if (!sh) {
    sh = ss.insertSheet(ACTION_LOG_SHEET_NAME_);
    sh.getRange(1, 1, 1, ACTION_LOG_HEADERS_.length).setValues([ACTION_LOG_HEADERS_]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function apiSaveAction_(customer, memo, manager, nextDate) {
  var c = String(customer == null ? '' : customer).trim();
  var m = String(memo == null ? '' : memo).trim();
  if (!c) throw new Error('거래처명(customer)이 필요합니다.');
  if (!m) throw new Error('조치내용(memo)이 필요합니다.');
  var sh = ensureActionLogSheet_();
  var tz = Session.getScriptTimeZone() || 'Asia/Seoul';
  var nowStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');
  var mgr = manager == null ? '' : String(manager);
  var nd = nextDate == null ? '' : String(nextDate);
  sh.appendRow([nowStr, c, m, mgr, nd]);
  return { success: true };
}

function apiGetActions_(customerName) {
  var cn = String(customerName == null ? '' : customerName).trim();
  if (!cn) return [];
  ensureActionLogSheet_();
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ACTION_LOG_SHEET_NAME_);
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  var header = values[0].map(function (h) {
    return String(h).trim();
  });
  var iTime = findColumnIndex_(header, ['일시']);
  var iName = findColumnIndex_(header, ['거래처명']);
  var iMemo = findColumnIndex_(header, ['조치내용']);
  var iMgr = findColumnIndex_(header, ['담당자']);
  var iNext = findColumnIndex_(header, ['다음예정일']);
  if (iTime === -1) iTime = 0;
  if (iName === -1) iName = 1;
  if (iMemo === -1) iMemo = 2;
  if (iMgr === -1) iMgr = 3;
  if (iNext === -1) iNext = 4;
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var nm = String(row[iName] == null ? '' : row[iName]).trim();
    if (nm !== cn) continue;
    out.push({
      time: row[iTime] == null ? '' : String(row[iTime]),
      customer: cn,
      memo: row[iMemo] == null ? '' : String(row[iMemo]),
      manager: row[iMgr] == null ? '' : String(row[iMgr]),
      nextDate: row[iNext] == null ? '' : String(row[iNext])
    });
  }
  out.sort(function (a, b) {
    return String(b.time).localeCompare(String(a.time));
  });
  return out.slice(0, 20);
}

function doGet(e) {
  e = e || {};
  const p = e.parameter || {};
  const action = String(p.action || '').trim();
  try {
    if (action === 'dashboard') {
      return jsonResponse_(computeDashboardPayload_());
    }
    if (action === 'customers') {
      var customerList = apiGetCustomersList_();
      return jsonResponse_(customerList);
    }
    if (action === 'debugSheet') {
      return jsonResponse_(apiDebugSheetInfo_());
    }
    if (action === 'detail') {
      var rawName = p.name;
      var name = '';
      if (rawName) {
        try {
          name = decodeURIComponent(String(rawName).replace(/\+/g, ' ')).trim();
        } catch (decErr) {
          name = String(rawName).trim();
        }
      }
      if (!name) {
        return jsonResponse_({ error: 'name parameter required' });
      }
      return jsonResponse_(apiGetCustomerDetail_(name));
    }
    if (action === 'getActions') {
      var rawC = p.customer;
      var cname = '';
      if (rawC) {
        try {
          cname = decodeURIComponent(String(rawC).replace(/\+/g, ' ')).trim();
        } catch (decErr2) {
          cname = String(rawC).trim();
        }
      }
      if (!cname) {
        return jsonResponse_({ error: 'customer parameter required' });
      }
      return jsonResponse_(apiGetActions_(cname));
    }
    return jsonResponse_({ error: 'unknown action', action: action || null });
  } catch (err) {
    return jsonResponse_({ error: err && err.message ? err.message : String(err) });
  }
}

/** POST: action=saveAction (application/x-www-form-urlencoded) */
function doPost(e) {
  e = e || {};
  try {
    var p = e.parameter || {};
    var action = String(p.action || '').trim();
    if (action === 'saveAction') {
      var customer = String(p.customer || '').trim();
      var memo = String(p.memo || '').trim();
      var manager = String(p.manager || '').trim();
      var nextDate = String(p.nextDate || '').trim();
      return jsonResponse_(apiSaveAction_(customer, memo, manager, nextDate));
    }
    return jsonResponse_({ error: 'unknown action', success: false });
  } catch (err) {
    return jsonResponse_({
      success: false,
      error: err && err.message ? err.message : String(err)
    });
  }
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 거래처요약 기반 — 대시보드 KPI (프론트 computeDashboardFromCustomers 와 동일 필드). 잔액>0 미수 거래처만 집계. */
function computeDashboardPayload_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = ss.getSheetByName('거래처요약');
  if (!summary) throw new Error('거래처요약 시트가 없습니다.');
  const values = summary.getDataRange().getValues();

  const tz = Session.getScriptTimeZone() || 'Asia/Seoul';
  const asOf = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss');

  if (values.length < 2) {
    return {
      totalBalance: 0,
      totalSales: 0,
      totalCollect: 0,
      collectionRate: 0,
      critical: 0,
      high: 0,
      mid: 0,
      riskCnt: 0,
      manage: 0,
      over30: 0,
      over60: 0,
      asOf: asOf
    };
  }

  const header = values[0].map(function (h) { return String(h).trim(); });
  const rows = values.slice(1);

  const idxName = findColumnIndex_(header, ['거래처명']);
  const idxSales = findColumnIndex_(header, ['총판매']);
  const idxCollect = findColumnIndex_(header, ['총수금']);
  const idxBalance = findColumnIndex_(header, ['현재잔액']);
  const idxLastPay = findColumnIndex_(header, ['최근수금일']);
  const idxRisk = findColumnIndex_(header, ['위험등급']);

  if (idxName === -1 || idxBalance === -1 || idxLastPay === -1 || idxRisk === -1) {
    throw new Error('거래처요약 헤더가 맞지 않습니다.');
  }

  const rec = receivableSummaryRows_(rows, idxBalance);

  const totalBalance = rec.reduce(function (s, r) { return s + toNumber_(r[idxBalance]); }, 0);
  const totalSales =
    idxSales === -1 ? 0 : rec.reduce(function (s, r) { return s + toNumber_(r[idxSales]); }, 0);
  const totalCollect =
    idxCollect === -1 ? 0 : rec.reduce(function (s, r) { return s + toNumber_(r[idxCollect]); }, 0);
  const collectionRate = totalSales > 0 ? totalCollect / totalSales : 0;

  var critical = 0;
  var high = 0;
  var mid = 0;
  rec.forEach(function (row) {
    var rk = String(row[idxRisk]).trim();
    if (rk === '매우위험') critical++;
    else if (rk === '위험') high++;
    else if (rk === '주의') mid++;
  });

  const riskCnt = critical + high;
  const manage = critical + high + mid;

  const today = new Date();
  var over30 = 0;
  var over60 = 0;
  var noLastPayDate = 0;
  /* 핵심: row[idxLastPay]를 String()으로 변환하지 말고 원본값(Date 객체 포함)을 직접 전달.
     String(dateObj) → "Mon Mar 01 2026..." 형태가 되어 parseLedgerDateNo_ 정규식 파싱 실패함. */
  rec.forEach(function (row) {
    var lpRaw = row[idxLastPay];
    if (lpRaw == null || lpRaw === '' || String(lpRaw).trim() === '-') {
      noLastPayDate++;
      return;
    }
    // Date 객체면 그대로, 문자열이면 파싱 시도
    if (!hasUsableLastPayForAging_(lpRaw)) {
      noLastPayDate++;
      return;
    }
    var diffDays = getDaysFromLastPay_(lpRaw, today);
    if (diffDays >= 30) over30++;
    if (diffDays >= 60) over60++;
  });

  Logger.log('[computeDashboard] rec.length=' + rec.length + ' over30=' + over30 + ' over60=' + over60 + ' noLastPayDate=' + noLastPayDate);
  try { console.log('[computeDashboard] rec.length=' + rec.length + ' over30=' + over30 + ' noLastPayDate=' + noLastPayDate); } catch(e) {}

  return {
    totalBalance: totalBalance,
    totalSales: totalSales,
    totalCollect: totalCollect,
    collectionRate: collectionRate,
    critical: critical,
    high: high,
    mid: mid,
    riskCnt: riskCnt,
    manage: manage,
    over30: over30,
    over60: over60,
    asOf: asOf
  };
}

/** 거래처요약 시트에서 단일 고객 객체 (customers API와 동일 형태) */
function apiBuildCustomerObjFromSummaryRow_(row, header) {
  const iErp = findColumnIndex_(header, ['ERP출처']);
  const iName = findColumnIndex_(header, ['거래처명']);
  const iSales = findColumnIndex_(header, ['총판매']);
  const iCollect = findColumnIndex_(header, ['총수금']);
  const iBal = findColumnIndex_(header, ['현재잔액']);
  const iLast = findColumnIndex_(header, ['최근수금일']);
  const iCnt = findColumnIndex_(header, ['거래건수']);
  const iRisk = findColumnIndex_(header, ['위험등급']);
  const iMemo = findColumnIndex_(header, ['관리메모']);
  var nm = String(row[iName] == null ? '' : row[iName]).trim();
  var memoStr = iMemo === -1 ? '' : String(row[iMemo] == null ? '' : row[iMemo]);
  var erpDirect = '';
  if (iErp !== -1 && row[iErp] != null && String(row[iErp]).trim()) {
    erpDirect = String(row[iErp]).trim();
    if (memoStr.indexOf('ERP:') === -1) {
      memoStr = combineSummaryMemoErp_(erpDirect, memoStr);
    }
  }
  return {
    name: nm,
    erp: erpDirect,
    totalSales: iSales === -1 ? 0 : toNumber_(row[iSales]),
    totalCollection: iCollect === -1 ? 0 : toNumber_(row[iCollect]),
    balance: iBal === -1 ? 0 : toNumber_(row[iBal]),
    lastPayDate: (function() {
      var lpRaw = row[iLast];
      if (lpRaw == null) return '';
      // Date 객체면 YYYY-MM-DD 형식으로 변환 (String()으로 변환하면 파싱 불가능한 긴 문자열이 됨)
      if (Object.prototype.toString.call(lpRaw) === '[object Date]' && !isNaN(lpRaw.getTime())) {
        var mo = String(lpRaw.getMonth() + 1).padStart(2, '0');
        var dy = String(lpRaw.getDate()).padStart(2, '0');
        return lpRaw.getFullYear() + '-' + mo + '-' + dy;
      }
      return String(lpRaw).trim();
    })(),
    tradeCount: iCnt === -1 ? 0 : toNumber_(row[iCnt]),
    risk: String(row[iRisk] == null ? '' : row[iRisk]).trim(),
    memo: memoStr
  };
}

/**
 * 거래처요약 → 고객 배열 (프론트는 JSON 배열을 기대). 원장 없이도 잔액명세만으로 동작.
 *
 * 미수 대시보드 전용: 현재잔액 > 0 인 거래처만 반환(방법 B).
 * 방법 A(참고): 전체 거래처를 내려주고 risk === '미수외(음수)' 등으로 프론트에서 위험목록 제외.
 * 향후 미지급·매입처 관리 탭을 넣을 때는 balance < 0 전용 API/쿼리 파라미터로 분리하는 편이 안전함.
 */
function apiGetCustomersList_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = ss.getSheetByName('거래처요약');
  if (!summary) throw new Error('거래처요약 시트가 없습니다.');
  const values = summary.getDataRange().getValues();
  if (values.length < 2) return [];

  const header = values[0].map(function (h) { return String(h).trim(); });
  const iName = findColumnIndex_(header, ['거래처명']);
  const iBal = findColumnIndex_(header, ['현재잔액']);

  if (iName === -1 || iBal === -1) {
    throw new Error('거래처요약 헤더가 맞지 않습니다.');
  }

  const out = [];
  var totalRows = 0;
  var skippedNoName = 0;
  var skippedNoBalance = 0;
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var nm = String(row[iName] == null ? '' : row[iName]).trim();
    if (!nm) { skippedNoName++; continue; }
    totalRows++;
    if (toNumber_(row[iBal]) <= 0) { skippedNoBalance++; continue; }
    out.push(apiBuildCustomerObjFromSummaryRow_(row, header));
  }
  Logger.log('[apiGetCustomersList] totalRows=' + totalRows + ' skippedNoBalance=' + skippedNoBalance + ' returned=' + out.length);
  try { console.log('[apiGetCustomersList] totalRows=' + totalRows + ' skippedNoBalance=' + skippedNoBalance + ' returned=' + out.length); } catch(e) {}
  return out;
}

/**
 * 진단용: 거래처요약 시트의 원시 통계 반환.
 * ?action=debugSheet 로 호출. 브라우저에서 직접 API URL에 붙여 확인 가능.
 */
function apiDebugSheetInfo_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = ss.getSheetByName('거래처요약');
  if (!summary) return { error: '거래처요약 시트 없음' };

  const values = summary.getDataRange().getValues();
  if (values.length < 2) return { error: '데이터 없음', sheetRows: values.length };

  const header = values[0].map(function(h) { return String(h).trim(); });
  const iName = findColumnIndex_(header, ['거래처명']);
  const iBal  = findColumnIndex_(header, ['현재잔액']);
  const iLast = findColumnIndex_(header, ['최근수금일']);
  const iRisk = findColumnIndex_(header, ['위험등급']);

  var totalDataRows = 0;
  var balancePositive = 0;
  var balanceZero = 0;
  var balanceNegative = 0;
  var noName = 0;
  var sample = [];

  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var nm = iName === -1 ? '' : String(row[iName] == null ? '' : row[iName]).trim();
    if (!nm) { noName++; continue; }
    totalDataRows++;
    var bal = iBal === -1 ? 0 : toNumber_(row[iBal]);
    if (bal > 0) {
      balancePositive++;
      if (sample.length < 20) {
        var lpRaw = iLast === -1 ? '' : row[iLast];
        var lpStr = '';
        if (lpRaw != null) {
          if (Object.prototype.toString.call(lpRaw) === '[object Date]' && !isNaN(lpRaw.getTime())) {
            lpStr = lpRaw.getFullYear() + '-' + String(lpRaw.getMonth()+1).padStart(2,'0') + '-' + String(lpRaw.getDate()).padStart(2,'0');
          } else {
            lpStr = String(lpRaw).trim();
          }
        }
        sample.push({
          name: nm,
          balance: bal,
          lastPayDate: lpStr,
          risk: iRisk === -1 ? '' : String(row[iRisk]).trim()
        });
      }
    } else if (bal === 0) {
      balanceZero++;
    } else {
      balanceNegative++;
    }
  }

  Logger.log('[debugSheet] totalDataRows=' + totalDataRows + ' balancePositive=' + balancePositive + ' balanceZero=' + balanceZero + ' balanceNegative=' + balanceNegative);

  return {
    header: header,
    sheetTotalRows: values.length - 1,
    noName: noName,
    totalDataRows: totalDataRows,
    balancePositive: balancePositive,
    balanceZero: balanceZero,
    balanceNegative: balanceNegative,
    iNameCol: iName,
    iBalCol: iBal,
    iLastPayCol: iLast,
    samplePositive: sample
  };
}

function apiGetCustomerDetail_(customerName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = ss.getSheetByName('거래처요약');
  if (!summary) {
    return { customer: null, transactions: [] };
  }
  const values = summary.getDataRange().getValues();
  if (values.length < 2) {
    return { customer: null, transactions: [] };
  }
  const header = values[0].map(function (h) { return String(h).trim(); });
  const iName = findColumnIndex_(header, ['거래처명']);
  if (iName === -1) {
    return { customer: null, transactions: [] };
  }

  var customer = null;
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (String(row[iName]).trim() === customerName) {
      customer = apiBuildCustomerObjFromSummaryRow_(row, header);
      break;
    }
  }
  if (!customer) {
    return { customer: null, transactions: [] };
  }

  var tx = rawTransactionsForCustomer_(customerName);
  return {
    customer: customer,
    transactions: tx
  };
}

/** RAW_거래처원장 — 상세조회 보조. 거래처원장이 없으면 빈 배열. */
function rawTransactionsForCustomer_(customerName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName('RAW_거래처원장');
  if (!raw) return [];
  const data = raw.getDataRange().getValues();
  if (data.length < 2) return [];

  const rh = data[0].map(function (h) { return String(h).trim(); });
  var iName = findColumnIndex_(rh, ['거래처명']);
  var iType = findColumnIndex_(rh, ['행구분']);
  var iDate = findColumnIndex_(rh, ['일자-No.']);
  var iItem = findColumnIndex_(rh, ['품목명[규격]']);
  var iAmt = findColumnIndex_(rh, ['금액']);
  var iBal = findColumnIndex_(rh, ['잔액']);

  if (iName === -1 || iType === -1 || iDate === -1 || iItem === -1 || iAmt === -1 || iBal === -1) {
    iName = 0;
    iType = 2;
    iDate = 3;
    iItem = 4;
    iAmt = 7;
    iBal = 13;
  }

  var acc = [];
  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    if (String(row[iName]).trim() !== customerName) continue;
    var type = String(row[iType]).trim();
    if (type === '월계' || type === '합계') continue;
    var content = String(row[iItem] == null ? '' : row[iItem]).trim();
    if (!content || content === '합계') continue;

    acc.push({
      date: row[iDate],
      type: type,
      content: content,
      amount: toNumber_(row[iAmt]),
      balance: toNumber_(row[iBal]),
      _sortKey: ledgerSortKey_(row[iDate]),
      _row: r
    });
  }

  acc.sort(function (a, b) {
    if (b._sortKey !== a._sortKey) return b._sortKey - a._sortKey;
    return b._row - a._row;
  });

  return acc.map(function (x) {
    return {
      date: x.date,
      type: x.type,
      content: x.content,
      amount: x.amount,
      balance: x.balance
    };
  });
}

function ledgerSortKey_(dateNo) {
  if (Object.prototype.toString.call(dateNo) === '[object Date]' && !isNaN(dateNo.getTime())) {
    return dateNo.getTime();
  }
  const now = new Date();
  const parsed = parseLedgerDateNo_(dateNo, now);
  if (parsed && !isNaN(parsed.getTime())) return parsed.getTime();
  return 0;
}

function getOrCreateErrorLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ERROR_LOG_SHEET_NAME);

  if (!sh) {
    sh = ss.insertSheet(ERROR_LOG_SHEET_NAME);
  }

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 4).setValues([[
      '발생시간',
      '발생위치',
      '파일명',
      '오류내용'
    ]]);

    sh.getRange(1, 1, 1, 4)
      .setBackground('#990000')
      .setFontColor('white')
      .setFontWeight('bold');

    sh.autoResizeColumns(1, 4);
  }

  return sh;
}

function appendErrorLog_(where, fileName, err) {
  try {
    const sh = getOrCreateErrorLogSheet_();
    const message = err && err.message ? err.message : String(err);

    sh.appendRow([
      new Date(),
      where || '',
      fileName || '',
      message
    ]);
  } catch (logErr) {
    Logger.log('오류로그 기록 실패: ' + logErr);
  }
}
/**
 * 스프레드시트 열릴 때 메뉴 생성
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📊 현금흐름관리')
    .addItem('전체 새로고침', 'runFullRefresh')
    .addItem('파일 가져오기 (요약만)', 'importLedgerFiles')
    .addItem('대시보드+관리대상 생성', 'runReportsFromSummary')
    .addItem('파일 이동 실행', 'moveProcessedFilesOnly')
    .addSeparator()
    .addItem('거래처요약 재생성', 'buildLedgerSummary')
    .addItem('대표대시보드 재생성', 'buildCEOdashboard')
    .addItem('관리대상 재생성', 'generateManagementList')
    .addSeparator()
    .addItem('위험등급 색상 적용', 'applyRiskColorFormatting')
    .addToUi();
}

/**
 * 가져오기 후 대시보드·관리대상만 생성 (요약 시트는 이미 있어야 함)
 */
function runReportsFromSummary() {
  buildCEOdashboard();
  generateManagementList();
}

/**
 * 전체 새로고침
 * ETL → 요약 → 대시보드 → 관리대상 → 색상 (단계별 실행 가능)
 */
function runFullRefresh() {
  const ui = SpreadsheetApp.getUi();

  try {
    ui.alert('전체 새로고침 시작');

    importLedgerFiles();
    runReportsFromSummary();

    applyRiskColorFormatting();

    ui.alert(
      '전체 새로고침 완료\n\n' +
      '✔ 거래처요약 갱신\n' +
      '✔ 대표대시보드 갱신\n' +
      '✔ 관리대상 갱신\n' +
      '✔ 위험등급 색상 적용 완료'
    );

  } catch (err) {

    Logger.log(err);
    appendErrorLog_('전체 새로고침', '', err);

    ui.alert(
      '오류 발생\n\n' +
      (err && err.message ? err.message : String(err))
    );
  }
}
function getOrCreateProcessedFolder_() {
  const parent = DriveApp.getFolderById(FOLDER_ID);
  const folders = parent.getFoldersByName(PROCESSED_FOLDER_NAME);

  if (folders.hasNext()) {
    return folders.next();
  }

  return parent.createFolder(PROCESSED_FOLDER_NAME);
}

/**
 * 처리 완료 파일을 처리완료 폴더로 이동 (importLedgerFiles 와 분리 실행).
 * 메뉴 "파일 이동 실행" 또는 독립 트리거로 호출.
 */
function moveProcessedFilesOnly() {
  const folder = DriveApp.getFolderById(FOLDER_ID);
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    const lower = f.getName().toLowerCase();
    if (lower.endsWith('.csv') || lower.endsWith('.xlsx')) {
      files.push(f);
    }
  }

  if (files.length === 0) {
    SpreadsheetApp.getUi().alert('이동할 파일이 없습니다.');
    return;
  }

  const movedCount = moveProcessedFiles_(files);
  Logger.log('[moveOnly] 이동 완료 ' + movedCount + '/' + files.length);
  SpreadsheetApp.getUi().alert('파일 이동 완료: ' + movedCount + '개');
}

function moveProcessedFiles_(files) {
  if (!files || files.length === 0) return 0;

  let processedFolder;
  try {
    processedFolder = getOrCreateProcessedFolder_();
  } catch (err) {
    Logger.log('처리완료 폴더 접근 실패: ' + (err && err.message ? err.message : String(err)));
    appendErrorLog_('처리완료 폴더 이동', '', err);
    return 0;
  }

  let moved = 0;

  files.forEach(function(file) {
    try {
      file.moveTo(processedFolder);
      moved++;
    } catch (err) {
      Logger.log('처리완료 이동 실패: ' + file.getName() + ' / ' + (err && err.message ? err.message : String(err)));
      appendErrorLog_('처리완료 폴더 이동', file.getName(), err);
    }
  });

  return moved;
}