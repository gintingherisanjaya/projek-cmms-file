/**
 * ===============================
 * DEPENDENCIES
 * ===============================
 */
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const https = require('https');
const axios = require('axios');


/**
 * ===============================
 * UTIL
 * ===============================
 */
const normalizeKey = v =>
  String(v ?? '').replace(/\s+/g, ' ').trim().toLowerCase();

/**
 * ===============================
 * FETCH CSV FROM GOOGLE SHEET
 * ===============================
 */
const MAX_RETRY = 3;
const TIMEOUT = 15000;
const CACHE_DIR = path.join(__dirname, 'csv_cache');

// pastikan folder cache ada
if (!fs.existsSync(CACHE_DIR)) {
  fs.mkdirSync(CACHE_DIR);
}

async function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

async function fetchCSVFromURL(url, retry = 0) {
  const cacheFile = path.join(
    CACHE_DIR,
    Buffer.from(url).toString('base64') + '.csv'
  );

  try {
    const response = await axios.get(url, {
      responseType: 'text',
      timeout: TIMEOUT,
      maxRedirects: 5,
      validateStatus: status => status >= 200 && status < 400
    });

    const contentType = response.headers['content-type'] || '';

    // ❌ BUKAN CSV
    if (!contentType.includes('text/csv')) {
      throw new Error(`Invalid content-type: ${contentType}`);
    }

    const data = response.data.trim();

    // safety check tambahan
    if (data.startsWith('<') || data.toLowerCase().includes('<html')) {
      throw new Error('CSV response is HTML');
    }

    // simpan cache
    fs.writeFileSync(cacheFile, data, 'utf8');

    return data;
  } catch (err) {
    // hapus cache jika ada
    if (fs.existsSync(cacheFile)) {
      fs.unlinkSync(cacheFile);
    }

    if (retry < MAX_RETRY) {
      const wait = 1000 * Math.pow(2, retry); // exponential backoff
      console.warn(
        `[RETRY ${retry + 1}/${MAX_RETRY}] ${url} (${err.message})`
      );
      await delay(wait);
      return fetchCSVFromURL(url, retry + 1);
    }

    // Buat log gagal akses
    const failedAccess = {
      url: url,
      timestamp: new Date().toISOString(),
      attemptNumber: retry + 1,
      status: 'FAILED',
      errorMessage: err.message,
      statusCode: err.response?.status || null
    };

    // Tambahkan ke file log gagal akses
    const failedAccessFile = path.join(__dirname, 'failed_access.json');
    let failedAccessLog = [];
    if (fs.existsSync(failedAccessFile)) {
      const existingLog = fs.readFileSync(failedAccessFile, 'utf8');
      failedAccessLog = JSON.parse(existingLog);
    }
    failedAccessLog.push(failedAccess);
    fs.writeFileSync(failedAccessFile, JSON.stringify(failedAccessLog, null, 2));

    console.error(`FAILED after ${MAX_RETRY} retries: ${url} | ${err.message}`);
    return null; // Return null instead of throwing error to continue processing other files
  }
}

/**
 * ===============================
 * CSV UTILITIES
 * ===============================
 */
function detectDelimiter(lines) {
  let comma = 0,
    semi = 0;
  lines.slice(0, 20).forEach(l => {
    comma += (l.match(/,/g) || []).length;
    semi += (l.match(/;/g) || []).length;
  });
  return semi > comma ? ';' : ',';
}

function parseCsvLine(line, delimiter) {
  const out = [];
  let cur = '';
  let quoted = false;

  for (const c of line) {
    if (c === '"') quoted = !quoted;
    else if (c === delimiter && !quoted) {
      out.push(cur.trim().replace(/^"|"$/g, ''));
      cur = '';
    } else cur += c;
  }
  out.push(cur.trim().replace(/^"|"$/g, ''));
  return out;
}

function csvToJson(csv, source) {
  csv = csv.replace(/\uFEFF/g, '').replace(/\r/g, '');
  const lines = csv.split('\n');

  const HEADER_ROW_INDEX = 2;
  const delimiter = detectDelimiter(lines);

  const header = parseCsvLine(lines[HEADER_ROW_INDEX], delimiter);
  const headerMap = {};
  header.forEach((h, i) => {
    const k = h.toLowerCase().replace(/\./g, '').trim();
    if (k) headerMap[k] = i;
  });

  const pick = (row, name) =>
    headerMap[name] !== undefined ? row[headerMap[name]] ?? '' : '';

  const data = [];

  for (let i = HEADER_ROW_INDEX + 1; i < lines.length; i++) {
    if (!lines[i].trim()) continue;
    const row = parseCsvLine(lines[i], delimiter);

    const obj = {
      excelRow: i + 1,
      filePath: source,
      equipment: pick(row, 'equipment number'),
      costCenter: pick(row, 'cost center before'),
      eqktu: pick(row, 'eqktu before'),
      objectType: pick(row, 'object type'),
      funcLoc: pick(row, 'functional location before'),
      funcDesc: pick(row, 'functloc desc before')
    };
    
    // If objectType is empty, try to get it from the fixed column index
    if (!obj.objectType && CHECK_MAPPING.objectType !== undefined) {
      obj.objectType = row[CHECK_MAPPING.objectType] ?? '';
    }

    if (obj.equipment?.trim()) data.push(obj);
  }

  return data;
}

/**
 * ===============================
 * EXCEL → JSON (XLSX)
 * ===============================
 */
function excelToJson(filePath, mapping) {
  if (!fs.existsSync(filePath)) return [];

  const wb = XLSX.readFile(path.resolve(filePath));
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

  const data = [];

  rows.forEach((row, i) => {
    if (i === 0) return;

    const obj = { excelRow: i + 1, filePath };
    for (const [k, idx] of Object.entries(mapping)) {
      obj[k] = row[idx] ?? '';
    }

    if (!obj.equipment?.trim()) return;
    data.push(obj);
  });

  return data;
}

const TRUTH_FILE = 'TRUTH_OF_DATA_30_JANUARI.xlsx';

const CHECK_FILE = [
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vQ1Ohel3lg3mT_99_d5vFGVNOuQquXzXkGm2MopT7CM-CJ3uX-k8Y0GnqTTh-uulA/pub?gid=1112545246&single=true&output=csv' , // sei daun
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vR0CLq-Ws2Qx5KC6VxaRRoxA8mlAbfY8GAsXSAUDOc4FRh5N5mo4RHg1vYFfERl9w/pub?gid=1508731010&single=true&output=csv' , // torgamba
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vSD1bSL9engMSgrbfX4jnsrLp0lS0hLEFN7LaFIYclKyZ8xneLqT2WfYOyviZJkPw/pub?gid=2026123966&single=true&output=csv' ,// baruhur
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vS6GJOiIhMA4JySopybdpd59HLiGlfddkBqSZ9W7dnogVKtloKVRfilZZm-n6E0FQ/pub?gid=242435122&single=true&output=csv', // 4.PKS AEK TOROP
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vQGi4WiW2WhDurivJhe2N3PZrpJxP1KoArDL4Zk_-j52S8Xrq1n5MoZmMunjAp_4Q/pub?gid=1338975567&single=true&output=csv', // 5.PKS AEK RASO
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vQ-NmEStyUV-olabZtZFA_DburI0yUP-QebjLRpdRUKUHtF8RER0qDWWAABRYgSUQ/pub?gid=1763009314&single=true&output=csv', // 6.PKS SISUMUT.xlsx
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vQjkJmSplmWAwuT5-vxJmnUlHZqdnkcMyvCmQ5XOyOObLDdPPgMCKA325GpTphWwA/pub?gid=1968431822&single=true&output=csv', // 7.PKS AEK NABARA
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vTIacdD_VAjOqcZXk5Gm0jt9_AS9A7gnhjiGDimfr39L0fmY6UDkrbQjApeioBZkg/pub?gid=1554806748&single=true&output=csv', // 8.PKS SEI SILAU
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vTehqYizF4rrf_jNamPd_ZjbhrKSCZGo4q3vo6bUe6jDfEy4lmOvb-PTpXvl7tS3A/pub?gid=419492571&single=true&output=csv', // 9.PKS SEI MANGKEI
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vRMsnTKrwuCAR3VOZ1gBlizhe-xuGukyDBVGUVWzts03uEY3PO0IH4Ql5TVFXanFw/pub?gid=566858846&single=true&output=csv', // 10.PKS RAMBUTAN
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vQXJ0rxH8tCW7MZSMzavrVg1oDaGrtcal8Q19_ulBOS_15apTlAXkFHP1JduaPobw/pub?gid=869966379&single=true&output=csv', // 11.PKS HAPESONG
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vREeKUr7xR7Kx_scr0ovElJyPqsbC_PnU8N1s6h5GsVWRMK8wae50H9j5gTEPXtAg/pub?gid=1707525293&single=true&output=csv', // 12.PKS CIKASUNGKA
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vS9MfapwKUAJRkupFOPGB9uJ3Nces8c62J_6j6JGDK9lLBmSBw3zdz5gyR0cC86Og/pub?gid=2032689621&single=true&output=csv', // 13.PKS KARTAJAYA
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vRxpcCCYFwhjxk8SWCyyBQr-mSNnxQqtQBOGXRAQKvlmXvsasBQGZfo2bRH7kTFPg/pub?gid=1775973386&single=true&output=csv', // 14.PKS SEI MERANTI

  'https://docs.google.com/spreadsheets/d/e/2PACX-1vQdyquDuwCYpQ73ocZyG8TKDLPGkOAQfO04KVyhYmsBhM_e_yYKYb5Qb33Nka3TNw/pub?gid=1385837078&single=true&output=csv', // 1.PKS BAH JAMBI => R2
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vR3P6NAT_8Rfdd98ue-eYB4UgYy1esV_yhcJBx6u6SnbRWbb3J5uGJqioRrZ6L1Dg/pub?gid=390681956&single=true&output=csv', // 2.PKS PASIR MANDOGE
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vTP7iehodVdr3YLcPuKsQhDZj7OSbvzT6zus75ar-glHe0gSF9XM9_5TuBRup013w/pub?gid=1608092781&single=true&output=csv', // 3.PKS DOLOK SINUMBA
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vQUX1hEj-SUpY-kKzr5qI7d4MsQxPdjEf9XzplA8TET8ZqDwImsgb32wuaFym_gIg/pub?gid=956359612&single=true&output=csv', // 4.PKS DOLOK ILIR
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vSs6g08ziMB64OFgzI8E7Ub-gGG66nOqV7cOfChEC0H7EgXxjhQu_IFjyVpMdWZZA/pub?gid=1347795556&single=true&output=csv', // 5.PKS GUNUNG BAYU
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vQNG1EImF3Uh__byYG7AZ1uxqe0xJnXwEq1P1JC9txCQ1CDYUqhoDPcfViKGP92nQ/pub?gid=785590884&single=true&output=csv', // 6.PKS MAYANG
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vQDFCDPsaVOVFa-cSowNUf8iTtBbOj139XGThto4yhBKEz2FBBv0y8ZJJDegcV_2g/pub?gid=1770948914&single=true&output=csv', // 7.PKS ADOLINA
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vSsI-D6hT8oyFkS6G0Cjod-dp55_1qYtuB03n9xZzjEZ33pItX4GUOl7BGvgskhkQ/pub?gid=511072893&single=true&output=csv', // 8.PKS PABATU
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vQbsepJWHzcXZVmOhMqImLQbErOdeoYWWZpKNCTqDFpePJ42xL30942a7gQLvrABw/pub?gid=522745381&single=true&output=csv', // 9.PKS SAWIT LANGKAT
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vRs_vhjq4OAdz0TwOhC2PRg8aDPg4sJm4Oe2wK51YdFq34T-5dp-rhxP3qAjpghTg/pub?gid=1735414182&single=true&output=csv', // 10.PKS TIMUR
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vTHlCVT0i5walptEHrL1cLnAtIO7u3q4Jz6k7uwphgs19PG5GJRU2KNHoXpoZkCRQ/pub?gid=1038713030&single=true&output=csv', // 11.PKS TINJOAN
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vRelWxwsquzeEPFbPJYUh_WQwF1cNs_g3m-HX1wodIlYEfTB6G8rDNLZFnG8MsdAw/pub?gid=967328015&single=true&output=csv', // 12.PKS AIR BATU
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vSduUVaNnkc4pjPws1hiPtXLq802WWANlpzGmsUgcuhb_oq-cgY_Oy23Tk2a1FmEg/pub?gid=566101832&single=true&output=csv', // 13.PKS PULU RAJA
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vTkD6ByKggL1tjwvabwhFDvTe4SQUrpQJOBF5b5yqc1MmGFVBJvfBkfGKYiCTE2lw/pub?gid=732669087&single=true&output=csv', // 14.PKS BERANGIR
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vQ8ZgXRuUpi544Qcg37zWJDXgu8VRXEJeFxP_A1So3odktjib09yjjNBl1LL0wK7A/pub?gid=730493562&single=true&output=csv', // 15.PKS AJAMU
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vQLCYEGZeEkq2qxPujfVT047VH8pqLa4fABKabKbNhBNXUI0IYaWgLhBhotvk9Sxg/pub?gid=814829854&single=true&output=csv', // 16.PKS SOSA
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vQbnsdaP4hVZgzo7HTLbAp_5QN5qmnz9mWyKHnhj8h0fgG9UvO7ayK64fv4b8EXWA/pub?gid=1974025468&single=true&output=csv', // 17.PKS KWALA SAWIT
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vT6UKXKNOD6jrjNSDe4cwqsmEUA3Na_SnR27R6pEpDIZyWAeo3SFC1lq2EkbYVdKg/pub?gid=667453919&single=true&output=csv', // 18.PKS SAWIT HULU
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vTzL2bKmtBw-eYX3lNhvlUolPATwMo-wHA8gcWonYcer76kQJMMeCCFjtg_dt5_4w/pub?gid=1146381308&single=true&output=csv', // 19.PKS SAWIT SEBERANG
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vQNR5KWWUWcRUQRkXTi9pWAgPJw0yonBiAof6cQeVa5QRL2Ha3Z0UXR6Z69OugM8w/pub?gid=751382986&single=true&output=csv', // 20.PKS PAGAR MERBAU
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vQUwWg_NudqRt_wMSp03C9tjyaWUBDp2iaoeLB-LbA0Py5L6Yes3n72WyaH3-OAzw/pub?gid=2035288453&single=true&output=csv', // 21.PKS LUWU

  'https://docs.google.com/spreadsheets/d/e/2PACX-1vT08ePA9H0B9z9TIX9dhN6Knsu3RbYvryoBlWQhKiqI_RzEmHssDREq6K5rgPCsRQ/pub?gid=480170633&single=true&output=csv', // 1.PKS TANAH PUTIH  => R3
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vR2jzVI108O9DLOROwff7vHJWzFcywZktLnAWU3wrezmqPVmUezE4oCOZ83ZMqKXA/pub?gid=282841402&single=true&output=csv', // 2.PKS TANJUNG MEDAN
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vRXXPqWEiQ9ZrjiaOt4jv-W2ORSdLPo2DUd3u7fkfRHNTmebZ1JlXiZmRaeaWhnGQ/pub?gid=1205515922&single=true&output=csv', // 3.PKS SEI GALUH
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vStkw3m5Qhh4vofwrlWLquVNWyUvzPXqVz-PRU5XVD6dhF6Poii_s9b4QddEdv_uA/pub?gid=2024406804&single=true&output=csv', // 4.PKS SEI PAGAR
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vSnFq1FlLt2fnUFiP1tPOe78c5K83xia5cYZWMq_BNDvuSKgwsRLqXner0eH1BL4g/pub?gid=1064920505&single=true&output=csv', // 5.PKS SEI GARO
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vQXGRmZswgOE1lYortkZsEK-k9TAxi2x4OdLtk0VqXU8-3zCNhLYp7NtdFHqF1qJw/pub?gid=2132027831&single=true&output=csv', // 6.PKS SEI BUATAN
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vR8-X9FfZQWrU_WDCxRSrHq2Em7radl2O_GyOYeB5l_ie8yRypUBhvHvFxygWN_iw/pub?gid=1792934369&single=true&output=csv', // 7.PKS LUBUK DALAM
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vS65NB5Alf6chr3zNryplxfvLDoW-6HfoHKUzSoCiBRWVDWMXxbQGUSZiaELaB6sA/pub?gid=8490190&single=true&output=csv', // 8.PKS TANDUN
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vQWW3KFBH9tQ4LN_ZUsQBWh7hV4z6FkIlBWKxeJN41Rfb2L1F7ldWipoZxKK0c6XA/pub?gid=26364533&single=true&output=csv', // 9.PKS TERANTAM
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vRqqB0uGKju3_tiCBMJBqczsj7EqVOSCslwEiAgEumaiDlA0pdOxNb0vX1lT6fMqQ/pub?gid=699233290&single=true&output=csv', // 10.PKS SEI TAPUNG
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vSb3uqDp64dmTvNhPG0Y0ExfdLuKObzhQfUyIKctV8Df70dq7PKepC6GkXw0oGHqg/pub?gid=43710913&single=true&output=csv', // 11.PKS SEI ROKAN
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vRscr_f7F8FDghwghpe01Rv4h-3OnxMvHFnr3QcC8JidCf9-wUVypDayraJUzDp3Q/pub?gid=2054649331&single=true&output=csv', // 12.PKS SEI INTAN

  'https://docs.google.com/spreadsheets/d/e/2PACX-1vRUHAvnW3Gy3_g1ary78z3_D_Iit-BylCvGdWwuPP0QTyrJ7lGlD4Eh1IDLoekjfg/pub?gid=388213494&single=true&output=csv', // 1.PKS AUR GADING => R4
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vSM6gCfMUAh2-6xlGPjE7IsoL08eAK4IY2KpHQMVEfresnga8FgqnDA5-beCB3Agg/pub?gid=1676843916&single=true&output=csv', // 2.PKS PINANG TINGGI
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vTcrEgoU-gLvgJE5sW6U4zvEXk_YNK2_2u4eh7gRIfcHMN1qE7DfRPDlRsoUH2MzA/pub?gid=868367737&single=true&output=csv', // 3.PKS BUNUT
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vQ-W1Sj1fEUWA1p9ZomfNZTI8eYSOCKPp1uGv8SE1fBp7RgTWvWS75ROkEiCLiDbA/pub?gid=1468518558&single=true&output=csv', // 4.PKS TANJUNG LEBAR
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vSBD3pdVmzZQLx2i8OVi-RtSM-8Ub_Wf8QDwup4NWv1tRlC45yldFPBPiQTTxBn1Q/pub?gid=1821767842&single=true&output=csv', // 5.PKS RIMBO DUO
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vSEnYSd4mv_e52J5XdPMGpmK8Q2kJZX9E5W6pHlLIBiEzd6GjaxLKwAEBRwN1viWg/pub?gid=1420018792&single=true&output=csv', // 6.PKS SOLOK SELATAN
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vT0R4YVgV5If4AHw2CiDgsQCOnvEVw-udD30IpDtVZmPTm-rUYZgGTD1f-uMkcZWA/pub?gid=1520177365&single=true&output=csv', // 7.PKS OPHIR
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vTbmAfU4oWsOL02PCqTfDmK1OAQejJhK12havpZaMORy1dON4aXzXrUYZUFDpQEfQ/pub?gid=255196905&single=true&output=csv', // 8.PKS PENGABUAN

  'https://docs.google.com/spreadsheets/d/e/2PACX-1vSEq0K6clhIGiOgVtcR6J-lXoAR30askU-ROh189nb1qE7r8zYuqqJnrvlwNP0Wbw/pub?gid=904994934&single=true&output=csv', // 1.PKS GUNUNG MELIAU => R5
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vRmOPUNqr0gzWGcGuDE5itYqCcgDfDHoGVPgbdDXFIIke-gJfkHmUc4st-4Ta9QDg/pub?gid=1861905872&single=true&output=csv', // 2.PKS RIMBA BELIAN
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vSWUxgpqyVNgyLuxYEDeY61iWEAW9pUKM5BQG3bbbWKXvCtOf5ZlrgYDFT1cM0HDg/pub?gid=2058943687&single=true&output=csv', // 3.PKS NGABANG
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vS11kO36HLHLHFccIjdAwW7aCDJJPH5ZPgxWnW3ovSHqba61hkr8YZmCx1XkM1agQ/pub?gid=702634424&single=true&output=csv', // 4.PKS PELAIHARI
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vQGJR0gmldRDgSkLTfUVXkw8wmt4m9X2wIy8EaEUTCqxt5Sl4jcrK-oUsXjn-N57A/pub?gid=1398434405&single=true&output=csv', // 5.PKS KEMBAYAN
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vSdS5wDPBV31GMMdYdWswGwzUobMIOMqiDB_8QWMIxTP2R6FrB-v0N7lewmyig2nw/pub?gid=789263867&single=true&output=csv', // 6.PKS SEMUNTAI
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vTupcbWvX1EXJNQ9wJzIAeEBDX7mMoVBYwzT3RocC6wCSFjNPYzc-LRDHFGH7PPuw/pub?gid=1297500265&single=true&output=csv', // 7.PKS PARINDU
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vSQDATAQ0wypzdhBiurY-cmNInCj4cdU-zmZHDLf_15Dl42bkW-RWoi8_ld50zm1g/pub?gid=1674164763&single=true&output=csv', // 8.PKS LONG PINANG
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vRUnw7M_8R17b77xYD8DmS0229ZRC2ROtDyHLBOnS2yvEhEAv91u_dhStC_K7tI1w/pub?gid=1278152290&single=true&output=csv', // 9.PKS PAMUKAN

  'https://docs.google.com/spreadsheets/d/e/2PACX-1vR6b7ZPTlKBS5Iqn1b4PfTvJNcRdhRtpsCvfC_VNBXbo2rUWqesnwgObBBCmDkh-g/pub?gid=709306947&single=true&output=csv', // 1.PKS PULAU TIGA => R6
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vR2t8AUgYIQd9zJ80f-gn7AvhEHESwmxRA3AvXqLA90Ry73mUgIi0qd0sOLjposkQ/pub?gid=1656447158&single=true&output=csv', // 2.PKS TANJUNG SEUMANTOH
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vTzdIiEymwaHwvL8CMbi3AGPVu_JhNWhZxynnOzqAc8oc0UbAXZbhCn2DUnezBseA/pub?gid=1309883327&single=true&output=csv', // 3.PKS COT GIREK

  'https://docs.google.com/spreadsheets/d/e/2PACX-1vRgEIVemxuRJBJvMv-sYBSR7n6ioQQNzW46BOKPOU80BtjITb1U5C53ndiE9RYByw/pub?gid=1113885983&single=true&output=csv', // 1.PKS BEKRI
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vR7SmPNEkRhJIPXAQXgNbTf9PClNmJJOM3-u5GfW8vbsgUukgoa8OptNlzL63Jr0g/pub?gid=1136275499&single=true&output=csv', // 2.PKS BETUNG
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vQkJefTqr1GGlxxGEsAs-9VoIPffbLTmXixrZGY5ui3m1izLGSG1PlsLTA0RKM4cA/pub?gid=1168888237&single=true&output=csv', // 3.PKS TALANG SAWIT
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vRyM9YtUcWfQ93eQoi9Bgwj-iUB5FWv7tM1cbW6bFpUgKB45n6uwb8zdEqHC6LaFQ/pub?gid=2103644044&single=true&output=csv' // 4.PKS SUNGAI LENGI
];


const TRUTH_MAPPING = {
  equipment: 1,
  costCenter: 4,
  eqktu: 5,
  objectType: 8,
  funcLoc: 16,
  funcDesc: 17
};

const CHECK_MAPPING = {
  equipment: 2,     
  costCenter: 5,    
  eqktu: 6,         
  objectType: 10,    
  funcLoc: 17,      
  funcDesc: 18      
};

// Helper function to determine most likely file for TIDAK DITEMUKAN records
function getRecommendedFileForCostCenter(costCenter, costCenterPrefixMap) {
  const costCenterStr = costCenter.toString().trim();
  if (costCenterStr.length >= 4) {
    const prefix = costCenterStr.substring(0, 4);
    if (costCenterPrefixMap.has(prefix)) {
      const files = Array.from(costCenterPrefixMap.get(prefix));
      // Return the first file (or we could implement more sophisticated logic)
      return files[0];
    }
  }
  return null;
}

async function buildEvaluatedExcel() {
  const evaluate = JSON.parse(fs.readFileSync('evaluate.json', 'utf8'));
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(TRUTH_FILE);
  const ws = wb.worksheets[0];
  const COLORS = {
    'SESUAI': 'FF92D050',
    'TIDAK SESUAI': 'FFFFEB9C',
    'TIDAK DITEMUKAN': 'FFFFC7CE',
    'TIDAK ADA': 'FFB3D9FF'
  };

  let jumlahSesuai = 0
  let jumlahTidakSesuai = 0
  let jumlahTidakDitemukan = 0
  let jumlahTidakAda = 0

  ws._conditionalFormattings = [];
  ws.eachRow((row, rowIndex) => {
    if (rowIndex === 1) return;

    const cellEquipment = row.getCell(TRUTH_MAPPING.equipment + 1);
    const cellCostCenter = row.getCell(TRUTH_MAPPING.costCenter + 1);
    const excelRowEquipment = cellEquipment.value;
    const excelRowCostCenter = cellCostCenter.value;
    
    // Cari entri di evaluate.json berdasarkan EQUIPMENT dan COST_CENTER
    const foundEntry = evaluate.find(d => d.EQUIPMENT === excelRowEquipment && d.COST_CENTER === excelRowCostCenter);
    
    // Gunakan status dari evaluate.json jika ditemukan, jika tidak maka TIDAK DITEMUKAN

    let finalStatus = ''
    if (!foundEntry) {
      finalStatus = 'TIDAK DITEMUKAN';
    }
    else if(foundEntry 
      && foundEntry.COST_CENTER 
      && foundEntry.COST_CENTER.toLowerCase().includes('stas') 
      && !foundEntry.COST_CENTER.toLowerCase().includes('e')
      && !foundEntry.COST_CENTER.toLowerCase().includes('2f19')
      && !foundEntry.COST_CENTER.toLowerCase().includes('2f21')
      && !foundEntry.COST_CENTER.toLowerCase().includes('5f20')
      //  extra (sementara)
      && !foundEntry.COST_CENTER.toLowerCase().includes('1f12')
      && !foundEntry.COST_CENTER.toLowerCase().includes('kf01')
      && !foundEntry.COST_CENTER.toLowerCase().includes('kf02')
      && !foundEntry.COST_CENTER.toLowerCase().includes('9f01')

    ) {
       finalStatus = foundEntry.STATUS;
    } else {
      finalStatus = 'TIDAK ADA';
    }
    
    switch(finalStatus) {
        case 'SESUAI':
            jumlahSesuai++;
            break;
        case 'TIDAK SESUAI':
            jumlahTidakSesuai++;
            break;
        case 'TIDAK DITEMUKAN':
            jumlahTidakDitemukan++;
            break;
        case 'TIDAK ADA':
            jumlahTidakAda++;
             break;
        default:
            break;
    }

    const argb = COLORS[`${finalStatus.trim()}`] ?? 'FFFFFFFF';
    row.eachCell(cell => {
        cell.fill = null;
        cell.style = {}; // penting
    });
    row.eachCell(cell => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb }
      };

    });
  });

  console.log(`📊 FINAL COUNTS`);
  console.log(`   ✅ Jumlah Sesuai     : ${jumlahSesuai}`);
  console.log(`   ❌ Jumlah Tidak Sesuai : ${jumlahTidakSesuai}`);
  console.log(`   ❌ Jumlah Tidak Ditemukan : ${jumlahTidakDitemukan}`);
  console.log(`   ❌ Jumlah Tidak Ada (Bukan alat di Stasiun PKS) : ${jumlahTidakAda}`);

  await wb.xlsx.writeFile('TRUTH_OF_DATA_EVALUATED.xlsx');
  console.log('🎨 TRUTH_OF_DATA_EVALUATED.xlsx SELESAI');
}

/**
 * ===============================
 * MAIN
 * ===============================
 */
(async () => {
  const truthData = excelToJson(TRUTH_FILE, TRUTH_MAPPING);

  let checkData = [];
  let processedCount = 0;
  let failedFiles = [];
  
  for (const src of CHECK_FILE) {
    try {
      if (/^https?:\/\//i.test(src)) {
        const csv = await fetchCSVFromURL(src);
        if (csv !== null) {
          const parsedData = csvToJson(csv, src);
          checkData.push(...parsedData);
          processedCount++;
        } else {
          failedFiles.push({
            filepath: src,
            error: 'Failed to fetch CSV'
          });
        }
      } else {
        const parsedData = excelToJson(src, CHECK_MAPPING);
        checkData.push(...parsedData);
        processedCount++;
      }
    } catch (error) {
      failedFiles.push({
        filepath: src,
        error: error.message
      });
    }
  }
  
  // Save failed files to json
  fs.writeFileSync('salah_template.json', JSON.stringify(failedFiles, null, 2));
  
  console.log(`📁 Total check files processed: ${processedCount}`);
  console.log(`❌ Failed files saved to salah_template.json: ${failedFiles.length}`);

  fs.writeFileSync('truth.json', JSON.stringify(truthData, null, 2));
  fs.writeFileSync('check.json', JSON.stringify(checkData, null, 2));

  // Buat dua map: satu untuk pencocokan berdasarkan kombinasi equipment dan cost center, satu lagi hanya berdasarkan equipment
  const primaryCheckMap = new Map(); // kombinasi equipment dan cost center
  const equipmentCheckMap = new Map(); // hanya equipment
  
  // Create cost center prefix mapping to identify which file records should belong to
  const costCenterPrefixMap = new Map(); // 4-digit cost center prefix -> file path
  
  checkData.forEach(c => {
    if (c) { // Pastikan objek tidak null/undefined
      const primaryKey = `${c.equipment.toString().trim().toLowerCase()}_${c.costCenter.toString().trim().toLowerCase()}`;
      const equipmentKey = c.equipment.toString().trim().toLowerCase();
      
      // Extract first 4 digits of cost center for file mapping
      const costCenterStr = c.costCenter.toString().trim();
      if (costCenterStr.length >= 4) {
        const prefix = costCenterStr.substring(0, 4);
        if (!costCenterPrefixMap.has(prefix)) {
          costCenterPrefixMap.set(prefix, new Set());
        }
        costCenterPrefixMap.get(prefix).add(c.filePath);
      }
      
      if (!primaryCheckMap.has(primaryKey)) primaryCheckMap.set(primaryKey, []);
      primaryCheckMap.get(primaryKey).push(c);
      
      if (!equipmentCheckMap.has(equipmentKey)) equipmentCheckMap.set(equipmentKey, []);
      equipmentCheckMap.get(equipmentKey).push(c);
    }
  });

  const results = [];

  truthData.forEach(t => {
    const primaryKey = `${t.equipment.toString().trim().toLowerCase()}_${t.costCenter.toString().trim().toLowerCase()}`;
    const equipmentKey = t.equipment.toString().trim().toLowerCase();

    // Coba temukan kecocokan pertama berdasarkan kombinasi equipment dan cost center
    if (primaryCheckMap.has(primaryKey)) {
      // Kecocokan ditemukan berdasarkan kombinasi equipment dan cost center
      const c = primaryCheckMap.get(primaryKey)[0];
      
      const cmp = (a, b) =>
        String(a || '').trim().toLowerCase() !==
        String(b || '').trim().toLowerCase();

      // Bandingkan semua 6 parameter: equipment, costCenter, eqktu, objectType, funcLoc, funcDesc
      // ObjectType hanya ada di truth data, jadi kita hanya bisa membandingkan jika ada di kedua data
      const equipmentMatch = !cmp(t.equipment, c.equipment);
      const costCenterMatch = !cmp(t.costCenter, c.costCenter);
      const eqktuMatch = !cmp(t.eqktu, c.eqktu);
      const objectTypeMatch = !cmp(t.objectType, c.objectType || '');
      const funcLocMatch = !cmp(t.funcLoc, c.funcLoc);
      const funcDescMatch = !cmp(t.funcDesc, c.funcDesc);
      
      // Logika baru: jika 2 key awal (equipment dan costCenter) sama, cek 4 key berikutnya
      // Jika 2 key awal beda, masuk kategori TIDAK DITEMUKAN
      // Jika 2 key awal sama tapi 4 key berikutnya beda, masuk kategori TIDAK SESUAI
      // Jika semua 6 key sama, masuk kategori SESUAI
      if (equipmentMatch && costCenterMatch) {
        // 2 key awal sama, cek 4 key berikutnya
        if (eqktuMatch && objectTypeMatch && funcLocMatch && funcDescMatch) {
          // Semua 6 key sama
          results.push({
            TRUTH_ROW: t.excelRow,
            TRUTH_FILEPATH: t.filePath,
            COST_CENTER: t.costCenter,
            CHECK_ROW: c.excelRow,
            CHECK_FILEPATH: c.filePath,
            CHECK_COST_CENTER: c.costCenter,
            CATATAN: '',
            STATUS: 'SESUAI',
            EQUIPMENT: t.equipment
          });
        } else {
          // 2 key awal sama tapi 4 key berikutnya beda
          // Buat catatan detail tentang parameter mana yang tidak sesuai
          const catatan = [];
          if (!eqktuMatch) catatan.push({ PARAMETER: 'EQKTU', TRUTH: t.eqktu, CURRENT_CHECK: c.eqktu });
          if (!objectTypeMatch) catatan.push({ PARAMETER: 'OBJECTTYPE', TRUTH: t.objectType, CURRENT_CHECK: c.objectType });
          if (!funcLocMatch) catatan.push({ PARAMETER: 'FUNCLOC', TRUTH: t.funcLoc, CURRENT_CHECK: c.funcLoc });
          if (!funcDescMatch) catatan.push({ PARAMETER: 'FUNCDesc', TRUTH: t.funcDesc, CURRENT_CHECK: c.funcDesc });
          
          results.push({
            TRUTH_ROW: t.excelRow,
            TRUTH_FILEPATH: t.filePath,
            COST_CENTER: t.costCenter,
            CHECK_ROW: c.excelRow,
            CHECK_FILEPATH: c.filePath,
            CHECK_COST_CENTER: c.costCenter,
            CATATAN: catatan,
            STATUS: 'TIDAK SESUAI',
            EQUIPMENT: t.equipment,
            X: t,
            Y: c
          });
        }
      } else {
        // 2 key awal beda, masuk kategori TIDAK DITEMUKAN
        let status = 'TIDAK DITEMUKAN';

        if (
          foundEntry?.COST_CENTER &&
          !foundEntry.COST_CENTER.toLowerCase().includes('stas') &&
          foundEntry.COST_CENTER.toLowerCase().includes('e') &&
          (
            foundEntry.COST_CENTER.toLowerCase().includes('2f19') ||
            foundEntry.COST_CENTER.toLowerCase().includes('2f21') ||
            foundEntry.COST_CENTER.toLowerCase().includes('5f20')
          )
        ) {
          status = 'TIDAK ADA';
        }

        results.push({
          TRUTH_ROW: t.excelRow,
          TRUTH_FILEPATH: t.filePath,
          COST_CENTER: t.costCenter,
          STATUS: status,
          EQUIPMENT: t.equipment
        });

      }
    } else if (equipmentCheckMap.has(equipmentKey)) {
      // Jika tidak ditemukan kecocokan berdasarkan kombinasi, coba hanya berdasarkan equipment
      const c = equipmentCheckMap.get(equipmentKey)[0];
      
      const cmp = (a, b) =>
        String(a || '').trim().toLowerCase() !==
        String(b || '').trim().toLowerCase();

      // Bandingkan semua 6 parameter: equipment, costCenter, eqktu, objectType, funcLoc, funcDesc
      // ObjectType hanya ada di truth data, jadi kita hanya bisa membandingkan jika ada di kedua data
      const equipmentMatch = !cmp(t.equipment, c.equipment);
      const costCenterMatch = !cmp(t.costCenter, c.costCenter);
      const eqktuMatch = !cmp(t.eqktu, c.eqktu);
      const objectTypeMatch = !cmp(t.objectType, c.objectType || '');
      const funcLocMatch = !cmp(t.funcLoc, c.funcLoc);
      const funcDescMatch = !cmp(t.funcDesc, c.funcDesc);
      
      // Logika baru: jika 2 key awal (equipment dan costCenter) sama, cek 4 key berikutnya
      // Jika 2 key awal beda, masuk kategori TIDAK DITEMUKAN
      // Jika 2 key awal sama tapi 4 key berikutnya beda, masuk kategori TIDAK SESUAI
      // Jika semua 6 key sama, masuk kategori SESUAI
      if (equipmentMatch && costCenterMatch) {
        // 2 key awal sama, cek 4 key berikutnya
        if (eqktuMatch && objectTypeMatch && funcLocMatch && funcDescMatch) {
          // Semua 6 key sama
          results.push({
            TRUTH_ROW: t.excelRow,
            TRUTH_FILEPATH: t.filePath,
            COST_CENTER: t.costCenter,
            CHECK_ROW: c.excelRow,
            CHECK_FILEPATH: c.filePath,
            CHECK_COST_CENTER: c.costCenter,
            CATATAN: '',
            STATUS: 'SESUAI',
            EQUIPMENT: t.equipment
          });
        } else {
          // 2 key awal sama tapi 4 key berikutnya beda
          // Buat catatan detail tentang parameter mana yang tidak sesuai
          const catatan = [];
          if (!eqktuMatch) catatan.push({ PARAMETER: 'EQKTU', TRUTH: t.eqktu, CURRENT_CHECK: c.eqktu });
          if (!objectTypeMatch) catatan.push({ PARAMETER: 'OBJECTTYPE', TRUTH: t.objectType, CURRENT_CHECK: c.objectType });
          if (!funcLocMatch) catatan.push({ PARAMETER: 'FUNCLOC', TRUTH: t.funcLoc, CURRENT_CHECK: c.funcLoc });
          if (!funcDescMatch) catatan.push({ PARAMETER: 'FUNCDesc', TRUTH: t.funcDesc, CURRENT_CHECK: c.funcDesc });
          
          results.push({
            TRUTH_ROW: t.excelRow,
            TRUTH_FILEPATH: t.filePath,
            COST_CENTER: t.costCenter,
            CHECK_ROW: c.excelRow,
            CHECK_FILEPATH: c.filePath,
            CHECK_COST_CENTER: c.costCenter,
            CATATAN: JSON.stringify(catatan),
            STATUS: 'TIDAK SESUAI',
            EQUIPMENT: t.equipment
          });
        }
      } else {
        // 2 key awal beda, masuk kategori TIDAK DITEMUKAN
        results.push({
          TRUTH_ROW: t.excelRow,
          TRUTH_FILEPATH: t.filePath,
          COST_CENTER: t.costCenter,
          STATUS: 'TIDAK DITEMUKAN',
          EQUIPMENT: t.equipment
        });
      }
    } else {
      // Jika tidak ditemukan kecocokan berdasarkan equipment pun, maka TIDAK DITEMUKAN
      results.push({
        TRUTH_ROW: t.excelRow,
        TRUTH_FILEPATH: t.filePath,
        COST_CENTER: t.costCenter,
        STATUS: 'TIDAK DITEMUKAN',
        EQUIPMENT: t.equipment
      });
    }
  });

  fs.writeFileSync('evaluate.json', JSON.stringify(results, null, 2));

  // Create evaluate_2.json with summary per cost center prefix
  // Create a copy of results to avoid any potential interference with the main process
  const resultsCopy = JSON.parse(JSON.stringify(results));
  
  // Extract all unique cost center prefixes from TRUTH file
  const truthCostCenterPrefixes = new Set();
  truthData.forEach(t => {
    const costCenterStr = t.costCenter.toString().trim();
    if (costCenterStr.length >= 4) {
      const prefix = costCenterStr.substring(0, 4);
      truthCostCenterPrefixes.add(prefix);
    }
  });
  
  // Initialize summary by prefix
  const summaryByPrefix = {};
  
  // Initialize all prefixes with zero counts
  truthCostCenterPrefixes.forEach(prefix => {
    summaryByPrefix[prefix] = {
      COST_CENTER_PREFIX: prefix,
      FILE: 'NO_MATCHING_FILE_FOUND',
      Sesuai: 0,
      'Tidak Sesuai': 0,
      'Tidak Ditemukan': 0
    };
  });
  
  resultsCopy.forEach(result => {
    const costCenterStr = result.COST_CENTER.toString().trim();
    if (costCenterStr.length >= 4) {
      const prefix = costCenterStr.substring(0, 4);
      
      // Initialize if not exists
      if (!summaryByPrefix[prefix]) {
        summaryByPrefix[prefix] = {
          COST_CENTER_PREFIX: prefix,
          FILE: 'NO_MATCHING_FILE_FOUND',
          Sesuai: 0,
          'Tidak Sesuai': 0,
          'Tidak Ditemukan': 0
        };
      }
      
      // Update the file path if we have a matching file
      if (result.CHECK_FILEPATH) {
        summaryByPrefix[prefix].FILE = result.CHECK_FILEPATH;
      } else {
        // For TIDAK DITEMUKAN records, try to determine the most likely file based on cost center prefix
        const recommendedFile = getRecommendedFileForCostCenter(result.COST_CENTER, costCenterPrefixMap);
        if (recommendedFile) {
          summaryByPrefix[prefix].FILE = recommendedFile;
        }
      }
      
      // Update counts
      if (result.STATUS === 'SESUAI') {
        summaryByPrefix[prefix].Sesuai++;
      } else if (result.STATUS === 'TIDAK SESUAI') {
        summaryByPrefix[prefix]['Tidak Sesuai']++;
      } else if (result.STATUS === 'TIDAK DITEMUKAN') {
        summaryByPrefix[prefix]['Tidak Ditemukan']++;
      }
    }
  });
  
  // Convert summary object to array format and sort by cost center prefix
  const summaryArray = Object.values(summaryByPrefix).sort((a, b) => a.COST_CENTER_PREFIX.localeCompare(b.COST_CENTER_PREFIX));
  
  fs.writeFileSync('evaluate_per_pks.json', JSON.stringify(summaryArray, null, 2));
  
  // Create evaluate_2.xlsx with the same summary data
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Summary');
  
  // Define columns for the worksheet
  worksheet.columns = [
    { header: 'COST_CENTER_PREFIX', key: 'COST_CENTER_PREFIX', width: 20 },
    { header: 'FILE', key: 'FILE', width: 80 },
    { header: 'Sesuai', key: 'Sesuai', width: 15 },
    { header: 'Tidak Sesuai', key: 'Tidak Sesuai', width: 15 },
    { header: 'Tidak Ditemukan', key: 'Tidak Ditemukan', width: 15 }
  ];
  
  // Add summary data to the worksheet
  summaryArray.forEach(item => {
    worksheet.addRow(item);
  });
  
  // Style the header row
  worksheet.getRow(1).eachCell(cell => {
    cell.font = { bold: true };
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFE0E0E0' }
    };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
  });
  
  // Auto-filter for the header row
  worksheet.autoFilter = 'A1:D1';
  
  // Write the Excel file
  await workbook.xlsx.writeFile('evaluate_per_pks.xlsx');

  console.log(`✔ truth.json (${truthData.length})`);
  console.log(`✔ check.json (${checkData.length})`);
  console.log(`✔ evaluate.json (${results.length})`);
  console.log(`✔ evaluate_per_pks.json (${summaryArray.length} files summarized)`);
  console.log(`✔ evaluate_per_pks.xlsx (${summaryArray.length} files summarized)`);

  await buildEvaluatedExcel();
})();
