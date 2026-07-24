/**
 * Test cục bộ bằng Node (không cần deploy lên Apps Script, không cần dữ
 * liệu vận hành thật). Nạp toàn bộ file nguồn (.js, trừ Tests.js) vào
 * chung một scope bằng Function() rồi chạy assertion - vì các file này
 * là code Apps Script thuần (biến toàn cục `QDD`), không dùng module ES.
 *
 * Chạy: node tests/run_tests.js
 */
'use strict';
const fs = require('fs');
const path = require('path');

const SRC_DIR = path.join(__dirname, '..');
const SOURCE_FILES = [
  'Config.js',
  'CommandFilter.js',
  'RampEngine.js',
  'Segments.js',
  'AreaIntegration.js',
  'CsvParser.js',
  'QddCalculator.js',
  'BatchCalculator.js',
  'MonthlyReport.js',
];

const code = SOURCE_FILES
  .map((f) => fs.readFileSync(path.join(SRC_DIR, f), 'utf8'))
  .join('\n;\n') + '\n;\nreturn QDD;\n';

const QDD = new Function(code)();

let pass = 0, fail = 0;

function approxEqual(a, b, eps) {
  eps = eps === undefined ? 1e-6 : eps;
  return Math.abs(a - b) <= eps;
}

function check(name, actual, expected, eps) {
  const ok = approxEqual(actual, expected, eps);
  if (ok) {
    pass++;
    console.log('  OK   ' + name);
  } else {
    fail++;
    console.log('  FAIL ' + name + ' -- expected ' + expected + ', got ' + actual);
  }
}

function checkTrue(name, condition) {
  if (condition) {
    pass++;
    console.log('  OK   ' + name);
  } else {
    fail++;
    console.log('  FAIL ' + name);
  }
}

// ---------------------------------------------------------------------
// 1) RampEngine + Segments + AreaIntegration: 1 lệnh, ramp trong chu kỳ 1
//    P0=100, lệnh tại t=0 đưa target=200, tốc độ mặc định 3.5 MW/phút.
//    Thời lượng ramp = |200-100|/3.5*60 = 1714.285714286 giây.
// ---------------------------------------------------------------------
console.log('1) Ramp co ban (P0=100 -> 200 MW tai t=0)');
{
  const effectiveCommands = [{ seconds: 0, p: 200 }];
  const qdd = QDD.AreaIntegration.computeDay(effectiveCommands, 100);
  // Chu ky 1: phan ramp (0..1714.2857) + phan giu (1714.2857..1800) o 200
  const H = Math.abs(200 - 100) / QDD.Config.RAMP_RATE_MW_PER_MIN * 60;
  const rampArea = (100 + 200) / 2 * H;
  const holdArea = 200 * (1800 - H);
  const expectedQdd1 = (rampArea + holdArea) / 1800;
  check('Chu ky 1 (co ramp + giu)', qdd[0], expectedQdd1, 1e-6);
  check('Chu ky 2 (giu nguyen 200 MW)', qdd[1], 200, 1e-9);
  check('Chu ky 48 (giu nguyen 200 MW)', qdd[47], 200, 1e-9);
}

// ---------------------------------------------------------------------
// 2) Ngay khong co lenh nao: Qdd = P0 het 48 chu ky (UAT-14)
// ---------------------------------------------------------------------
console.log('2) Ngay khong co lenh (UAT-14)');
{
  const qdd = QDD.AreaIntegration.computeDay([], 300);
  checkTrue('Tat ca 48 chu ky = P0', qdd.every((v) => approxEqual(v, 300, 1e-9)));
}

// ---------------------------------------------------------------------
// 3) Ngat ramp (R06): lenh thu 2 den giua ramp cua lenh 1
//    P0=0, lenh1 target=100 tai t=0 (H1=100/3.5*60=1714.2857s, ket thuc I1).
//    lenh2 tai t=900 (giua ramp1) voi target=50.
//    F2 phai = noi suy tuyen tinh tren duong ramp1 tai t=900.
// ---------------------------------------------------------------------
console.log('3) Ngat ramp - noi suy (R06)');
{
  const cmds = [{ seconds: 0, p: 100 }, { seconds: 900, p: 50 }];
  const rows = QDD.RampEngine.buildRows(cmds, 0);
  const H1 = 100 / QDD.Config.RAMP_RATE_MW_PER_MIN * 60;
  const expectedF2 = 0 + (100 - 0) * (900 - 0) / H1; // F=0, D=100, noi suy tai t=900
  check('P bat dau cua lenh 2 (noi suy tren ramp 1)', rows[1].F, expectedF2, 1e-6);
  checkTrue('Lenh 2 la RAMP_DOWN (D=50 < F noi suy)', rows[1].D < rows[1].F);
}

// ---------------------------------------------------------------------
// 4) CommandFilter - R01/R02/R03 + UAT-32 (lenh 0-0 KHONG duoc tinh)
// ---------------------------------------------------------------------
console.log('4) CommandFilter R01-R03 + UAT-32');
{
  const day = new Date(2026, 6, 7); // 07/07/2026 (thang 0-based)
  const raw = [
    // Lenh MO hop le, target 435.7
    {
      id: 'A', toMay: 'S2', csRaLenh: 435.7, csHoanThanh: 0,
      bdth: new Date(2026, 6, 7, 10, 4, 0), hoanThanh: 1, dungLenh: false, nguonLenh: 'MO',
    },
    // Lenh 0-0 (trip) - PHAI bi loai theo UAT-32
    {
      id: 'B', toMay: 'S2', csRaLenh: 0, csHoanThanh: 0,
      bdth: new Date(2026, 6, 7, 10, 34, 0), hoanThanh: 1, dungLenh: false, nguonLenh: 'MO',
    },
    // Lenh SO hop le
    {
      id: 'C', toMay: 'S1', csRaLenh: 100, csHoanThanh: 120,
      bdth: new Date(2026, 6, 7, 8, 0, 0), hoanThanh: 1, dungLenh: false, nguonLenh: 'SO',
    },
    // Khac ngay - phai bi loai
    {
      id: 'D', toMay: 'S2', csRaLenh: 200, csHoanThanh: 200,
      bdth: new Date(2026, 6, 8, 8, 0, 0), hoanThanh: 1, dungLenh: false, nguonLenh: 'MO',
    },
  ];

  const s2 = QDD.CommandFilter.selectEffective(raw, day, 'S2');
  checkTrue('S2: chi 1 lenh hop le (loai lenh 0-0 va khac ngay)', s2.length === 1);
  checkTrue('S2: lenh con lai la id A', s2.length === 1 && s2[0].id === 'A');
  check('S2: P hieu luc = CS ra lenh (R02)', s2.length === 1 ? s2[0].p : NaN, 435.7);

  const s1 = QDD.CommandFilter.selectEffective(raw, day, 'S1');
  checkTrue('S1: 1 lenh SO hop le', s1.length === 1);
  check('S1: P hieu luc = CS hoan thanh (R01)', s1.length === 1 ? s1[0].p : NaN, 120);
}

// ---------------------------------------------------------------------
// 5) CsvParser.extractKwhGiao
// ---------------------------------------------------------------------
console.log('5) CsvParser.extractKwhGiao');
{
  const period = Array.from({ length: 48 }, (_, i) => String(100 + i));
  const rows = [
    ['07-07-26', 'KwhGiao'].concat(period),
    ['07-07-26', 'KwhNhan'].concat(period.map(() => '0')),
  ];
  const values = QDD.CsvParser.extractKwhGiao(rows);
  checkTrue('Doc dung 48 gia tri', values.length === 48);
  check('Gia tri dau tien', values[0], 100);
  check('Gia tri cuoi cung', values[47], 147);
}

// ---------------------------------------------------------------------
// 6) QddCalculator - Qdu trong/ngoai dai +-3% (R13-R14)
// ---------------------------------------------------------------------
console.log('6) QddCalculator R08-R14');
{
  // Khong co lenh nao trong ngay -> Qdd giu nguyen P0=200 MW ca 48 chu ky.
  // Qdc(MWh) = KwhGiao/1000, P_Qdc = Qdc*2. Muon P_Qdc = 200 (bang Qdd) thi KwhGiao = 200/2*1000 = 100000
  const qdcInBand = Array(48).fill(100000);
  const qmpAny = Array(48).fill(50000); // khong quan trong khi trong dai (Qdu=0)

  const resultsInBand = QDD.QddCalculator.calculateDay({
    effectiveCommands: [], p0: 200, qdc48: qdcInBand, qmp48: qmpAny,
  });
  checkTrue('Trong dai +-3%: Qdu = 0 het 48 chu ky', resultsInBand.every((r) => r.qdu === 0 && r.dauHieu === 'trong ±3%'));

  // Ngoai dai: KwhGiao thap han nhieu -> P_Qdc < nguong duoi -> "am"
  const qdcOutOfBand = Array(48).fill(10000); // P_Qdc = 10000/1000*2 = 20 << nguong duoi (194)
  const qmpForOut = Array(48).fill(300000); // Qmp = 300 MWh
  const resultsOut = QDD.QddCalculator.calculateDay({
    effectiveCommands: [], p0: 200, qdc48: qdcOutOfBand, qmp48: qmpForOut,
  });
  const qddV = 200 / 2 * QDD.Config.QDD_V_COEF;
  const expectedQdu = 300 - qddV;
  checkTrue('Ngoai dai (duoi): dau hieu = am', resultsOut.every((r) => r.dauHieu === 'âm'));
  check('Ngoai dai: Qdu = Qmp - Qdd_V', resultsOut[0].qdu, expectedQdu, 1e-6);
}

// ---------------------------------------------------------------------
// 7) BatchCalculator - nhieu ngay, nhieu to may cung luc
// ---------------------------------------------------------------------
console.log('7) BatchCalculator - nhieu ngay/nhieu to may');
{
  const qdcFlat = Array(48).fill(100000); // -> P_Qdc = 200 MW, trong dai neu Qdd=200
  const qmpFlat = Array(48).fill(50000);

  const dayInputs = [
    { date: '2026-07-01', unit: 'S1', effectiveCommands: [], p0: 200, qdc48: qdcFlat, qmp48: qmpFlat },
    { date: '2026-07-01', unit: 'S2', effectiveCommands: [], p0: 150, qdc48: qdcFlat, qmp48: qmpFlat },
    { date: '2026-07-02', unit: 'S1', effectiveCommands: [], p0: 200, qdc48: qdcFlat, qmp48: qmpFlat },
    // Ngay co du lieu loi (qdc48 thieu gia tri) - khong duoc lam hong ca lo
    { date: '2026-07-03', unit: 'S1', effectiveCommands: [], p0: 200, qdc48: [1, 2, 3], qmp48: qmpFlat },
  ];

  const results = QDD.BatchCalculator.calculateMultiple(dayInputs);
  checkTrue('Tra ve du 4 ket qua (dung thu tu)', results.length === 4);
  checkTrue('01/07 S1: tinh thanh cong, 48 chu ky', results[0].periods && results[0].periods.length === 48);
  checkTrue('01/07 S2: doc lap voi S1 (P0=150 khac P0=200)', results[1].periods[0].qdd === 150);
  checkTrue('03/07: loi duoc bat, khong lam dung chuong trinh', !!results[3].error && results[3].periods === null);

  // buildDayUnitInputs
  const raw = [{
    id: 'X', toMay: 'S1', csRaLenh: 250, csHoanThanh: 250,
    bdth: new Date(2026, 6, 5, 6, 0, 0), hoanThanh: 1, dungLenh: false, nguonLenh: 'MO',
  }];
  const built = QDD.BatchCalculator.buildDayUnitInputs(raw, [new Date(2026, 6, 5)], ['S1', 'S2']);
  checkTrue('buildDayUnitInputs: 1 ngay x 2 to may = 2 ket qua', built.length === 2);
  const s1built = built.filter((b) => b.unit === 'S1')[0];
  checkTrue('buildDayUnitInputs: S1 co dung 1 lenh hieu luc', s1built.effectiveCommands.length === 1);
}

// ---------------------------------------------------------------------
// 8) MonthlyReport - tong hop nhieu ngay, khong can snapshot
// ---------------------------------------------------------------------
console.log('8) MonthlyReport');
{
  const qdcFlat = Array(48).fill(100000);
  const qmpFlat = Array(48).fill(50000);
  const dayInputs = [
    { date: '2026-07-01', unit: 'S1', effectiveCommands: [], p0: 200, qdc48: qdcFlat, qmp48: qmpFlat },
    { date: '2026-07-02', unit: 'S1', effectiveCommands: [], p0: 200, qdc48: qdcFlat, qmp48: qmpFlat },
    { date: '2026-07-03', unit: 'S1', effectiveCommands: [], p0: 200, qdc48: [1, 2, 3], qmp48: qmpFlat }, // loi
  ];
  const dayResults = QDD.BatchCalculator.calculateMultiple(dayInputs);
  const report = QDD.MonthlyReport.aggregate(dayResults);

  checkTrue('perDay co du 3 dong', report.perDay.length === 3);
  checkTrue('Ngay loi duoc ghi nhan rieng, khong tinh vao tong', report.perDay[2].error !== undefined);
  check('tongHop.soNgay = 2 (bo qua ngay loi)', report.tongHop.soNgay, 2);
  check('tongHop.soNgayLoi = 1', report.tongHop.soNgayLoi, 1);
  checkTrue('perUnit.S1 ton tai va cong don dung 2 ngay', report.perUnit.S1.soNgay === 2);
  // Moi ngay Qdc = sum(100 MWh x 48) = 4800; 2 ngay = 9600
  check('tongHop.tongQdc = 2 ngay x 48 chu ky x 100 MWh', report.tongHop.tongQdc, 9600, 1e-6);
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail > 0 ? 1 : 0);
