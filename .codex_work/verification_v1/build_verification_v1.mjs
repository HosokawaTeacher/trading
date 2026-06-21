import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const root = process.cwd();
const inputPath = path.join(root, "outputs/trading_sim_20260614/5_model_trading_simulation_1month.xlsx");
const qaDir = path.join(root, ".codex_work/verification_v1/qa");
const tempPath = path.join(root, ".codex_work/verification_v1/verification_v1_tmp.xlsx");
await fs.mkdir(qaDir, { recursive: true });

const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);

const colors = {
  navy: "#1F4E78",
  blue: "#D9EAF7",
  paleBlue: "#EAF3F8",
  green: "#E2F0D9",
  yellow: "#FFF2CC",
  red: "#F4CCCC",
  gray: "#E7E6E6",
  darkGray: "#666666",
  border: "#B7C9D6",
  input: "#DDEBF7",
};

function resetSheet(sheet) {
  const used = sheet.getUsedRange();
  if (used) {
    used.unmerge();
    used.clear({ applyTo: "all" });
  }
  sheet.deleteAllDrawings();
  sheet.showGridLines = false;
}

function title(sheet, range, text) {
  const r = sheet.getRange(range);
  r.merge();
  r.values = [[text]];
  r.format.fill = colors.navy;
  r.format.font = { bold: true, color: "#FFFFFF", size: 14 };
  r.format.verticalAlignment = "center";
  r.format.rowHeight = 26;
}

function section(sheet, range, text) {
  const r = sheet.getRange(range);
  r.merge();
  r.values = [[text]];
  r.format.fill = colors.blue;
  r.format.font = { bold: true, color: "#1F1F1F" };
  r.format.borders = { preset: "outside", style: "thin", color: colors.border };
}

function header(range) {
  range.format.fill = colors.navy;
  range.format.font = { bold: true, color: "#FFFFFF" };
  range.format.wrapText = true;
  range.format.horizontalAlignment = "center";
  range.format.verticalAlignment = "center";
  range.format.borders = { preset: "all", style: "thin", color: colors.border };
}

function body(range) {
  range.format.borders = { preset: "all", style: "thin", color: colors.border };
  range.format.verticalAlignment = "center";
}

function inputStyle(range) {
  range.format.fill = colors.input;
  range.format.font = { color: "#0070C0" };
}

function setWidths(sheet, widths) {
  widths.forEach((width, i) => {
    sheet.getRangeByIndexes(0, i, 1, 1).format.columnWidth = width;
  });
}

const main = workbook.worksheets.getItem("ステップ2_ルール検証");
const forward = workbook.worksheets.getOrAdd("前進検証_v1");
const history = workbook.worksheets.getOrAdd("過去検証_v1");
const live = workbook.worksheets.getOrAdd("実運用比較_v1");
const definitions = workbook.worksheets.getOrAdd("判定ルール_v1");
for (const sheet of [main, forward, history, live, definitions]) resetSheet(sheet);

// Main dashboard / fixed assumptions.
title(main, "A1:P1", "ステップ2：損切り・利確ルールの効き目検証（仕様v1）");
main.getRange("A2:P2").merge();
main.getRange("A2").values = [["前提日 2026-06-21／基準日 2026-06-12終値／前進・過去検証は混合しない／n=2のため現在は動作確認のみ"]];
main.getRange("A2:P2").format.font = { italic: true, color: colors.darkGray };

section(main, "A4:D4", "■ 固定前提（v1）");
main.getRange("A5:D12").values = [
  ["項目", "値", "単位/状態", "備考"],
  ["基準日", new Date("2026-06-12T00:00:00+09:00"), "終値", "金曜日・同一時点で建値とベンチを固定"],
  ["評価日", new Date("2026-06-19T00:00:00+09:00"), "終値", "前進検証の現時点"],
  ["片道コスト率", 0.0005, "0.05%", "往復0.1%；保有中は評価日決済コストも見込む"],
  ["DD警告ライン", -0.08, "対投下資金", "運用を点検"],
  ["DD失格ライン", -0.12, "対投下資金", "不採用候補"],
  ["MAE要注意", -15000, "円/取引", "想定損失1万円の1.5倍"],
  ["取引単位", "1建玉=1取引", "確定", "部分利確は合算損益で判定"],
];
header(main.getRange("A5:D5"));
body(main.getRange("A6:D12"));
main.getRange("B6:B7").format.numberFormat = "yyyy-mm-dd";
main.getRange("B8:B10").format.numberFormat = "0.0%;[Red](0.0%)";
main.getRange("B11").format.numberFormat = "#,##0;[Red](#,##0)";

section(main, "F4:I4", "■ 比較群（出口ルール単体）");
main.getRange("F5:I9").values = [
  ["群", "出口", "再投資", "評価時点"],
  ["A", "終値条件確認後、翌営業日始値で全決済", "なし", "同一固定期末"],
  ["B", "buy&hold（評価日終値で決済換算）", "なし", "同一固定期末"],
  ["C", "第1利確で半分、残りは評価日まで", "なし", "同一固定期末"],
  ["N", "日経平均buy&hold", "-", "2026-06-12終値起点"],
];
header(main.getRange("F5:I5"));
body(main.getRange("F6:I9"));

section(main, "K4:P4", "■ データ出典（2026-06-21取得）");
main.getRange("K5:P9").values = [
  ["対象", "基準日値", "評価日値", "基準日", "評価日", "出典URL"],
  ["8306 三菱UFJ", 3162, 3278, new Date("2026-06-12T00:00:00+09:00"), new Date("2026-06-19T00:00:00+09:00"), "https://finance.yahoo.co.jp/quote/8306.T/history"],
  ["9433 KDDI", 2769, 2721.5, new Date("2026-06-12T00:00:00+09:00"), new Date("2026-06-19T00:00:00+09:00"), "https://finance.yahoo.co.jp/quote/9433.T/history"],
  ["日経平均", 66020.04, 71250.06, new Date("2026-06-12T00:00:00+09:00"), new Date("2026-06-19T00:00:00+09:00"), "https://indexes.nikkei.co.jp/nkave/archives/summary/"],
  ["検証仕様", "v1", "固定", new Date("2026-06-21T00:00:00+09:00"), null, "VERIFICATION_SPEC_v1_20260621.md"],
];
header(main.getRange("K5:P5"));
body(main.getRange("K6:P9"));
main.getRange("L6:M8").format.numberFormat = "#,##0.00";
main.getRange("N6:O9").format.numberFormat = "yyyy-mm-dd";

section(main, "A14:P14", "■ 前進検証・群別集計（n<10は動作確認、加重スコアは算出保留）");
main.getRange("A15:P18").values = [
  ["群", "総損益(コスト後)", "対投下リターン", "最大DD", "DD判定", "日経リターン", "超過リターン", "期待値/取引", "実現損益比", "勝率(補助)", "n", "期待値スコア40%", "DDスコア25%", "超過スコア20%", "損益比スコア15%", "加重スコア"],
  ["A", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null],
  ["B", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null],
  ["C", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null],
];
header(main.getRange("A15:P15"));
body(main.getRange("A16:P18"));
for (let row = 16; row <= 18; row++) {
  const src = row + 26; // Forward summary rows 42:44.
  main.getRange(`B${row}:K${row}`).formulas = [[
    `='前進検証_v1'!G${src}`,
    `='前進検証_v1'!H${src}`,
    `='前進検証_v1'!I${src}`,
    `='前進検証_v1'!J${src}`,
    `='前進検証_v1'!K${src}`,
    `='前進検証_v1'!L${src}`,
    `='前進検証_v1'!M${src}`,
    `='前進検証_v1'!N${src}`,
    `='前進検証_v1'!O${src}`,
    `='前進検証_v1'!P${src}`,
  ]];
  main.getRange(`L${row}`).formulas = [[`=IF(K${row}<10,"",MAX(0,MIN(100,50+H${row}/10000*50)))`]];
  main.getRange(`M${row}`).formulas = [[`=IF(K${row}<10,"",MAX(0,MIN(100,(D${row}-$B$10)/ABS($B$10)*100)))`]];
  main.getRange(`N${row}`).formulas = [[`=IF(K${row}<10,"",MAX(0,MIN(100,50+G${row}*500)))`]];
  main.getRange(`O${row}`).formulas = [[`=IF(K${row}<10,"",MAX(0,MIN(100,I${row}/3*100)))`]];
  main.getRange(`P${row}`).formulas = [[`=IF(COUNT(L${row}:O${row})<4,"",SUMPRODUCT(L${row}:O${row},$L$20:$O$20))`]];
}
main.getRange("L20:O20").values = [[0.4, 0.25, 0.2, 0.15]];
main.getRange("K20").values = [["重み"]];
main.getRange("K20:O20").format.fill = colors.yellow;
main.getRange("L20:O20").format.numberFormat = "0%";
main.getRange("B16:B18").format.numberFormat = "#,##0;[Red](#,##0)";
main.getRange("C16:G18").format.numberFormat = "0.00%;[Red](0.00%)";
main.getRange("H16:H18").format.numberFormat = "#,##0;[Red](#,##0)";
main.getRange("I16:I18").format.numberFormat = "0.00x;[Red](0.00x)";
main.getRange("J16:J18").format.numberFormat = "0.0%";
main.getRange("L16:P18").format.numberFormat = "0.0";

section(main, "A22:P22", "■ 判定注意");
main.getRange("A23:P27").merge(true);
main.getRange("A23:A27").values = [
  ["・10件=動作確認、30件以上=暫定評価、50〜100件=統計評価の目安。前進と過去は混合集計しない。"],
  ["・加重スコアはn≥10のときのみ参考表示。期待値は1万円=1R、超過は±10%、損益比は3.0を100点とする線形正規化。"],
  ["・DDは日次終値の資産曲線のピークからの下落率。MAEは個別取引の最大逆行額で、最大含み損/DDと区別。"],
  ["・同日に利確と損切りが両方到達し、日足で順序不明な場合は保守的に損切り優先。"],
  ["・売買推奨ではなく模擬検証。最終判断と発注はユーザーが行う。"],
];
main.getRange("A23:P27").format.wrapText = true;
main.getRange("A23:P27").format.borders = { preset: "outside", style: "thin", color: colors.border };
setWidths(main, [9, 16, 15, 12, 12, 13, 14, 14, 13, 12, 8, 13, 11, 12, 12, 13]);
main.freezePanes.freezeRows(2);

// Forward test: auditable trade and daily equity calculations.
title(forward, "A1:AA1", "前進検証_v1（出口ルール比較・再投資なし）");
forward.getRange("A2:AA2").merge();
forward.getRange("A2").values = [["基準日 2026-06-12終値／評価日 2026-06-19終値／往復コスト0.1%／保有中も評価日決済換算／青字セルが入力"]];
section(forward, "A4:AA4", "■ 1建玉=1取引（部分利確は合算）");
forward.getRange("A5:AA7").values = [
  ["ID", "コード", "銘柄", "株数", "建値日", "建値", "第1利確", "発火判定日(終値)", "翌営業日", "翌日始値", "評価日終値", "期中最安値", "A決済値", "B決済値", "C半分決済値", "A粗損益", "Aコスト", "A純損益", "B粗損益", "Bコスト", "B純損益", "C粗損益", "Cコスト", "C純損益", "MAE", "MAE判定", "メモ"],
  [1, "8306", "三菱UFJ FG", 100, new Date("2026-06-12T00:00:00+09:00"), 3162, 3300, new Date("2026-06-18T00:00:00+09:00"), new Date("2026-06-19T00:00:00+09:00"), 3370, 3278, 3186, null, null, null, null, null, null, null, null, null, null, null, null, null, null, "6/18終値で利確条件確認→6/19始値"],
  [2, "9433", "KDDI", 100, new Date("2026-06-12T00:00:00+09:00"), 2769, 2800, null, null, null, 2721.5, 2670, null, null, null, null, null, null, null, null, null, null, null, null, null, null, "条件未発火。評価日終値で決済換算"],
];
header(forward.getRange("A5:AA5"));
body(forward.getRange("A6:AA7"));
inputStyle(forward.getRange("A6:L7"));
for (let row = 6; row <= 7; row++) {
  forward.getRange(`M${row}:O${row}`).formulas = [[
    `=IF(J${row}<>"",J${row},K${row})`,
    `=K${row}`,
    `=IF(J${row}<>"",J${row},K${row})`,
  ]];
  forward.getRange(`P${row}:Y${row}`).formulas = [[
    `=(M${row}-F${row})*D${row}`,
    `=(F${row}*D${row}+M${row}*D${row})*'ステップ2_ルール検証'!$B$8`,
    `=P${row}-Q${row}`,
    `=(N${row}-F${row})*D${row}`,
    `=(F${row}*D${row}+N${row}*D${row})*'ステップ2_ルール検証'!$B$8`,
    `=S${row}-T${row}`,
    `=IF(J${row}<>"",(O${row}-F${row})*D${row}/2+(K${row}-F${row})*D${row}/2,(K${row}-F${row})*D${row})`,
    `=(F${row}*D${row}+IF(J${row}<>"",O${row}*D${row}/2+K${row}*D${row}/2,K${row}*D${row}))*'ステップ2_ルール検証'!$B$8`,
    `=V${row}-W${row}`,
    `=MIN(0,(L${row}-F${row})*D${row})`,
  ]];
  forward.getRange(`Z${row}`).formulas = [[`=IF(Y${row}<='ステップ2_ルール検証'!$B$11,"要注意","OK")`]];
}
forward.getRange("E6:E7").format.numberFormat = "yyyy-mm-dd";
forward.getRange("F6:G7").format.numberFormat = "#,##0.00";
forward.getRange("H6:I7").format.numberFormat = "yyyy-mm-dd";
forward.getRange("J6:O7").format.numberFormat = "#,##0.00";
forward.getRange("P6:Y7").format.numberFormat = "#,##0;[Red](#,##0)";
forward.getRange("Z6:Z7").format.fill = colors.green;

section(forward, "A10:Q10", "■ 日次資産曲線（終値DD主指標）");
forward.getRange("A11:Q17").values = [
  ["日付", "8306終値", "9433終値", "日経終値", "A粗損益", "A純資産", "Aピーク", "A DD", "B粗損益", "B純資産", "Bピーク", "B DD", "C粗損益", "C純資産", "Cピーク", "C DD", "日経リターン"],
  [new Date("2026-06-12T00:00:00+09:00"), 3162, 2769, 66020.04, null, null, null, null, null, null, null, null, null, null, null, null, null],
  [new Date("2026-06-15T00:00:00+09:00"), 3242, 2718, 68199.34, null, null, null, null, null, null, null, null, null, null, null, null, null],
  [new Date("2026-06-16T00:00:00+09:00"), 3228, 2714.5, 69404.50, null, null, null, null, null, null, null, null, null, null, null, null, null],
  [new Date("2026-06-17T00:00:00+09:00"), 3272, 2728, 69902.25, null, null, null, null, null, null, null, null, null, null, null, null, null],
  [new Date("2026-06-18T00:00:00+09:00"), 3374, 2715.5, 71053.49, null, null, null, null, null, null, null, null, null, null, null, null, null],
  [new Date("2026-06-19T00:00:00+09:00"), 3278, 2721.5, 71250.06, null, null, null, null, null, null, null, null, null, null, null, null, null],
];
header(forward.getRange("A11:Q11"));
body(forward.getRange("A12:Q17"));
inputStyle(forward.getRange("A12:D17"));
for (let row = 12; row <= 17; row++) {
  const isExit = row === 17;
  forward.getRange(`E${row}:Q${row}`).formulas = [[
    isExit ? "=(3370-3162)*100+(C17-2769)*100" : `=(B${row}-3162)*100+(C${row}-2769)*100`,
    isExit ? "=SUMPRODUCT($D$6:$D$7,$F$6:$F$7)+$E17-((3162*100+2769*100)+(3370*100+C17*100))*'ステップ2_ルール検証'!$B$8" : `=SUMPRODUCT($D$6:$D$7,$F$6:$F$7)+E${row}-((3162*100+2769*100)+(B${row}*100+C${row}*100))*'ステップ2_ルール検証'!$B$8`,
    `=MAX($F$12:F${row})`,
    `=F${row}/G${row}-1`,
    `=(B${row}-3162)*100+(C${row}-2769)*100`,
    `=SUMPRODUCT($D$6:$D$7,$F$6:$F$7)+I${row}-((3162*100+2769*100)+(B${row}*100+C${row}*100))*'ステップ2_ルール検証'!$B$8`,
    `=MAX($J$12:J${row})`,
    `=J${row}/K${row}-1`,
    isExit ? "=(3370-3162)*50+(B17-3162)*50+(C17-2769)*100" : `=(B${row}-3162)*100+(C${row}-2769)*100`,
    isExit ? "=SUMPRODUCT($D$6:$D$7,$F$6:$F$7)+M17-((3162*100+2769*100)+(3370*50+B17*50+C17*100))*'ステップ2_ルール検証'!$B$8" : `=SUMPRODUCT($D$6:$D$7,$F$6:$F$7)+M${row}-((3162*100+2769*100)+(B${row}*100+C${row}*100))*'ステップ2_ルール検証'!$B$8`,
    `=MAX($N$12:N${row})`,
    `=N${row}/O${row}-1`,
    `=D${row}/$D$12-1`,
  ]];
}
forward.getRange("A12:A17").format.numberFormat = "yyyy-mm-dd";
forward.getRange("B12:D17").format.numberFormat = "#,##0.00";
forward.getRange("E12:G17").format.numberFormat = "#,##0;[Red](#,##0)";
forward.getRange("I12:K17").format.numberFormat = "#,##0;[Red](#,##0)";
forward.getRange("M12:O17").format.numberFormat = "#,##0;[Red](#,##0)";
forward.getRange("H12:H17").format.numberFormat = "0.00%;[Red](0.00%)";
forward.getRange("L12:L17").format.numberFormat = "0.00%;[Red](0.00%)";
forward.getRange("P12:Q17").format.numberFormat = "0.00%;[Red](0.00%)";

section(forward, "A20:K20", "■ 補助：日中安値・MAE");
forward.getRange("A21:K27").values = [
  ["日付", "8306安値", "9433安値", "8306 MAE", "9433 MAE", "合計日中逆行", "備考", null, null, null, null],
  [new Date("2026-06-12T00:00:00+09:00"), null, null, 0, 0, 0, "建値は終値のため同日安値は建後MAEに含めない", null, null, null, null],
  [new Date("2026-06-15T00:00:00+09:00"), 3225, 2701.5, null, null, null, null, null, null, null, null],
  [new Date("2026-06-16T00:00:00+09:00"), 3186, 2670, null, null, null, null, null, null, null, null],
  [new Date("2026-06-17T00:00:00+09:00"), 3260, 2708.5, null, null, null, null, null, null, null, null],
  [new Date("2026-06-18T00:00:00+09:00"), 3317, 2688.5, null, null, null, null, null, null, null, null],
  [new Date("2026-06-19T00:00:00+09:00"), 3258, 2681.5, null, null, null, null, null, null, null, null],
];
header(forward.getRange("A21:G21"));
body(forward.getRange("A22:G27"));
inputStyle(forward.getRange("A22:C27"));
for (let row = 23; row <= 27; row++) {
  forward.getRange(`D${row}:F${row}`).formulas = [[
    `=MIN(0,(B${row}-$F$6)*$D$6)`,
    `=MIN(0,(C${row}-$F$7)*$D$7)`,
    `=D${row}+E${row}`,
  ]];
}
forward.getRange("A22:A27").format.numberFormat = "yyyy-mm-dd";
forward.getRange("B23:C27").format.numberFormat = "#,##0.00";
forward.getRange("D22:F27").format.numberFormat = "#,##0;[Red](#,##0)";

section(forward, "A30:K30", "■ 集計（メインシート参照先）");
forward.getRange("A31:K44").values = [
  ["項目", "A", "B", "C", null, null, null, null, null, null, null],
  ["投下資金", null, null, null, null, null, null, null, null, null, null],
  ["総損益", null, null, null, null, null, null, null, null, null, null],
  ["対投下リターン", null, null, null, null, null, null, null, null, null, null],
  ["最大DD", null, null, null, null, null, null, null, null, null, null],
  ["DD判定", null, null, null, null, null, null, null, null, null, null],
  ["日経リターン", null, null, null, null, null, null, null, null, null, null],
  ["超過リターン", null, null, null, null, null, null, null, null, null, null],
  ["期待値/取引", null, null, null, null, null, null, null, null, null, null],
  ["実現損益比", null, null, null, null, null, null, null, null, null, null],
  ["勝率", null, null, null, null, null, null, null, null, null, null],
  ["n", null, null, null, null, null, null, null, null, null, null],
  [null, null, null, null, null, null, null, null, null, null, null],
  [null, null, null, null, null, null, null, null, null, null, null],
];
header(forward.getRange("A31:D31"));
body(forward.getRange("A32:D42"));
forward.getRange("B32:D32").formulas = [["=SUMPRODUCT($D$6:$D$7,$F$6:$F$7)", "=B32", "=B32"]];
forward.getRange("B33:D33").formulas = [["=SUM($R$6:$R$7)", "=SUM($U$6:$U$7)", "=SUM($X$6:$X$7)"]];
forward.getRange("B34:D34").formulas = [["=B33/B32", "=C33/C32", "=D33/D32"]];
forward.getRange("B35:D35").formulas = [["=MIN($H$12:$H$17)", "=MIN($L$12:$L$17)", "=MIN($P$12:$P$17)"]];
for (const col of ["B", "C", "D"]) {
  forward.getRange(`${col}36`).formulas = [[`=IF(${col}35<='ステップ2_ルール検証'!$B$10,"失格候補",IF(${col}35<='ステップ2_ルール検証'!$B$9,"警告","OK"))`]];
}
forward.getRange("B37:D37").formulas = [["=$Q$17", "=$Q$17", "=$Q$17"]];
forward.getRange("B38:D38").formulas = [["=B34-B37", "=C34-C37", "=D34-D37"]];
forward.getRange("B39:D39").formulas = [["=AVERAGE($R$6:$R$7)", "=AVERAGE($U$6:$U$7)", "=AVERAGE($X$6:$X$7)"]];
forward.getRange("B40:D40").formulas = [[
  "=IFERROR(AVERAGEIF($R$6:$R$7,\">0\")/ABS(AVERAGEIF($R$6:$R$7,\"<0\")),0)",
  "=IFERROR(AVERAGEIF($U$6:$U$7,\">0\")/ABS(AVERAGEIF($U$6:$U$7,\"<0\")),0)",
  "=IFERROR(AVERAGEIF($X$6:$X$7,\">0\")/ABS(AVERAGEIF($X$6:$X$7,\"<0\")),0)",
]];
forward.getRange("B41:D41").formulas = [["=COUNTIF($R$6:$R$7,\">0\")/COUNT($R$6:$R$7)", "=COUNTIF($U$6:$U$7,\">0\")/COUNT($U$6:$U$7)", "=COUNTIF($X$6:$X$7,\">0\")/COUNT($X$6:$X$7)"]];
forward.getRange("B42:D42").formulas = [["=COUNT($R$6:$R$7)", "=COUNT($U$6:$U$7)", "=COUNT($X$6:$X$7)"]];

// Rows 42:44 are a compact transposed block consumed by the dashboard.
forward.getRange("F41:P44").values = [
  ["群", "総損益", "リターン", "最大DD", "DD判定", "日経", "超過", "期待値", "損益比", "勝率", "n"],
  ["A", null, null, null, null, null, null, null, null, null, null],
  ["B", null, null, null, null, null, null, null, null, null, null],
  ["C", null, null, null, null, null, null, null, null, null, null],
];
header(forward.getRange("F41:P41"));
body(forward.getRange("F42:P44"));
for (let row = 42; row <= 44; row++) {
  const sourceCol = ["B", "C", "D"][row - 42];
  forward.getRange(`G${row}:P${row}`).formulas = [[
    `=${sourceCol}$33`, `=${sourceCol}$34`, `=${sourceCol}$35`, `=${sourceCol}$36`, `=${sourceCol}$37`,
    `=${sourceCol}$38`, `=${sourceCol}$39`, `=${sourceCol}$40`, `=${sourceCol}$41`, `=${sourceCol}$42`,
  ]];
}
forward.getRange("B32:D33").format.numberFormat = "#,##0;[Red](#,##0)";
forward.getRange("B34:D35").format.numberFormat = "0.00%;[Red](0.00%)";
forward.getRange("B37:D38").format.numberFormat = "0.00%;[Red](0.00%)";
forward.getRange("B39:D39").format.numberFormat = "#,##0;[Red](#,##0)";
forward.getRange("B40:D40").format.numberFormat = "0.00x;[Red](0.00x)";
forward.getRange("B41:D41").format.numberFormat = "0.0%";
forward.getRange("G42:G44").format.numberFormat = "#,##0;[Red](#,##0)";
forward.getRange("H42:L44").format.numberFormat = "0.00%;[Red](0.00%)";
forward.getRange("M42:M44").format.numberFormat = "#,##0;[Red](#,##0)";
forward.getRange("N42:N44").format.numberFormat = "0.00x;[Red](0.00x)";
forward.getRange("O42:O44").format.numberFormat = "0.0%";
setWidths(forward, [11, 10, 20, 9, 12, 11, 11, 14, 12, 11, 12, 12, 11, 11, 13, 12, 11, 12, 12, 11, 12, 12, 11, 12, 11, 11, 34]);
forward.freezePanes.freezeRows(5);

// Historical test template: deliberately separate, with bias controls and exclusions.
title(history, "A1:Z1", "過去検証_v1（ウォークフォワード補助・前進と混合禁止）");
history.getRange("A2:Z2").merge();
history.getRange("A2").values = [["当時入手可能な情報のみを使用。対象母集団・選外銘柄・選外理由も記録し、生存者/後知恵バイアスを明示する。青字=入力。"]];
history.getRange("A4:Z4").values = [[
  "ID", "ルール版", "判定日", "コード", "銘柄", "当時母集団", "選定/選外", "選外理由", "テーマ群", "建値日", "株数", "建値", "ストップ", "第1利確", "発火日", "翌日始値", "評価日", "評価値", "期中最安", "粗損益", "コスト", "純損益", "MAE", "勝敗", "バイアス注記", "出典URL"
]];
header(history.getRange("A4:Z4"));
body(history.getRange("A5:Z104"));
inputStyle(history.getRange("A5:S104"));
inputStyle(history.getRange("Y5:Z104"));
for (let row = 5; row <= 104; row++) {
  history.getRange(`T${row}:X${row}`).formulas = [[
    `=IF(OR(G${row}<>"選定",L${row}="",R${row}=""),"",(R${row}-L${row})*K${row})`,
    `=IF(T${row}="","",(L${row}*K${row}+R${row}*K${row})*'ステップ2_ルール検証'!$B$8)`,
    `=IF(T${row}="","",T${row}-U${row})`,
    `=IF(OR(S${row}="",L${row}=""),"",MIN(0,(S${row}-L${row})*K${row}))`,
    `=IF(V${row}="","",IF(V${row}>0,"勝ち",IF(V${row}<0,"負け","引分")))`,
  ]];
}
history.getRange("C5:C104").format.numberFormat = "yyyy-mm-dd";
history.getRange("J5:J104").format.numberFormat = "yyyy-mm-dd";
history.getRange("O5:O104").format.numberFormat = "yyyy-mm-dd";
history.getRange("Q5:Q104").format.numberFormat = "yyyy-mm-dd";
history.getRange("L5:N104").format.numberFormat = "#,##0.00";
history.getRange("P5:S104").format.numberFormat = "#,##0.00";
history.getRange("T5:W104").format.numberFormat = "#,##0;[Red](#,##0)";
history.getRange("B5:B104").dataValidation = { rule: { type: "list", values: ["v1"] } };
history.getRange("G5:G104").dataValidation = { rule: { type: "list", values: ["選定", "選外"] } };
setWidths(history, [8, 10, 12, 10, 18, 24, 12, 30, 18, 12, 8, 11, 11, 11, 12, 11, 12, 11, 11, 11, 11, 11, 11, 10, 34, 38]);
history.freezePanes.freezeRows(4);

// Live portfolio comparison: reinvestment allowed and kept separate from the exit-rule test.
title(live, "A1:N1", "実運用比較_v1（空き資金の再投入あり・出口ルール比較と別建て）");
live.getRange("A2:N2").merge();
live.getRange("A2").values = [["このシートは実運用の資金推移用。出口ルール単体A/B/Cと混合集計しない。青字=入力。"]];
live.getRange("A4:N4").values = [["日付", "期首現金", "新規投入", "売却回収", "コスト", "期末現金", "保有時価", "総資産", "ピーク", "DD", "再投入銘柄", "選定理由", "同日/同テーマ群", "備考"]];
header(live.getRange("A4:N4"));
body(live.getRange("A5:N104"));
inputStyle(live.getRange("A5:E104"));
inputStyle(live.getRange("G5:G104"));
inputStyle(live.getRange("K5:N104"));
for (let row = 5; row <= 104; row++) {
  live.getRange(`F${row}`).formulas = [[`=IF(A${row}="","",B${row}-C${row}+D${row}-E${row})`]];
  live.getRange(`H${row}`).formulas = [[`=IF(A${row}="","",F${row}+G${row})`]];
  live.getRange(`I${row}`).formulas = [[`=IF(H${row}="","",MAX($H$5:H${row}))`]];
  live.getRange(`J${row}`).formulas = [[`=IF(H${row}="","",H${row}/I${row}-1)`]];
}
live.getRange("A5:A104").format.numberFormat = "yyyy-mm-dd";
live.getRange("B5:I104").format.numberFormat = "#,##0;[Red](#,##0)";
live.getRange("J5:J104").format.numberFormat = "0.00%;[Red](0.00%)";
setWidths(live, [12, 14, 14, 14, 12, 14, 14, 14, 14, 11, 18, 34, 22, 32]);
live.freezePanes.freezeRows(4);

// Quantified qualitative rules and working formulas.
title(definitions, "A1:N1", "判定ルール_v1（定性語の数値定義）");
definitions.getRange("A2:N2").merge();
definitions.getRange("A2").values = [["仕様v1のしきい値を固定。青字入力で判定式を動作確認できる。"]];
definitions.getRange("A4:N8").values = [
  ["用語", "入力1", "入力2", "入力3", "入力4", "入力5", "入力6", "入力7", "入力8", "判定", "閾値1", "閾値2", "閾値3", "用途"],
  ["急騰", "当日騰落率", "3日騰落率", "出来高/20日平均", null, null, null, null, null, null, 0.05, 0.12, 2, "買い本フラグにしない"],
  ["押し目反発", "25日線", "前日25日線", "終値/25日線-1", "高値からの調整率", "反発確認(1/0)", null, null, null, null, -0.03, -0.08, 1, "買い候補"],
  ["上抜け定着", "節目", "当日終値", "翌日終値", null, null, null, null, null, null, 2, null, null, "2営業日連続で維持"],
  ["長い上ヒゲ", "始値", "高値", "安値", "終値", null, null, null, null, null, 0.5, 1 / 3, 1 / 3, "利確・見送り優先"],
];
header(definitions.getRange("A4:N4"));
body(definitions.getRange("A5:N8"));
inputStyle(definitions.getRange("B5:I8"));
definitions.getRange("J5:J8").formulas = [
  ["=IF(COUNT(B5:D5)<3,\"\",IF(AND(OR(B5>=K5,C5>=L5),D5>=M5),\"該当\",\"非該当\"))"],
  ["=IF(COUNT(B6:F6)<5,\"\",IF(AND(B6>C6,D6>=K6,E6<=K6,E6>=L6,F6=M6),\"該当\",\"非該当\"))"],
  ["=IF(COUNT(B7:D7)<3,\"\",IF(AND(C7>B7,D7>=B7),\"該当\",\"非該当\"))"],
  ["=IF(COUNT(B8:E8)<4,\"\",IF(AND((C8-MAX(B8,E8))/(C8-D8)>=K8,ABS(E8-B8)/(C8-D8)<=L8,E8<=D8+(C8-D8)*M8),\"該当\",\"非該当\"))"],
];
definitions.getRange("B5:C5").format.numberFormat = "0.0%";
definitions.getRange("D5").format.numberFormat = "0.0x";
definitions.getRange("D6:E6").format.numberFormat = "0.0%";
definitions.getRange("K5:L5").format.numberFormat = "0.0%";
definitions.getRange("M5").format.numberFormat = "0.0x";
definitions.getRange("K6:L6").format.numberFormat = "0.0%";
definitions.getRange("K8:M8").format.numberFormat = "0.0%";
definitions.getRange("A10:N14").merge(true);
definitions.getRange("A10:A14").values = [
  ["押し目反発の反発確認=1：陽線で前日高値超え、または下ヒゲ陽線で前日安値を割らず引け。"],
  ["押し目反発は25日線が上向き、終値が25日線-3%以上、直近高値から-3〜-8%調整したものとする。"],
  ["急騰は（当日+5%以上 OR 3日合計+12%以上）AND 出来高20日平均比2倍以上。"],
  ["上抜け定着は当日終値が節目超え、翌営業日終値も節目以上。発注はその次の営業日始値。"],
  ["同日の利確/損切り両到達は、時系列不明なら損切り優先で記録。"],
];
definitions.getRange("A10:N14").format.wrapText = true;
setWidths(definitions, [17, 14, 14, 16, 16, 14, 13, 13, 13, 12, 11, 11, 11, 28]);

// Compact checks before export.
const errorScan = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "verification v1 formula error scan",
});
console.log(errorScan.ndjson);

const inspectMain = await workbook.inspect({ kind: "table", range: "ステップ2_ルール検証!A14:P20", include: "values,formulas", tableMaxRows: 12, tableMaxCols: 16 });
console.log(inspectMain.ndjson);
const inspectForward = await workbook.inspect({ kind: "table", range: "前進検証_v1!A31:P44", include: "values,formulas", tableMaxRows: 20, tableMaxCols: 16 });
console.log(inspectForward.ndjson);

for (const name of ["ステップ2_ルール検証", "前進検証_v1", "過去検証_v1", "実運用比較_v1", "判定ルール_v1"]) {
  const preview = await workbook.render({ sheetName: name, autoCrop: "all", scale: 1, format: "png" });
  await fs.writeFile(path.join(qaDir, `${name}.png`), new Uint8Array(await preview.arrayBuffer()));
}

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(tempPath);
await fs.rename(tempPath, inputPath);
console.log(`saved:${inputPath}`);
