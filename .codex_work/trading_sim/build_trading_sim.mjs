import fs from "node:fs/promises";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const outputDir = "/Users/hosokawa/Library/CloudStorage/GoogleDrive-hosokawa.edu@gmail.com/マイドライブ/トレーディング/outputs/trading_sim_20260614";
const outputPath = `${outputDir}/5_model_trading_simulation_1month.xlsx`;

const wb = Workbook.create();
const dashboard = wb.worksheets.add("Dashboard");
const models = wb.worksheets.add("5 Models");
const capital = wb.worksheets.add("Capital Plan");
const watchlist = wb.worksheets.add("Watchlist");
const log = wb.worksheets.add("Trade Log");
const weekly = wb.worksheets.add("Weekly Review");
const candidates = wb.worksheets.add("Candidate Selection");
const searchTemplate = wb.worksheets.add("Search Template");
const prompts = wb.worksheets.add("AI Prompts");
const rules = wb.worksheets.add("Rules");
const checks = wb.worksheets.add("Checks");

const theme = {
  navy: "1f3a5f",
  teal: "0f766e",
  green: "2f7d32",
  amber: "f59e0b",
  red: "b91c1c",
  gray: "f3f4f6",
  paleBlue: "eaf2ff",
  paleGreen: "e9f7ef",
  paleAmber: "fff7e6",
  white: "ffffff",
  black: "111827",
};

function setWidths(sheet, widths) {
  const toCol = (n) => {
    let s = "";
    while (n > 0) {
      const m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - m - 1) / 26);
    }
    return s;
  };
  widths.forEach((w, i) => {
    const col = toCol(i + 1);
    sheet.getRange(`${col}1:${col}1`).format.columnWidth = w;
  });
}

function title(sheet, range, text) {
  const r = sheet.getRange(range);
  r.merge();
  r.values = [[text]];
  r.format.fill.color = theme.navy;
  r.format.font.color = theme.white;
  r.format.font.bold = true;
  r.format.font.size = 16;
  r.format.horizontalAlignment = "center";
  r.format.verticalAlignment = "middle";
}

function section(sheet, range, text, color = theme.teal) {
  const r = sheet.getRange(range);
  r.merge();
  r.values = [[text]];
  r.format.fill.color = color;
  r.format.font.color = theme.white;
  r.format.font.bold = true;
}

function header(range) {
  range.format.fill.color = theme.paleBlue;
  range.format.font.bold = true;
  range.format.font.color = theme.black;
  range.format.borders.color = "cbd5e1";
}

function body(range) {
  range.format.borders.color = "e5e7eb";
  range.format.verticalAlignment = "top";
}

// Rules
title(rules, "A1:H1", "1ヶ月シミュレーション基本ルール");
rules.getRange("A3:B12").values = [
  ["目的", "5つの売買モデルを同じ条件で観察し、1ヶ月後に残すモデルを選ぶ"],
  ["期間", "2026-06-14開始。4週間の模擬運用を想定"],
  ["実売買", "原則なし。記録は仮想エントリー／仮想決済として扱う"],
  ["推奨対象", "まずは流動性の高い日本株・米国株・ETFから選ぶ"],
  ["1回の最大損失", "総資金の0.5%〜1.0%を上限目安"],
  ["同時観察数", "各モデル3〜5候補まで"],
  ["禁止", "損切りなし、ナンピン前提、SNS人気だけの採用、AI丸投げ"],
  ["週末作業", "候補更新、ルール逸脱確認、翌週の見送り条件を決める"],
  ["平日作業", "価格・ニュース・損切り条件だけ確認"],
  ["判断基準", "収益より、損失管理・再現性・自分が運用できるかを重視"],
];
header(rules.getRange("A3:B3"));
body(rules.getRange("A3:B12"));
setWidths(rules, [22, 95, 14, 14, 14, 14, 14, 14]);

// Models
title(models, "A1:J1", "採用する5つのモデル");
models.getRange("A3:J3").values = [[
  "Model ID", "モデル名", "狙い", "週末チェック項目", "エントリー条件", "損切り方針", "利確方針", "AIの使い方", "主なリスク", "採用判定基準"
]];
const modelRows = [
  ["M1", "トレンドフォロー", "上昇トレンドが続く銘柄を狙う", "指数、セクター、移動平均、直近高値", "上昇基調で押し目または高値更新後の定着", "直近安値・移動平均割れ・想定損失上限", "分割利確またはトレーリング", "チャート条件の言語化、候補比較", "急反転、過熱、遅いエントリー", "損小利大が守れる"],
  ["M2", "決算モメンタム", "好決算・上方修正後の継続上昇を狙う", "決算、上方修正、出来高、翌週イベント", "好材料後に高値圏で崩れず推移", "決算窓埋め・材料否定・損失上限", "材料の織り込み感が強い時に縮小", "決算短信・説明資料・ニュース要約", "材料出尽くし、ギャップダウン", "材料と価格が同方向に続く"],
  ["M3", "ETFコア＋個別株サテライト", "市場全体の流れを取りつつ個別で上乗せ", "指数ETF、金利、為替、セクターETF", "地合い良好時にETF中心、個別は少額", "ETFは広め、個別は厳しめ", "ETFは段階利確、個別は短め", "マクロ要約、ETF比較、個別候補整理", "個別の寄与が薄い、相場急変", "安定して継続しやすい"],
  ["M4", "ニュース確認型モメンタム", "価格上昇をニュースの質で確認する", "直近ニュース、テーマ、競合、規制", "価格が強く、ニュースが一過性でない", "ニュース否定・出来高減少・損失上限", "ニュース熱が冷めたら縮小", "ニュースを強気/弱気/中立で採点", "誤情報、煽り、AIの読み違い", "ニュース確認で悪い取引を避ける"],
  ["M5", "リスク最小・高流動性リバランス", "無理に当てず、流動性と損失制御を重視", "出来高、ボラ、スプレッド、相関", "低ボラ・高流動性で条件が明確", "ATRや直近安値を使い小さく管理", "小さく確定、悪化時は撤退", "ポジションサイズ計算、相関チェック", "リターンが小さい、機会損失", "損失が最も安定して小さい"],
];
models.getRange("A4:J8").values = modelRows;
header(models.getRange("A3:J3"));
body(models.getRange("A4:J8"));
setWidths(models, [10, 24, 35, 35, 38, 34, 34, 38, 30, 28]);

// Capital Plan
title(capital, "A1:J1", "100万円前提 資金配分案");
capital.getRange("A3:B8").values = [
  ["総資金", 1000000],
  ["初回投入目安", 583200],
  ["現金温存目安", 416800],
  ["初回投入比率", null],
  ["現金比率", null],
  ["方針", "上限一杯に投資せず、良さそうな銘柄や押し目が出たときに回せる余剰金を確保する"],
];
capital.getRange("B6").formulas = [["=B4/B3"]];
capital.getRange("B7").formulas = [["=B5/B3"]];
capital.getRange("A3:B8").format.borders.color = "e5e7eb";
capital.getRange("B3:B5").numberFormat = "#,##0";
capital.getRange("B6:B7").numberFormat = "0.0%";
capital.getRange("A10:K10").values = [[
  "区分", "コード", "銘柄", "株数", "Entry目安", "投資額", "Stop", "第1利確", "第2利確", "損切り損益", "第1利確損益"
]];
capital.getRange("A11:K13").values = [
  ["初回候補", "8306", "三菱UFJ FG", 100, 3162, null, 3000, 3300, 3450, null, null],
  ["初回候補", "9433", "KDDI", 100, 2670, null, 2580, 2800, 2900, null, null],
  ["押し目待ち", "8035", "東京エレクトロン", 2, 66000, null, 62000, 74000, 78000, null, null],
];
capital.getRange("L10:M10").values = [["第2利確損益", "状態"]];
capital.getRange("F11:F13").formulas = [["=D11*E11"], ["=D12*E12"], ["=D13*E13"]];
capital.getRange("J11:J13").formulas = [["=(G11-E11)*D11"], ["=(G12-E12)*D12"], ["=(G13-E13)*D13"]];
capital.getRange("K11:K13").formulas = [["=(H11-E11)*D11"], ["=(H12-E12)*D12"], ["=(H13-E13)*D13"]];
capital.getRange("L11:L13").formulas = [["=(I11-E11)*D11"], ["=(I12-E12)*D12"], ["=(I13-E13)*D13"]];
capital.getRange("M11:M13").values = [["実行候補"], ["実行候補"], ["待機"]];
capital.getRange("A16:B20").values = [
  ["初回候補のみ投資額", null],
  ["初回候補のみ現金", null],
  ["初回候補損切り", null],
  ["初回候補第1利確", null],
  ["初回候補第2利確", null],
];
capital.getRange("B16:B20").formulas = [
  ["=SUM(F11:F12)"],
  ["=B3-B16"],
  ["=SUM(J11:J12)"],
  ["=SUM(K11:K12)"],
  ["=SUM(L11:L12)"],
];
capital.getRange("D16:E20").values = [
  ["8035まで入れた投資額", null],
  ["8035まで入れた現金", null],
  ["8035込み損切り", null],
  ["8035込み第1利確", null],
  ["8035込み第2利確", null],
];
capital.getRange("E16:E20").formulas = [
  ["=SUM(F11:F13)"],
  ["=B3-E16"],
  ["=SUM(J11:J13)"],
  ["=SUM(K11:K13)"],
  ["=SUM(L11:L13)"],
];
header(capital.getRange("A10:M10"));
body(capital.getRange("A11:M13"));
body(capital.getRange("A16:E20"));
capital.getRange("E11:I13").numberFormat = "#,##0";
capital.getRange("F11:L13").numberFormat = "#,##0;[Red](#,##0);-";
capital.getRange("B16:B20").numberFormat = "#,##0;[Red](#,##0);-";
capital.getRange("E16:E20").numberFormat = "#,##0;[Red](#,##0);-";
setWidths(capital, [14, 10, 22, 10, 12, 13, 12, 12, 12, 13, 13, 13, 12]);

// Watchlist
title(watchlist, "A1:K1", "監視リストと条件");
watchlist.getRange("A3:K3").values = [[
  "優先", "コード", "銘柄", "扱い", "上昇確認", "見送り/損切り", "資金方針", "確認するデータ", "検索時の判断", "次アクション", "メモ"
]];
watchlist.getRange("A4:K10").values = [
  [1, "8306", "三菱UFJ FG", "初回候補", "3,220円超え、銀行セクター強含み", "3,000円割れ", "100株まで", "価格、出来高、金利/日銀ニュース", "継続候補か利確/撤退を判定", "検索時に最新価格で再計算", "主力候補"],
  [2, "9433", "KDDI", "初回候補", "2,750円超え、2,650円台維持", "2,580円割れ", "100株まで", "価格、出来高、通信セクター", "守り枠として継続可否を判定", "検索時に最新価格で再計算", "守り候補"],
  [3, "8035", "東京エレクトロン", "押し目待ち", "65,000〜66,000円台から反発、または70,000円超え", "65,000円をすぐ割る、半導体全体反落", "S株1〜2株のみ", "価格、出来高、半導体指数、AI関連ニュース", "追わずに押し目を待てるか判定", "条件到達時だけ検討", "急騰後なので注意"],
  [4, "6857", "アドバンテスト", "保留", "28,500円超え、出来高増", "26,000円を保てない", "S株少額のみ", "価格、出来高、半導体指数", "反発継続か荒さ優先で見送りか", "週末再評価", "値動き荒い"],
  [5, "9984", "ソフトバンクG", "保留", "高値からの調整後に反発確認", "資金調達/AI関連悪材料で続落", "小さく観察のみ", "価格、ニュース、AI/Arm/OpenAI関連", "ニュースと価格反応の一致を確認", "週末再評価", "ニュース感応度高い"],
  [6, "7011", "三菱重工業", "見送り寄り", "下降基調から反転し高値更新", "下降継続", "当面なし", "価格、出来高、防衛/受注ニュース", "材料とチャートが一致するか確認", "週末再評価", "材料は強いが形が弱い"],
  [7, "6501", "日立製作所", "見送り寄り", "下げ止まり後に移動平均回復", "下落継続", "当面なし", "価格、出来高、決算/インフラ材料", "下げ止まり確認が先", "週末再評価", "今は様子見"],
];
header(watchlist.getRange("A3:K3"));
body(watchlist.getRange("A4:K10"));
setWidths(watchlist, [8, 10, 22, 14, 38, 34, 18, 34, 36, 26, 24]);

// Candidate Selection
title(candidates, "A1:L1", "モデル別 銘柄選定（模擬観察候補）");
candidates.getRange("A3:L3").values = [[
  "Model ID", "モデル名", "コード", "銘柄/ETF", "市場", "役割", "選定理由", "観察条件", "見送り条件", "確認ソース", "メモ", "優先度"
]];
const candidateRows = [
  ["M5", "リスク最小・高流動性", "8306", "三菱UFJ FG", "日本株", "主力候補", "3ヶ月日足の上昇基調が比較的きれいで、100万円口座でも100株単位で扱いやすい", "3,220円超え、銀行セクター強含み", "3,000円割れ、金利低下や銀行セクター悪化", "SBIチャート/金利ニュース", "初回100株候補。価格更新時に損益再計算", "高"],
  ["M5", "リスク最小・高流動性", "9433", "KDDI", "日本株", "守り候補", "通信大型株で値動きが比較的落ち着きやすく、余剰資金を残す方針と相性がよい", "2,750円超え、2,650円台維持", "2,580円割れ、ディフェンシブ株まで売られる地合い", "SBIチャート/通信セクター", "初回100株候補。上昇力より安定性重視", "高"],
  ["M1", "トレンドフォロー", "8035", "東京エレクトロン", "日本株", "押し目待ち", "候補内で最も強いチャートだが急騰直後。単元では買えないためS株で少額観察", "65,000〜66,000円台から反発、または70,000円超え出来高あり", "65,000円をすぐ割る、半導体全体反落", "SBIチャート/半導体ニュース", "S株1〜2株のみ検討。飛び乗り禁止", "中"],
  ["M1", "トレンドフォロー", "6857", "アドバンテスト", "日本株", "保留", "反発力はあるが値動きが荒く、100万円口座ではS株前提", "28,500円超え、出来高増", "26,000円を保てない、半導体全体反落", "SBIチャート/半導体ニュース", "今週は追わず週末再評価", "中"],
  ["M4", "ニュース確認型モメンタム", "9984", "ソフトバンクグループ", "日本株", "保留", "AIニュース感応度は高いが、高値からの調整中で振られやすい", "調整後の反発確認、AI関連ニュースが価格を支える", "資金調達/AI関連悪材料で続落", "SBIチャート/AIニュース", "ニュースと価格反応が一致するまで待つ", "中"],
  ["M2", "決算モメンタム", "7011", "三菱重工業", "日本株", "見送り寄り", "受注や防衛材料は強いが、直近チャートは下降基調で反転確認が弱い", "下降基調から反転し高値更新", "下降継続、材料に株価が反応しない", "SBIチャート/FT等ニュース", "材料は継続監視。今は買い候補から外す", "低"],
  ["M2", "決算モメンタム", "6501", "日立製作所", "日本株", "見送り寄り", "AI・インフラ材料はあるが、直近は下落が目立ち下げ止まり確認が先", "移動平均回復、下げ止まり確認", "下落継続、出来高を伴う下抜け", "SBIチャート/決算ニュース", "週末再評価", "低"],
  ["M3", "ETFコア＋個別株サテライト", "1306", "TOPIX連動型ETF", "日本ETF", "地合い確認", "個別株リスクを抑えて日本株全体の地合いを見る基準", "TOPIXが上昇基調、下値切り上げ", "指数が週足で崩れる", "指数/ETFチャート", "今回は実行候補ではなく比較用", "中"],
  ["M3", "ETFコア＋個別株サテライト", "1321", "日経225連動型ETF", "日本ETF", "地合い確認", "大型・値がさ株主導の相場を確認する基準", "日経平均が高値圏維持", "値がさ半導体の急落が続く", "指数/ETFチャート", "今回は実行候補ではなく比較用", "中"],
];
candidates.getRange("A4:L12").values = candidateRows;
header(candidates.getRange("A3:L3"));
body(candidates.getRange("A4:L12"));
setWidths(candidates, [10, 24, 10, 22, 12, 18, 44, 38, 36, 28, 32, 10]);

// Trade Log
title(log, "A1:T1", "模擬取引ログ");
log.getRange("A3:T3").values = [[
  "No.", "週", "Model ID", "銘柄/ティッカー", "市場", "候補日", "仮想Entry日", "Entry価格", "Stop価格", "Target価格",
  "想定損失率", "仮想Exit日", "Exit価格", "損益率", "状態", "根拠", "AI要約", "ルール逸脱", "週末メモ", "次アクション"
]];
header(log.getRange("A3:T3"));
body(log.getRange("A4:T103"));
log.getRange("A4:A103").formulas = Array.from({ length: 100 }, (_, i) => [`=IF(D${i + 4}<>"",ROW()-3,"")`]);
const logRows = candidateRows.slice(0, 3);
log.getRange("B4:E6").values = logRows.map(r => ["Week 1", r[0], `${r[2]} ${r[3]}`, r[4]]);
log.getRange("H4:J6").values = [
  [3162, 3000, 3300],
  [2670, 2580, 2800],
  [66000, 62000, 74000],
];
log.getRange("O4:O6").values = [["候補"], ["候補"], ["候補"]];
log.getRange("P4:P6").values = logRows.map(r => [r[6]]);
log.getRange("Q4:Q6").values = logRows.map(r => [r[7]]);
log.getRange("R4:R6").values = logRows.map(() => ["なし"]);
log.getRange("S4:S6").values = logRows.map(r => [r[10]]);
log.getRange("T4:T6").values = [
  ["最新検索時に価格を更新し、100株で再計算"],
  ["最新検索時に価格を更新し、100株で再計算"],
  ["押し目条件到達時のみS株1〜2株を検討"],
];
log.getRange("K4:K103").formulas = Array.from({ length: 100 }, (_, i) => [`=IF(AND(H${i + 4}<>"",I${i + 4}<>""),(I${i + 4}-H${i + 4})/H${i + 4},"")`]);
log.getRange("N4:N103").formulas = Array.from({ length: 100 }, (_, i) => [`=IF(AND(H${i + 4}<>"",M${i + 4}<>""),(M${i + 4}-H${i + 4})/H${i + 4},"")`]);
log.getRange("B4:B103").dataValidation = { type: "list", formula1: '"Week 1,Week 2,Week 3,Week 4"' };
log.getRange("C4:C103").dataValidation = { type: "list", formula1: '"M1,M2,M3,M4,M5"' };
log.getRange("O4:O103").dataValidation = { type: "list", formula1: '"候補,保有中,利確,損切り,見送り,取消"' };
log.getRange("R4:R103").dataValidation = { type: "list", formula1: '"なし,軽微,重大"' };
log.getRange("F4:G103").numberFormat = "yyyy-mm-dd";
log.getRange("L4:L103").numberFormat = "yyyy-mm-dd";
log.getRange("H4:J103").numberFormat = "#,##0.00";
log.getRange("M4:M103").numberFormat = "#,##0.00";
log.getRange("K4:K103").numberFormat = "0.0%;[Red](0.0%);-";
log.getRange("N4:N103").numberFormat = "0.0%;[Red](0.0%);-";
setWidths(log, [7, 11, 10, 18, 12, 13, 13, 12, 12, 12, 12, 13, 12, 11, 12, 35, 45, 12, 35, 24]);

// Weekly Review
title(weekly, "A1:N1", "週末レビュー");
weekly.getRange("A3:N3").values = [[
  "週", "レビュー日", "市場環境", "強いセクター/テーマ", "弱いセクター/テーマ", "新規候補数", "実行候補数", "見送り数",
  "最大懸念", "翌週の見送り条件", "残すモデル", "弱めるモデル", "学び", "次週アクション"
]];
weekly.getRange("A4:N7").values = [
  ["Week 1", null, "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Week 2", null, "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Week 3", null, "", "", "", "", "", "", "", "", "", "", "", ""],
  ["Week 4", null, "", "", "", "", "", "", "", "", "", "", "", ""],
];
header(weekly.getRange("A3:N3"));
body(weekly.getRange("A4:N7"));
weekly.getRange("B4:B7").numberFormat = "yyyy-mm-dd";
setWidths(weekly, [11, 14, 32, 28, 28, 12, 12, 12, 32, 34, 18, 18, 42, 32]);

// Search Template
title(searchTemplate, "A1:H1", "「検索して」依頼時の確認テンプレ");
searchTemplate.getRange("A3:B10").values = [
  ["目的", "このチャット上で最新データ確認からシミュレーションまで完結する"],
  ["確認タイミング", "ユーザーが「検索して」と依頼した時"],
  ["資金前提", "100万円。初回は上限一杯に使わず、余剰金を残す"],
  ["現在の初回候補", "8306 三菱UFJ FG、9433 KDDI"],
  ["押し目待ち", "8035 東京エレクトロン"],
  ["保留/再評価", "6857、9984、7011、6501、1306、1321"],
  ["判断原則", "買い推奨ではなく、模擬運用の候補更新として扱う"],
  ["出力", "最新データ、判断、資金配分、損切り/利確、残す現金"],
];
searchTemplate.getRange("A12:H12").values = [[
  "コード", "銘柄", "最新価格", "前日比", "出来高/売買代金", "ニュース/材料", "チャート判断", "今回の判断"
]];
searchTemplate.getRange("A13:H21").values = [
  ["8306", "三菱UFJ FG", "", "", "", "", "3,220円超えなら強気、3,000円割れなら撤退/見送り", ""],
  ["9433", "KDDI", "", "", "", "", "2,750円超えなら強気、2,580円割れなら撤退/見送り", ""],
  ["8035", "東京エレクトロン", "", "", "", "", "65,000〜66,000円台の反発、または70,000円超えを確認", ""],
  ["6857", "アドバンテスト", "", "", "", "", "28,500円超えなら再評価、26,000円割れは見送り", ""],
  ["9984", "ソフトバンクG", "", "", "", "", "調整後の反発とAIニュースの一致を確認", ""],
  ["7011", "三菱重工業", "", "", "", "", "下降基調からの反転確認が先", ""],
  ["6501", "日立製作所", "", "", "", "", "下げ止まりと移動平均回復が先", ""],
  ["1306", "TOPIX ETF", "", "", "", "", "日本株全体の地合い確認", ""],
  ["1321", "日経225 ETF", "", "", "", "", "大型株/半導体主導の地合い確認", ""],
];
searchTemplate.getRange("A24:H24").values = [[
  "シナリオ", "投資額", "現金", "損切り損益", "第1利確損益", "第2利確損益", "採用/見送り", "理由"
]];
searchTemplate.getRange("A25:H27").values = [
  ["8306+9433のみ", null, null, null, null, null, "基本案", "約40万円を残せる"],
  ["8306+9433+8035", null, null, null, null, null, "条件付き", "8035が押し目/再上昇確認できた場合"],
  ["全額投入", "", "", "", "", "", "原則見送り", "余剰資金を残す方針に反する"],
];
searchTemplate.getRange("B25:F26").formulas = [
  ["=SUM('Capital Plan'!F11:F12)", "='Capital Plan'!B3-B25", "=SUM('Capital Plan'!J11:J12)", "=SUM('Capital Plan'!K11:K12)", "=SUM('Capital Plan'!L11:L12)"],
  ["=SUM('Capital Plan'!F11:F13)", "='Capital Plan'!B3-B26", "=SUM('Capital Plan'!J11:J13)", "=SUM('Capital Plan'!K11:K13)", "=SUM('Capital Plan'!L11:L13)"],
];
searchTemplate.getRange("A30:B37").values = [
  ["返答テンプレ", "今日の検索結果では、実行候補/見送り候補を以下の通り更新します。"],
  ["1", "最新価格とチャート形状"],
  ["2", "ニュース/決算/セクター材料"],
  ["3", "100万円前提の投資額と残す現金"],
  ["4", "損切り/利確シナリオ"],
  ["5", "買わない条件"],
  ["6", "次回検索までの監視価格"],
  ["注意", "このチャット内で完結し、証券会社アプリは補助として扱う"],
];
header(searchTemplate.getRange("A12:H12"));
header(searchTemplate.getRange("A24:H24"));
body(searchTemplate.getRange("A3:B10"));
body(searchTemplate.getRange("A13:H21"));
body(searchTemplate.getRange("A25:H27"));
body(searchTemplate.getRange("A30:B37"));
searchTemplate.getRange("B25:F27").numberFormat = "#,##0;[Red](#,##0);-";
setWidths(searchTemplate, [14, 24, 12, 12, 18, 34, 48, 34]);

// AI Prompts
title(prompts, "A1:D1", "AIプロンプト集");
prompts.getRange("A3:D3").values = [["用途", "使うタイミング", "プロンプト", "出力形式"]];
const promptRows = [
  ["市場環境チェック", "週末", "今週の株式市場を、指数、金利、為替、セクター、リスクイベントに分けて整理してください。売買推奨ではなく、来週のスイング候補を探す前提条件として、強い領域・弱い領域・避けるべき条件を表で出してください。", "強い/弱い/注意/来週の見送り条件"],
  ["銘柄スクリーニング", "週末", "以下の候補銘柄を、トレンド、出来高、材料、決算、流動性、損切りの置きやすさで採点してください。買い推奨ではなく、模擬運用候補として順位付けし、除外理由を重視してください。", "銘柄別スコア、除外理由、確認すべきデータ"],
  ["決算・ニュース要約", "候補選定時", "この決算資料またはニュースを、株価に影響し得るポジティブ要因、ネガティブ要因、一過性要因、未確認リスクに分けて要約してください。最後に、価格モメンタムを支持するかを0〜5点で評価してください。", "要因別要約、0〜5点、注意点"],
  ["売買計画レビュー", "仮想エントリー前", "この模擬取引計画をレビューしてください。エントリー理由、損切り、利確、ポジションサイズ、見送り条件の曖昧さを指摘し、ルール違反があれば明示してください。", "OK/要修正、修正点、見送り条件"],
  ["週末振り返り", "毎週末", "今週の模擬取引ログをもとに、モデル別の良かった点、悪かった点、ルール逸脱、翌週に残すべきモデルを整理してください。損益よりも再現性と損失管理を重視してください。", "モデル別評価、改善点、翌週の方針"],
];
prompts.getRange("A4:D8").values = promptRows;
header(prompts.getRange("A3:D3"));
body(prompts.getRange("A4:D8"));
setWidths(prompts, [22, 18, 95, 38]);

// Dashboard formulas and chart helper
title(dashboard, "A1:L1", "5モデル 1ヶ月シミュレーション Dashboard");
dashboard.getRange("A3:B8").values = [
  ["開始日", new Date("2026-06-14")],
  ["対象期間", "4週間"],
  ["実売買", "なし（模擬運用）"],
  ["資金前提", "100万円。初回投入は約58万円、約42万円を温存"],
  ["最重要ルール", "上限一杯に投資せず、損切りとポジションサイズを先に決める"],
  ["次の作業", "検索依頼時に最新データを取り、このチャット上で再シミュレーション"],
];
dashboard.getRange("B3").numberFormat = "yyyy-mm-dd";
section(dashboard, "D3:L3", "モデル別サマリー");
dashboard.getRange("D4:L4").values = [["Model ID", "モデル名", "候補数", "完了数", "平均損益率", "勝率", "ルール逸脱", "状態", "コメント"]];
dashboard.getRange("D5:E9").values = modelRows.map(r => [r[0], r[1]]);
dashboard.getRange("F5:F9").formulas = [["=COUNTIF('Trade Log'!C:C,D5)"],["=COUNTIF('Trade Log'!C:C,D6)"],["=COUNTIF('Trade Log'!C:C,D7)"],["=COUNTIF('Trade Log'!C:C,D8)"],["=COUNTIF('Trade Log'!C:C,D9)"]];
dashboard.getRange("G5:G9").formulas = [["=COUNTIFS('Trade Log'!C:C,D5,'Trade Log'!M:M,\"<>\")"],["=COUNTIFS('Trade Log'!C:C,D6,'Trade Log'!M:M,\"<>\")"],["=COUNTIFS('Trade Log'!C:C,D7,'Trade Log'!M:M,\"<>\")"],["=COUNTIFS('Trade Log'!C:C,D8,'Trade Log'!M:M,\"<>\")"],["=COUNTIFS('Trade Log'!C:C,D9,'Trade Log'!M:M,\"<>\")"]];
dashboard.getRange("H5:H9").formulas = [["=IFERROR(AVERAGEIFS('Trade Log'!N:N,'Trade Log'!C:C,D5),\"\")"],["=IFERROR(AVERAGEIFS('Trade Log'!N:N,'Trade Log'!C:C,D6),\"\")"],["=IFERROR(AVERAGEIFS('Trade Log'!N:N,'Trade Log'!C:C,D7),\"\")"],["=IFERROR(AVERAGEIFS('Trade Log'!N:N,'Trade Log'!C:C,D8),\"\")"],["=IFERROR(AVERAGEIFS('Trade Log'!N:N,'Trade Log'!C:C,D9),\"\")"]];
dashboard.getRange("I5:I9").formulas = [["=IF(G5=0,\"\",COUNTIFS('Trade Log'!C:C,D5,'Trade Log'!N:N,\">0\")/G5)"],["=IF(G6=0,\"\",COUNTIFS('Trade Log'!C:C,D6,'Trade Log'!N:N,\">0\")/G6)"],["=IF(G7=0,\"\",COUNTIFS('Trade Log'!C:C,D7,'Trade Log'!N:N,\">0\")/G7)"],["=IF(G8=0,\"\",COUNTIFS('Trade Log'!C:C,D8,'Trade Log'!N:N,\">0\")/G8)"],["=IF(G9=0,\"\",COUNTIFS('Trade Log'!C:C,D9,'Trade Log'!N:N,\">0\")/G9)"]];
dashboard.getRange("J5:J9").formulas = [["=COUNTIFS('Trade Log'!C:C,D5,'Trade Log'!R:R,\"<>なし\",'Trade Log'!R:R,\"<>\")"],["=COUNTIFS('Trade Log'!C:C,D6,'Trade Log'!R:R,\"<>なし\",'Trade Log'!R:R,\"<>\")"],["=COUNTIFS('Trade Log'!C:C,D7,'Trade Log'!R:R,\"<>なし\",'Trade Log'!R:R,\"<>\")"],["=COUNTIFS('Trade Log'!C:C,D8,'Trade Log'!R:R,\"<>なし\",'Trade Log'!R:R,\"<>\")"],["=COUNTIFS('Trade Log'!C:C,D9,'Trade Log'!R:R,\"<>なし\",'Trade Log'!R:R,\"<>\")"]];
dashboard.getRange("K5:K9").formulas = Array.from({ length: 5 }, (_, i) => [`=IF(F${i+5}=0,"未入力",IF(J${i+5}>0,"要確認","観察中"))`]);
dashboard.getRange("L5:L9").values = [[""],[""],[""],[""],[""]];
header(dashboard.getRange("D4:L4"));
body(dashboard.getRange("A3:B8"));
body(dashboard.getRange("D5:L9"));
dashboard.getRange("H5:I9").numberFormat = "0.0%;[Red](0.0%);-";
dashboard.getRange("B3:B8").format.fill.color = theme.paleAmber;
dashboard.getRange("N4:O9").values = [["Model", "平均損益率"], ["M1", null], ["M2", null], ["M3", null], ["M4", null], ["M5", null]];
dashboard.getRange("O5:O9").formulas = [["=H5"], ["=H6"], ["=H7"], ["=H8"], ["=H9"]];
dashboard.getRange("O5:O9").numberFormat = "0.0%;[Red](0.0%);-";
setWidths(dashboard, [18, 34, 4, 10, 24, 10, 10, 13, 10, 12, 12, 34, 3, 8, 12]);
const avgChart = dashboard.charts.add("ColumnClustered", dashboard.getRange("N4:O9"), "Auto");
avgChart.title.text = "モデル別 平均損益率";
avgChart.setPosition(dashboard.getRange("D11:L24"));
avgChart.width = 540;
avgChart.height = 250;
section(dashboard, "A11:C11", "現在の実行案", theme.amber);
dashboard.getRange("A12:C16").values = [
  ["初回候補", "8306 三菱UFJ FG", "100株"],
  ["初回候補", "9433 KDDI", "100株"],
  ["押し目待ち", "8035 東京エレクトロン", "S株1〜2株"],
  ["初回投入", "約583,000円", ""],
  ["温存資金", "約417,000円", ""],
];
body(dashboard.getRange("A12:C16"));

// Checks
title(checks, "A1:F1", "チェック");
checks.getRange("A3:F3").values = [["項目", "Actual", "Expected", "差分", "Status", "Notes"]];
checks.getRange("A4:F8").values = [
  ["Model IDが未入力の取引", "", 0, "", "", "Trade LogのC列"],
  ["Entry/Stop入力時の想定損失率", "", "算出", "", "", "K列が空なら未完成"],
  ["完了取引の損益率", "", "算出", "", "", "N列が空なら未完成"],
  ["重大なルール逸脱", "", 0, "", "", "R列が重大の件数"],
  ["レビュー週数", "", 4, "", "", "Weekly Reviewは4週分"],
];
checks.getRange("B4:B8").formulas = [
  ["=COUNTIFS('Trade Log'!D:D,\"<>\",'Trade Log'!C:C,\"\")"],
  ["=COUNTIFS('Trade Log'!H:H,\"<>\",'Trade Log'!I:I,\"<>\",'Trade Log'!K:K,\"\")"],
  ["=COUNTIFS('Trade Log'!H:H,\"<>\",'Trade Log'!M:M,\"<>\",'Trade Log'!N:N,\"\")"],
  ["=COUNTIF('Trade Log'!R:R,\"重大\")"],
  ["=COUNTA('Weekly Review'!A4:A7)"],
];
checks.getRange("D4:D8").formulas = [
  ["=B4-C4"],["=B5"],["=B6"],["=B7-C7"],["=B8-C8"],
];
checks.getRange("E4:E8").formulas = [
  ["=IF(D4=0,\"OK\",\"要確認\")"],["=IF(D5=0,\"OK\",\"要確認\")"],["=IF(D6=0,\"OK\",\"要確認\")"],["=IF(D7=0,\"OK\",\"要確認\")"],["=IF(D8=0,\"OK\",\"要確認\")"],
];
header(checks.getRange("A3:F3"));
body(checks.getRange("A4:F8"));
setWidths(checks, [28, 12, 12, 12, 12, 45]);

for (const sheet of [dashboard, models, capital, watchlist, log, weekly, candidates, searchTemplate, prompts, rules, checks]) {
  sheet.showGridlines = false;
  sheet.getUsedRange().format.font.name = "Aptos";
  sheet.getUsedRange().format.font.size = 10;
}

await fs.mkdir(outputDir, { recursive: true });

const inspectDashboard = await wb.inspect({
  kind: "table",
  range: "Dashboard!A1:L12",
  include: "values,formulas",
  tableMaxRows: 20,
  tableMaxCols: 12,
});
console.log(inspectDashboard.ndjson);

const errors = await wb.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 300 },
  summary: "final formula error scan",
});
console.log(errors.ndjson);

await wb.render({ sheetName: "Dashboard", range: "A1:L14", scale: 1 });
await wb.render({ sheetName: "5 Models", range: "A1:J10", scale: 1 });
await wb.render({ sheetName: "Capital Plan", range: "A1:M22", scale: 1 });
await wb.render({ sheetName: "Watchlist", range: "A1:K12", scale: 1 });
await wb.render({ sheetName: "Trade Log", range: "A1:T18", scale: 1 });
await wb.render({ sheetName: "Weekly Review", range: "A1:N9", scale: 1 });
await wb.render({ sheetName: "Candidate Selection", range: "A1:L20", scale: 1 });
await wb.render({ sheetName: "Search Template", range: "A1:H28", scale: 1 });
await wb.render({ sheetName: "AI Prompts", range: "A1:D10", scale: 1 });
await wb.render({ sheetName: "Rules", range: "A1:B14", scale: 1 });
await wb.render({ sheetName: "Checks", range: "A1:F10", scale: 1 });

const out = await SpreadsheetFile.exportXlsx(wb);
await out.save(outputPath);
console.log(outputPath);
