# HANDOFF_TO_CODEX_20260616.md

# Codex引き継ぎ：TRADING_MAP.md への登録（前提日 2026-06-16）

## あなた（Codex）への依頼

Cowork（Claude）側はDriveコネクタで新規ファイルを作成できるが、**既存 .md の上書き編集ができない**（Drive UIでも .md はインライン編集不可）。
そこで、Mac上のあなた（Codex）に、既存ファイル `TRADING_MAP.md` への追記と、Gitへの反映をお願いしたい。

## 前提（重要）

- 正本は `TRADING_MAP.md`（状態・構成の唯一の正）。表を他ファイルに重複させない。
- このフォルダ（トレーディング）は AI-Hub とは独立運用。AI-Hub側のコンテキストは読み込まない。
- 売買の断定的推奨はしない。最終判断はユーザー。
- 機微情報（口座番号・資産額等）は書かない。
- Git操作はMac側（サンドボックス外）で行う。Drive同期下のため push 中は同期を一時停止すると `.git` 破損リスクを下げられる。

## 今セッション（2026-06-16）でCowork側が新規作成したファイル

1. `LEARNING_LOG.md` … セッションごとの学習・検証ログ（蓄積・AI引き継ぎ用）。本日分（候補7銘柄の直近値・監視ライン照合、損益シミュレーションA/B比較、学習方針）を出典つきで記録済み。
2. `EXTERNAL_MEMORY_NOTES.md` … AI-Hubの外部記憶化手法（入口固定／現在・履歴の二層分離／3層記憶／終了プロトコル／命名・上書き対応）を抽出し、トレーディングへの適用案を整理。

## やってほしいこと

### 1. `TRADING_MAP.md` の「## ファイル構成」表に2行追加

`| trading_policy_20260614.md | 資金管理・候補銘柄・損切利確の方針メモ | 🟢 最新 |` の直後に、以下を挿入：

```
| LEARNING_LOG.md | セッションごとの学習・検証ログ（蓄積・AI引き継ぎ用） | 🟢 新規 |
| EXTERNAL_MEMORY_NOTES.md | 外部記憶化の運用ノウハウ（AI-Hub手法の抽出・適用） | 🟢 新規 |
```

### 2. `TRADING_MAP.md` の「## 更新履歴」の先頭に2行追加

```
- 2026-06-16：EXTERNAL_MEMORY_NOTES.md 新規作成。AI-Hubの外部記憶化手法を抽出し適用案を整理。前提日 2026-06-16
- 2026-06-16：LEARNING_LOG.md 新規作成（学習・検証ログの蓄積開始）。データ更新（候補7銘柄の直近値・監視ライン照合）・損益シミュレーション（建て方A/B比較）・学習方針を記録。前提日 2026-06-16
```

### 3. （任意・推奨）入口の読む順序を明記

`EXTERNAL_MEMORY_NOTES.md` の適用案に沿い、`TRADING_MAP.md` 冒頭の「## このファイルの役割」付近に一文追加してよい：
「作業AIはまず TRADING_MAP.md を読み、次に LEARNING_LOG.md（履歴）と EXTERNAL_MEMORY_NOTES.md（運用ノウハウ）を読む。」

### 4. 「最終更新」日付の確認

`TRADING_MAP.md` 冒頭の「最終更新」は 2026-06-16 のままでよい（同日内のため）。

### 5. Gitへ反映

```
git add TRADING_MAP.md LEARNING_LOG.md EXTERNAL_MEMORY_NOTES.md HANDOFF_TO_CODEX_20260616.md
git commit -m "Add learning log & external-memory notes; register in TRADING_MAP (前提日 2026-06-16)"
git push
```

## 完了後の確認

- `TRADING_MAP.md` のファイル構成表に2ファイルが載っているか。
- 更新履歴に2行が入っているか。
- `git status` がクリーンで、リモート `main` と一致しているか。

## 補足（Cowork側の制約メモ・恒久ルール候補）

- 新規ファイル作成＝Driveコネクタ（Cowork）で可。
- 既存 .md の編集・改名・削除、Git操作＝Mac側（Codex／Claude Code）で実施。
- 詳細な外部記憶化の運用方針は `EXTERNAL_MEMORY_NOTES.md` を参照。
