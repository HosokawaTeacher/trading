# Claude Code 引き継ぎプロンプト（GitHub連携の仕上げ）

作成：2026-06-16 ／ 依頼者：細川昭彦

トレーディングフォルダで `claude` を起動し、下の「依頼文」をそのまま貼って実行してください。
Claude Codeはサンドボックス外でMacのファイルに直接アクセスできるため、保護された `.git` も削除できます。

---

## 依頼文（このまま貼る）

このフォルダ（Googleドライブ同期下のトレーディング検討ワークスペース）を、GitHubの既存Privateリポジトリ `https://github.com/HosokawaTeacher/trading.git` に正式に紐づいた稼働リポジトリへ仕上げてください。

背景：初回pushは別環境のサンドボックス制約を回避して「.gitを除いた一時コピー」から行ったため、この作業フォルダ自体はまだリモートに紐づいていません。フォルダ内には壊れた作りかけの `.git` が残っています（過去の環境では削除が `Operation not permitted` で拒否されたが、Claude Codeはサンドボックス外なので削除できるはず）。今回これをきれいにやり直し、以後はこのフォルダから通常のcommit/pushで運用できる状態にしてください。

実行内容：

1. 壊れた `.git` を完全に削除する（`rm -rf .git`）。消せない場合は原因を調べて対処する。
2. `git init -b main`、`git config user.name "HosokawaTeacher"`、`git config user.email "hosokawa.edu@gmail.com"`。
3. `git remote add origin https://github.com/HosokawaTeacher/trading.git`。
4. 既存リモート履歴を取り込む：`git fetch origin` → `git reset --soft origin/main`（リモートの `main` は `a41c30cb87e1fd658a42f0bef29ec1843b3044d7`）。失敗時はその行を飛ばす。
5. `git add -A` の前に `git status` で追跡対象を確認。`.gitignore` 整備済み（.DS_Store / *.tmp / *.bak / ~$* / node_modules 等を除外）。`node_modules` や一時ファイルが入っていないこと、`CODEX_HANDOFF.md` / `CLAUDE_CODE_HANDOFF.md` / `TRADING_MAP.md` の変更が含まれることを確認。
6. `git commit -m "Sync working folder to repo (前提日 2026-06-16)"`（差分が無ければスキップ）。
7. `git push -u origin main`。
8. 仕上げ確認：`git log --oneline -3` と `git ls-remote origin main` が同じコミットを指すこと。

注意：
- 機微情報（口座番号・資産額・パスワード等）はコミットに含めない。
- Googleドライブ同期下のため、可能ならpush中は同期を一時停止すると `.git` 破損リスクを下げられる。
- 完了後、最新コミットのハッシュを報告してください（`TRADING_MAP.md` の更新履歴に反映します）。
