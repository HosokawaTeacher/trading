# Codex 引き継ぎプロンプト（GitHub連携の仕上げ）

作成：2026-06-16 ／ 依頼者：細川昭彦

そのまま下の「依頼文」をCodexに貼って実行してください。

---

## 依頼文（このまま貼る）

トレーディングフォルダを、GitHubの既存Privateリポジトリ `https://github.com/HosokawaTeacher/trading.git` に正式に紐づいた稼働リポジトリへ仕上げてください。

背景：初回pushはCowork経由の制約回避で「.gitを除いた一時コピー」から行ったため、作業フォルダ自体はまだリモートに紐づいていません（壊れた作りかけの `.git` が残っている可能性あり）。今回これをきれいにやり直し、以後はこのフォルダから通常のcommit/pushで運用できる状態にしてください。

手順：

```bash
# 1. フォルダへ移動（パスは環境に合わせて読み替え可）
cd "/Users/akihiko/Library/CloudStorage/GoogleDrive-hosokawa.edu@gmail.com/マイドライブ/トレーディング"

# 2. 作りかけ/壊れた .git を完全削除
rm -rf .git

# 3. 初期化と識別情報
git init -b main
git config user.name "HosokawaTeacher"
git config user.email "hosokawa.edu@gmail.com"

# 4. リモート登録
git remote add origin https://github.com/HosokawaTeacher/trading.git

# 5. 既存リモート(main)を取り込む（初回pushの履歴があるため）
git fetch origin
git reset --soft origin/main   # 既存履歴の上に乗せる。失敗時はこの行を飛ばす

# 6. 現在のフォルダ内容をコミット
git add -A
git status            # 追跡対象を確認（node_modules等が入っていないこと）
git commit -m "Sync working folder to repo (前提日 2026-06-16)"   # 差分が無ければスキップ可

# 7. push
git push -u origin main
```

確認事項：
- `.gitignore` は整備済み（.DS_Store / *.tmp / *.bak / ~$* / node_modules 等を除外）。`git status` で `node_modules` や一時ファイルが追跡対象に入っていないことを確認。
- push後、`git log --oneline -3` とGitHub上の `main` が同じコミットを指しているか確認。
- 機微情報（口座番号・資産額・パスワード等）はコミットに含めない。
- Googleドライブ同期下のため、可能ならpush中は同期を一時停止すると `.git` 破損リスクを下げられる。

完了したら、最新コミットのハッシュを教えてください（TRADING_MAP.md の更新履歴に反映します）。
