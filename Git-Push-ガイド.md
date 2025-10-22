# Git Push ガイド - ItemPicking プロジェクト

## 📋 現在の状態

✅ Gitリポジトリ初期化完了  
✅ 23ファイルをコミット完了  
✅ ローカルリポジトリ準備完了  

**コミットID:** feec0d4  
**ブランチ:** master  
**ファイル数:** 23個 (361.7 KB)  

---

## 🚀 リモートリポジトリへのPush方法

### オプション1: 新規GitHubリポジトリを作成してPush

#### ステップ1: GitHubで新規リポジトリを作成
1. https://github.com/new にアクセス
2. リポジトリ名を入力（例: `ItemPicking-CodeGenerator`）
3. 説明を追加（例: `YAML-based code generation system for Infor CSI ItemPicking form`）
4. **Public** または **Private** を選択
5. **"Add a README file"のチェックを外す**（既にローカルにファイルがあるため）
6. **"Create repository"** をクリック

#### ステップ2: リモートリポジトリを追加してPush
```powershell
# GitHubのリポジトリURLを設定（例）
git remote add origin https://github.com/ytsutsumi30/ItemPicking-CodeGenerator.git

# リモートリポジトリを確認
git remote -v

# mainブランチにリネーム（GitHubの標準）
git branch -M main

# 初回Push
git push -u origin main
```

---

### オプション2: 既存のリモートリポジトリにPush

既にリモートリポジトリがある場合：

```powershell
# リモートリポジトリを追加
git remote add origin <あなたのリポジトリURL>

# 例: 
# git remote add origin https://github.com/ytsutsumi30/existing-repo.git
# または
# git remote add origin git@github.com:ytsutsumi30/existing-repo.git

# ブランチをmainにリネーム
git branch -M main

# Push
git push -u origin main
```

---

### オプション3: Azure DevOps / GitLab / Bitbucket

#### Azure DevOps
```powershell
git remote add origin https://dev.azure.com/<organization>/<project>/_git/<repo>
git branch -M main
git push -u origin main
```

#### GitLab
```powershell
git remote add origin https://gitlab.com/<username>/<repo>.git
git branch -M main
git push -u origin main
```

#### Bitbucket
```powershell
git remote add origin https://bitbucket.org/<username>/<repo>.git
git branch -M main
git push -u origin main
```

---

## 🔐 認証方法

### 方法1: HTTPS（トークン認証）

GitHubの場合、Personal Access Tokenが必要です：

1. https://github.com/settings/tokens にアクセス
2. **"Generate new token"** → **"Classic"**
3. 権限を選択：`repo` にチェック
4. トークンをコピー
5. Pushコマンド実行時にパスワードの代わりにトークンを入力

```powershell
git push -u origin main
# Username: ytsutsumi30
# Password: <Personal Access Token>
```

### 方法2: SSH（推奨）

SSH鍵を使う方が便利です：

```powershell
# SSH鍵を生成（まだない場合）
ssh-keygen -t ed25519 -C "y.tsutsumi30@gmail.com"

# 公開鍵を表示
Get-Content ~/.ssh/id_ed25519.pub

# GitHubに公開鍵を登録
# https://github.com/settings/ssh/new
```

SSH URLを使用：
```powershell
git remote add origin git@github.com:ytsutsumi30/<repo>.git
git push -u origin main
```

---

## 📝 今後の更新フロー

### 日常的な変更をコミット＆Push

```powershell
# YAMLファイルを編集
code ItemPiking.YAML

# 変更を確認
git status
git diff

# 変更をステージング
git add ItemPiking.YAML

# または、すべての変更をステージング
git add -A

# コミット
git commit -m "Update: グリッド列の追加"

# Push
git push
```

### ブランチを使った開発

```powershell
# 新機能用のブランチを作成
git checkout -b feature/add-new-property

# 変更を加える
code ItemPiking.YAML
.\generate_cs_from_yaml.ps1 -Verbose

# コミット
git add -A
git commit -m "Add new property to cItem class"

# Pushしてプルリクエスト作成
git push -u origin feature/add-new-property
```

---

## 🎯 推奨リポジトリ構成

### README.md を追加する

```powershell
# プロジェクト完成サマリーをREADMEとしてコピー
Copy-Item "プロジェクト完成サマリー.md" "README.md"

git add README.md
git commit -m "Add README.md"
git push
```

### .github/workflows でCI/CD設定（オプション）

```yaml
# .github/workflows/generate-code.yml
name: Generate Code
on:
  push:
    paths:
      - 'ItemPiking.YAML'
  pull_request:
    paths:
      - 'ItemPiking.YAML'

jobs:
  generate:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v3
      
      - name: Install powershell-yaml
        shell: pwsh
        run: Install-Module -Name powershell-yaml -Force
      
      - name: Generate C# code
        shell: pwsh
        run: .\generate_cs_from_yaml.ps1 -OutputFile "ItemPicking.cs" -Verbose
      
      - name: Upload generated code
        uses: actions/upload-artifact@v3
        with:
          name: generated-code
          path: ItemPicking.cs
```

---

## 🔍 トラブルシューティング

### 問題: `remote: Permission denied`

**解決策:**
- Personal Access Tokenの権限を確認
- SSHの場合、鍵がGitHubに登録されているか確認

### 問題: `! [rejected] master -> master (fetch first)`

**解決策:**
```powershell
# リモートの変更を取得してマージ
git pull origin main --rebase
git push origin main
```

### 問題: ファイルサイズが大きすぎる

**解決策:**
```powershell
# 大きなファイルを.gitignoreに追加
echo "*.xlsm" >> .gitignore
git rm --cached *.xlsm
git commit -m "Remove large Excel files from tracking"
```

---

## 📊 現在のリポジトリ統計

```
ブランチ: master
コミット数: 1
ファイル数: 23
総サイズ: 361.7 KB
最終コミット: feec0d4

ファイル構成:
├── .gitignore
├── ItemPiking.YAML (★マスター定義)
├── generate_cs_from_yaml.ps1 (★推奨ツール)
├── ItemPicking.cs (元ファイル)
├── ItemPicking.html
├── ItemPicking_generated.cs
├── Excel生成ツール (.xlsx, .xlsm)
├── PowerShellスクリプト (.ps1)
└── ドキュメント (.md)
```

---

## 🎉 次のステップ

1. **リモートリポジトリを作成** (GitHub推奨)
2. **リモートを追加** (`git remote add origin <URL>`)
3. **Push実行** (`git push -u origin main`)
4. **README.mdを追加** (プロジェクト完成サマリーをコピー)
5. **GitHub Actionsで自動化** (オプション)

---

## 📞 クイックコマンド集

```powershell
# リモート追加（GitHub）
git remote add origin https://github.com/ytsutsumi30/<repo-name>.git

# ブランチをmainにリネーム
git branch -M main

# 初回Push
git push -u origin main

# 以降の通常Push
git add -A
git commit -m "更新内容"
git push

# リモートURL変更
git remote set-url origin <新しいURL>

# リモート削除
git remote remove origin

# ブランチ確認
git branch -a

# リモート確認
git remote -v
```

---

**準備完了！リモートリポジトリのURLを決めてPushしましょう！** 🚀
