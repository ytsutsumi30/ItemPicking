# generate_cs_from_yaml.ps1 使用方法

## 概要

`generate_cs_from_yaml.ps1`は、YAMLファイルから直接C#コード（ItemPicking.cs）を生成するPowerShellスクリプトです。

**Excel不要！PowerShellだけで完結！**

## 特徴

✅ **Excel不要**: PowerShellとテキストエディタだけで作業可能  
✅ **高速**: Excelを起動しないため、数秒で生成完了  
✅ **自動化対応**: CI/CDパイプラインに組み込み可能  
✅ **クロスプラットフォーム**: PowerShell Core (pwsh) 対応  
✅ **詳細ログ**: -Verboseオプションで詳細情報表示  

## 前提条件

1. **PowerShell 5.1以上** または **PowerShell Core 7.x**
2. **powershell-yaml モジュール**

```powershell
# モジュールのインストール（初回のみ）
Install-Module -Name powershell-yaml -Scope CurrentUser -Force
```

## 基本的な使い方

### 1. 最もシンプルな実行

```powershell
# ItemPiking.YAML から ItemPicking.cs を生成
.\generate_cs_from_yaml.ps1
```

YAMLファイルの`metadata.outputFileName`に指定されたファイル名で生成されます。

### 2. 出力ファイル名を指定

```powershell
# 別のファイル名で出力
.\generate_cs_from_yaml.ps1 -OutputFile "ItemPicking_v2.cs"
```

### 3. 詳細情報を表示

```powershell
# 詳細なログを表示
.\generate_cs_from_yaml.ps1 -Verbose
```

出力例:
```
==========================================
ItemPicking.cs Code Generator from YAML
==========================================

Checking powershell-yaml module...
Loading YAML file: ItemPiking.YAML
YAML loaded successfully!
Output file: ItemPicking.cs

Generating references...
Generating using statements...
Generating global variables...
Generating cItem class...
Generating callAPI method...
Generating callAPIwithTarget method...
Loading common processing from source file...
Reading lines 152 to 481 from ItemPicking.cs
Common processing loaded: 330 lines

Writing to file: ItemPicking.cs
File written successfully!

==========================================
Code generation completed!
==========================================
Total lines: 483
Output file: ItemPicking.cs
File size: 19.85 KB

Generated sections:
  - References: 1 items
  - Usings: 4 categories
  - Global variables: 7 items
  - Item properties: 10 items
  - Methods: callAPI, callAPIwithTarget
  - Common processing: 330 lines

Done!
```

### 4. 別のYAMLファイルを使用

```powershell
# 別のYAMLファイルから生成
.\generate_cs_from_yaml.ps1 -YamlFile "ItemPiking_v2.YAML" -OutputFile "ItemPicking_v2.cs"
```

## パラメータ一覧

| パラメータ | 型 | デフォルト | 説明 |
|-----------|-----|----------|------|
| `-YamlFile` | string | "ItemPiking.YAML" | 入力YAMLファイルのパス |
| `-OutputFile` | string | (YAMLで指定) | 出力C#ファイルのパス |
| `-Verbose` | switch | false | 詳細情報を表示 |

## ワークフロー

### 開発フロー

```
1. YAMLファイルを編集
   ↓
2. PowerShellスクリプトを実行
   .\generate_cs_from_yaml.ps1
   ↓
3. 生成されたC#ファイルを確認
   ↓
4. ビルド・テスト
```

### 例: 新しいプロパティを追加

#### ステップ1: YAMLを編集

`ItemPiking.YAML`を開いて、新しいプロパティを追加：

```yaml
itemClass:
  properties:
    - category: "data"
      items:
        # 既存のプロパティ...
        - jsonName: "NewProperty"
          propertyName: "NewProperty"
          type: "string"
          description: "新しいプロパティ"
```

#### ステップ2: コード生成

```powershell
.\generate_cs_from_yaml.ps1 -Verbose
```

#### ステップ3: 確認

生成された`ItemPicking.cs`を開いて確認。

## CI/CD統合

### Azure DevOps Pipeline

```yaml
steps:
  - task: PowerShell@2
    displayName: 'Generate C# from YAML'
    inputs:
      targetType: 'filePath'
      filePath: 'generate_cs_from_yaml.ps1'
      arguments: '-Verbose'
      pwsh: true
  
  - task: VSBuild@1
    displayName: 'Build Solution'
    inputs:
      solution: 'EditWinStudioScript.sln'
```

### GitHub Actions

```yaml
name: Generate and Build

on: [push, pull_request]

jobs:
  build:
    runs-on: windows-latest
    steps:
      - uses: actions/checkout@v2
      
      - name: Install powershell-yaml
        shell: pwsh
        run: Install-Module -Name powershell-yaml -Force -Scope CurrentUser
      
      - name: Generate C# from YAML
        shell: pwsh
        run: ./generate_cs_from_yaml.ps1 -Verbose
      
      - name: Build
        run: msbuild EditWinStudioScript.csproj /p:Configuration=Release
```

### GitLab CI

```yaml
stages:
  - generate
  - build

generate:
  stage: generate
  script:
    - pwsh -Command "Install-Module -Name powershell-yaml -Force -Scope CurrentUser"
    - pwsh -File generate_cs_from_yaml.ps1 -Verbose
  artifacts:
    paths:
      - ItemPicking.cs
```

## バッチ処理

### 複数のYAMLファイルを一括生成

```powershell
# batch_generate.ps1
$yamlFiles = Get-ChildItem -Filter "*.YAML"

foreach ($yaml in $yamlFiles) {
    Write-Host "Processing: $($yaml.Name)" -ForegroundColor Cyan
    $outputName = $yaml.BaseName + ".cs"
    .\generate_cs_from_yaml.ps1 -YamlFile $yaml.Name -OutputFile $outputName -Verbose
    Write-Host ""
}

Write-Host "All files generated!" -ForegroundColor Green
```

## エラーハンドリング

### エラー例1: YAMLファイルが見つからない

```
Error: YAML file not found: ItemPiking.YAML
```

**解決策**: ファイル名とパスを確認してください。

### エラー例2: YAML構文エラー

```
Error: Failed to parse YAML file: ...
```

**解決策**: YAMLの構文を確認してください（インデント、引用符など）。

### エラー例3: powershell-yamlモジュールがない

```
powershell-yaml module not found. Installing...
```

**解決策**: 自動的にインストールされます。管理者権限が必要な場合があります。

### エラー例4: 共通処理ファイルが見つからない

```
Warning: Common processing source file not found: ItemPicking.cs
Skipping common processing section...
```

**解決策**: `commonProcessing.source`で指定されたファイルが存在することを確認してください。

## デバッグ

### 生成されたコードを確認

```powershell
# 最初の100行を表示
Get-Content ItemPicking.cs -TotalCount 100
```

### 元のファイルと比較

```powershell
# ファイルサイズ比較
$original = Get-Item "ItemPicking_original.cs"
$generated = Get-Item "ItemPicking.cs"

Write-Host "Original:  $($original.Length) bytes"
Write-Host "Generated: $($generated.Length) bytes"
Write-Host "Diff:      $($generated.Length - $original.Length) bytes"
```

### 差分チェック（Git）

```powershell
# Gitで差分を確認
git diff ItemPicking.cs
```

## パフォーマンス

### 実行時間の測定

```powershell
Measure-Command { .\generate_cs_from_yaml.ps1 }
```

一般的な実行時間:
- **PowerShellスクリプト**: 約2-3秒
- **Excelマクロ**: 約10-15秒

**PowerShellスクリプトは約5倍高速！**

## Excel方式との比較

| 項目 | PowerShellスクリプト | Excelマクロ |
|-----|-------------------|------------|
| **Excel必要** | ✗ | ✓ |
| **実行速度** | ⚡ 高速 (2-3秒) | 🐢 やや遅い (10-15秒) |
| **CI/CD統合** | ✓ 容易 | △ 困難 |
| **クロスプラットフォーム** | ✓ (PowerShell Core) | ✗ (Windowsのみ) |
| **自動化** | ✓ 完全自動化可能 | △ 制限あり |
| **デバッグ** | ✓ 容易 | △ やや困難 |
| **学習コスト** | 中 | 低 |
| **GUI操作** | ✗ | ✓ |

## ベストプラクティス

### 1. YAMLファイルのバージョン管理

```bash
git add ItemPiking.YAML
git commit -m "Update property definitions"
```

### 2. 生成前のバックアップ

```powershell
# 既存ファイルをバックアップ
if (Test-Path "ItemPicking.cs") {
    Copy-Item "ItemPicking.cs" "ItemPicking.cs.bak"
}

# コード生成
.\generate_cs_from_yaml.ps1
```

### 3. 自動ビルドスクリプト

```powershell
# build.ps1
Write-Host "Generating code from YAML..." -ForegroundColor Cyan
.\generate_cs_from_yaml.ps1 -Verbose

Write-Host "Building project..." -ForegroundColor Cyan
msbuild EditWinStudioScript.csproj /p:Configuration=Release

Write-Host "Build completed!" -ForegroundColor Green
```

### 4. コード品質チェック

```powershell
# 生成後に自動でコード品質をチェック
.\generate_cs_from_yaml.ps1

# 行数確認
$lines = (Get-Content ItemPicking.cs).Count
Write-Host "Total lines: $lines"

# 構文チェック（仮想的な例）
# csc.exe /t:library /nologo ItemPicking.cs
```

## トラブルシューティング

### Q1: スクリプトが実行できない

```
.\generate_cs_from_yaml.ps1 : このシステムではスクリプトの実行が無効になっているため...
```

**解決策**:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Q2: UTF-8エンコーディング問題

生成されたファイルの文字化け対策：

スクリプトは自動的にUTF-8 (BOMなし)で出力します。

### Q3: 行末の改行コード

Windows: CRLF (`\r\n`)  
Linux/Mac: LF (`\n`)

PowerShellは自動的にプラットフォーム標準の改行コードを使用します。

## 高度な使用例

### カスタムYAML処理

```powershell
# YAMLを読み込んで加工してから生成
$yaml = Get-Content -Raw ItemPiking.YAML | ConvertFrom-Yaml

# 動的にプロパティを追加
$newProp = @{
    jsonName = "DynamicProperty"
    propertyName = "DynamicProperty"
    type = "string"
    description = "動的に追加"
}
$yaml.itemClass.properties[0].items += $newProp

# 加工したYAMLを一時ファイルに保存
$yaml | ConvertTo-Yaml | Out-File temp.YAML

# 一時ファイルから生成
.\generate_cs_from_yaml.ps1 -YamlFile temp.YAML
```

## まとめ

`generate_cs_from_yaml.ps1`は、Excel不要でYAMLから直接C#コードを生成できる強力なツールです。

### 推奨される使い方

- **開発中**: Excel方式（GUI操作が直感的）
- **本番ビルド**: PowerShellスクリプト（自動化、高速）
- **CI/CD**: PowerShellスクリプト（完全自動化）

### 次のステップ

1. YAMLファイルをカスタマイズ
2. スクリプトを実行してコード生成
3. ビルド・テストで検証
4. CI/CDパイプラインに組み込み

---

**スクリプトバージョン**: 1.0  
**作成日**: 2025年10月13日  
**対応PowerShell**: 5.1+ / Core 7.x+
