# ItemPikingDetailGeneratorYML システム - 使用方法ガイド

## 📋 概要

YAMLベースのコード生成システムで、ItemPiking.mdの設計書から ItemPicking.cs を生成します。

## 🏗️ システム構成

```
ItemPiking.md (設計書)
    ↓
ItemPicking.cs (実装)
    ↓ create_yaml_from_cs.ps1
ItemPiking.YAML (定義ファイル)
    ↓ ItemPikingDetailGeneratorYML.xlsm
ItemPicking.cs (生成)
```

### ファイル一覧

| ファイル名 | 種類 | 説明 |
|-----------|------|------|
| **ItemPiking.YAML** | YAML | コード生成定義ファイル |
| **ItemPikingDetailGeneratorYML.xlsm** | Excel | YAMLからコードを生成するマクロツール |
| **create_yaml_from_cs.ps1** | PowerShell | C#からYAMLを生成 |
| **create_generator_yml.ps1** | PowerShell | Excelマクロツールを生成 |
| **ItemPicking.cs** | C# | 生成対象ファイル |

## 🚀 クイックスタート

### 前提条件

1. **PowerShell 5.1以上**
2. **Excel 2010以上**（マクロ有効）
3. **powershell-yaml モジュール**

```powershell
# モジュールのインストール
Install-Module -Name powershell-yaml -Scope CurrentUser -Force
```

### ステップ1: Excelマクロツールの生成

```powershell
cd "c:\work\work2\出荷検品\ItemPicking"
.\create_generator_yml.ps1
```

生成されるファイル:
- `ItemPikingDetailGeneratorYML.xlsm` (Excelマクロツール)

### ステップ2: YAMLファイルの確認・編集

`ItemPiking.YAML` を開いて、必要に応じて編集します。

### ステップ3: コード生成

1. `ItemPikingDetailGeneratorYML.xlsm` を開く
2. 開発タブ > マクロ > `GenerateCodeFromYaml` を実行
3. `ItemPicking.cs` が生成されます

## 📄 YAML定義ファイルの構造

### 全体構造

```yaml
metadata:           # 基本情報
references:         # DLL参照
usings:             # using文
globalVariables:    # グローバル変数
itemClass:          # cItemクラス定義
methods:            # メソッド定義（業務固有処理）
commonProcessing:   # 共通処理（152行目以降）
```

### セクション詳細

#### 1. metadata セクション

```yaml
metadata:
  outputFileName: "ItemPicking.cs"
  namespace: "Mongoose.FormScripts"
  className: "ItemPicking"
  baseClass: "FormScript"
  resultVariableName: "vJSONResult"
  commonProcessingStartLine: 152
```

| プロパティ | 説明 | 例 |
|-----------|------|-----|
| outputFileName | 出力ファイル名 | ItemPicking.cs |
| namespace | 名前空間 | Mongoose.FormScripts |
| className | クラス名 | ItemPicking |
| baseClass | 基底クラス | FormScript |
| resultVariableName | 結果変数名 | vJSONResult |
| commonProcessingStartLine | 共通処理開始行 | 152 |

#### 2. references セクション

```yaml
references:
  - assembly: "Newtonsoft.Json.dll"
```

生成されるコード:
```csharp
//<ref>Newtonsoft.Json.dll</ref>
```

#### 3. usings セクション

```yaml
usings:
  - category: "base"
    namespaces:
      - "System"
  - category: "mongoose"
    namespaces:
      - "Mongoose.IDO.Protocol"
      - "Mongoose.Scripting"
```

生成されるコード:
```csharp
using System;

using Mongoose.IDO.Protocol;
using Mongoose.Scripting;
```

#### 4. globalVariables セクション

```yaml
globalVariables:
  - name: "gIDOName"
    value: "SLLotLocs"
    description: "ターゲットIDO名"
```

生成されるコード:
```csharp
string gIDOName = "SLLotLocs";  // ターゲットIDO名
```

#### 5. itemClass セクション

```yaml
itemClass:
  name: "cItem"
  description: "取得データ、更新データ作成のクラス"
  properties:
    - category: "data"
      categoryComment: "取得予定のプロパティ"
      items:
        - jsonName: "Whse"
          propertyName: "Whse"
          type: "string"
          description: "倉庫"
```

生成されるコード:
```csharp
/// <summary>
/// 取得データ、更新データ作成のクラス
/// </summary>
public class cItem
{
   // 取得予定のプロパティ
   [JsonProperty("Whse")]
   public string Whse { get; set; }
   ...
}
```

#### 6. methods セクション

##### callAPI メソッド

```yaml
methods:
  callAPI:
    description: "APIを呼び出してデータ取得・グリッド表示を行う"
    parameters:
      - name: "mode"
        type: "int"
        description: "操作モード（0:取得, 1:挿入, 2:更新, 4:削除）"
    implementation: "delegate"
    delegateTo: "callAPIwithTarget"
```

##### callAPIwithTarget メソッド

```yaml
  callAPIwithTarget:
    description: "APIを呼び出してデータ取得・グリッド表示を行う"
    parameters:
      - name: "mode"
        type: "int"
        description: "操作モード"
      - name: "target"
        type: "string"
        description: "更新目標"
```

###### propertyListMode0 (mode=0時のプロパティ)

```yaml
    propertyListMode0:
      - name: "Whse"
        value: '""'
        modified: true
```

生成されるコード:
```csharp
propertiesList.Add(new Property { Name = "Whse",Value = "", Modified = true });
```

###### propertyListElse (更新時のプロパティ)

```yaml
    propertyListElse:
      - name: "Whse"
        value: 'ThisForm.Components["ResultGrid"].GetGridValue(...)'
        modified: true
```

###### propertyListKeys (キープロパティ)

```yaml
    propertyListKeys:
      condition: "mode != 1"
      items:
        - name: "RecordDate"
          value: 'ThisForm.Components["ResultGrid"].GetGridValue(...)'
          modified: true
```

###### gridDisplay (グリッド表示)

```yaml
    gridDisplay:
      gridName: "ResultGrid"
      deleteCondition: 'target == ""'
      mappings:
        - column: 1
          property: "Whse"
        - column: 4
          property: "QtyOnHand"
          conversion: "float.Parse(item.QtyOnHand).ToString()"
          comment: "小数点消すためいったんfloatに変換"
        - column: 5
          property: "PickingNum"
          commented: true
```

#### 7. commonProcessing セクション

```yaml
commonProcessing:
  source: "ItemPicking.cs"
  startLine: 152
  endLine: 481
  methods:
    - name: "getData"
      description: "データをアップデート・取得する"
    - name: "GenerateChangeSetJson"
      description: "Update用JSON文字列を自動生成する"
```

共通処理は元のItemPicking.csファイルから直接読み込まれます。

## 🔧 カスタマイズ方法

### 新しいプロパティを追加

1. `itemClass.properties` に新しいプロパティを追加
2. `methods.callAPIwithTarget.propertyListMode0` に追加
3. `methods.callAPIwithTarget.propertyListElse` に必要に応じて追加
4. `methods.callAPIwithTarget.gridDisplay.mappings` に表示定義を追加
5. マクロを実行して再生成

### グローバル変数を変更

```yaml
globalVariables:
  - name: "gIDOName"
    value: "NewIDOName"  # ← ここを変更
    description: "新しいIDO名"
```

### グリッド表示のカスタマイズ

```yaml
gridDisplay:
  mappings:
    - column: 12  # 新しい列
      property: "NewProperty"
      conversion: "item.NewProperty.ToString()"  # 任意の変換式
      comment: "新しいプロパティ"
```

## 🔄 ワークフロー

### 初回セットアップ

```powershell
# 1. Excelツール生成
.\create_generator_yml.ps1

# 2. YAMLを既存のC#から生成（オプション）
.\create_yaml_from_cs.ps1

# 3. powershell-yamlモジュールのインストール確認
Get-Module -ListAvailable -Name powershell-yaml
```

### 日常的な開発フロー

1. **YAML編集**: `ItemPiking.YAML` を編集して業務ロジックを定義
2. **コード生成**: Excelマクロ `GenerateCodeFromYaml` を実行
3. **検証**: 生成された `ItemPicking.cs` をビルド・テスト
4. **必要に応じてYAML調整**: エラーがあればYAMLを修正して再生成

### 共通処理の更新

共通処理（152行目以降）を更新する場合：

1. `ItemPicking.cs` の152行目以降を直接編集
2. YAMLファイルを再生成（オプション）:
   ```powershell
   .\create_yaml_from_cs.ps1
   ```

## 🧪 テスト・検証

### YAMLファイルの妥当性確認

```powershell
Import-Module powershell-yaml
$yaml = Get-Content -Raw -Path "ItemPiking.YAML" | ConvertFrom-Yaml
Write-Host "Class: $($yaml.metadata.className)"
Write-Host "Properties: $($yaml.itemClass.properties[0].items.Count)"
```

### 生成されたコードのビルド

```powershell
# Visual Studio環境でビルド
msbuild EditWinStudioScript.csproj /p:Configuration=Release
```

## 📊 比較表: Excel vs YAML方式

| 項目 | Excel方式 | YAML方式 |
|-----|----------|---------|
| **編集のしやすさ** | GUI操作 | テキストエディタ |
| **バージョン管理** | 困難（バイナリ） | 容易（テキスト） |
| **差分確認** | 困難 | 容易（git diff） |
| **自動化** | 制限あり | 容易（CI/CD対応） |
| **学習コスト** | 低い | 中程度（YAML構文） |
| **複雑な定義** | やや困難 | 容易（階層構造） |
| **可搬性** | Excelが必要 | テキストのみ |

## 🎯 使い分けガイド

### Excel方式が適している場合

- GUI操作が好き
- 小規模プロジェクト
- 単独開発
- バージョン管理不要

### YAML方式が適している場合

- テキストエディタでの作業が好き
- 中〜大規模プロジェクト
- チーム開発
- バージョン管理必須
- CI/CD自動化
- 複数プロジェクトでのテンプレート共有

## 🐛 トラブルシューティング

### YAMLパースエラー

```
Error: Failed to convert YAML to JSON via PowerShell.
```

**解決策**:
1. YAMLの構文をチェック（インデントなど）
2. powershell-yamlモジュールの再インストール:
   ```powershell
   Remove-Module powershell-yaml -Force
   Install-Module -Name powershell-yaml -Force -Scope CurrentUser
   ```

### マクロ実行エラー

```
Error: YAML file not found
```

**解決策**:
1. Configシートの`YamlPath`を確認
2. 相対パスの場合、Excelファイルと同じフォルダにYAMLがあるか確認

### 生成されたコードがビルドエラー

**解決策**:
1. YAMLの型定義を確認（`type: "string"` など）
2. valueフィールドの引用符をチェック
3. 共通処理のソースファイルパスを確認

## 📚 参考資料

- **設計書**: `ItemPiking.md`
- **Excel版**: `ItemPikingDetailGenerator.xlsx`
- **サンプルYAML**: `ItemPiking.YAML`
- **PowerShell-YAML**: https://github.com/cloudbase/powershell-yaml

## 🔗 関連ファイル

| ファイル | 説明 |
|---------|------|
| ItemPiking.md | 設計書（元資料） |
| ItemPicking.cs | 生成対象C#ファイル |
| ItemPiking.YAML | コード生成定義 |
| ItemPikingDetailGenerator.xlsx | Excel方式のツール |
| ItemPikingDetailGeneratorYML.xlsm | YAML方式のツール |
| create_yaml_from_cs.ps1 | C#→YAML変換 |
| create_generator_yml.ps1 | Excelツール生成 |

## 💡 ベストプラクティス

1. **YAMLファイルをバージョン管理**
   ```bash
   git add ItemPiking.YAML
   git commit -m "Update property definitions"
   ```

2. **コメントを活用**
   ```yaml
   # このプロパティは在庫数を表します
   - jsonName: "QtyOnHand"
     propertyName: "QtyOnHand"
   ```

3. **定期的なバックアップ**
   ```powershell
   Copy-Item ItemPiking.YAML ItemPiking.YAML.bak
   ```

4. **変更履歴の記録**
   ```yaml
   metadata:
     version: "1.0.1"
     lastModified: "2025-10-13"
     changeLog: "Added new property for tracking"
   ```

---

**作成日**: 2025年10月13日  
**バージョン**: 1.0  
**作成者**: AI Code Generator System
