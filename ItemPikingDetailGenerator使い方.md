# ItemPikingDetailGenerator.xlsx 使い方

## 概要

`ItemPikingDetailGenerator.xlsx`は、ItemPiking.mdの設計書から`ItemPicking.cs`を自動生成するためのExcelマクロツールです。

## ファイル構成

- **ItemPikingDetailGenerator.xlsx**: コード生成用Excelファイル
- **create_generator_ItemPicking.ps1**: Excel生成用PowerShellスクリプト
- **ItemPicking.cs**: 生成されるC#ソースファイル（152行目で業務固有処理と共通処理に分割）

## Excelシート構成

### 1. Readme シート
使い方と構造の説明

### 2. SpecData シート（業務固有処理：1-151行目）

#### 名前付き範囲
| 名前 | 値 | 説明 |
|------|-----|------|
| OutputPath | ItemPicking.cs | 出力ファイル名 |
| ReferenceLine | //<ref>Newtonsoft.Json.dll</ref> | DLL参照ディレクティブ |
| NamespaceName | Mongoose.FormScripts | 名前空間 |
| ClassName | ItemPicking | クラス名 |
| ResultVariableName | vJSONResult | JSON結果変数名 |

#### テーブル

**Usings**: 使用する名前空間
- Order: 順序
- Namespace: 名前空間名
- Category: カテゴリ（base/mongoose/framework/external）

**GlobalConfig**: グローバル変数設定
- Name: 変数名
- Value: 初期値
- Description: 説明

例：
```csharp
string gIDOName = "SLLotLocs";  // ターゲットIDO名
string gSSO = "1";              // SSOの使用、基本は1　1：はい　0：いいえ
```

**ItemProperties**: cItemクラスのプロパティ定義
- JsonName: JSON属性名
- PropertyName: C#プロパティ名
- Type: データ型
- Description: 説明

例：
```csharp
[JsonProperty("Whse")]
public string Whse { get; set; }
```

**CallAPIPropertiesMode0**: callAPI メソッドのmode=0時のプロパティリスト
- Name: プロパティ名
- InitialValue: 初期値
- Modified: Modified属性値
- Comment: コメント

**CallAPIPropertiesElse**: callAPI メソッドのmode=更新時のプロパティリスト
- Name: プロパティ名
- InitialValue: 値の取得元（例：ThisForm.Components["ResultGrid"].GetGridValue(...)）
- Modified: Modified属性値
- Comment: コメント

**CallAPIPropertiesKey**: データ特定用キープロパティ
- Name: プロパティ名（RecordDate, RowPointer, _ItemId）
- InitialValue: 値の取得元
- Modified: Modified属性値
- Comment: コメント

**GridDisplay**: グリッド表示のマッピング
- Column: グリッド列番号
- Property: cItemのプロパティ名
- Conversion: 変換式（例：float.Parse(item.QtyOnHand).ToString()）
- Comment: コメント

### 3. CommonTail シート（共通処理：152行目以降）

**TailLines**: ItemPicking.csの152行目以降の共通処理コード
- Code: 各行のコード

共通処理の内容：
- `getData()`: ION APIを呼び出してデータ取得・更新
- `GenerateChangeSetJson()`: Update用JSON生成
- `GenerateFilter()`: フィルター文字列生成
- `addScanTarget()`: QR読み取り対象をGridに追加
- `matchGridItems()`: Grid内の品目IDを照合
- `GenerateWebSetJson()`: web用JSON生成
- 各種JSONクラス定義（getJSONObject, Property, Change, UpdateJSONObject, WebJSONObject）

## 使用方法

### 1. Excelファイルの生成（初回のみ）

PowerShellで以下を実行：
```powershell
cd "c:\work\work2\出荷検品\ItemPicking"
.\create_generator_ItemPicking.ps1
```

### 2. Excelファイルを開く

`ItemPikingDetailGenerator.xlsx`を開く

### 3. 業務固有の設定を編集（必要に応じて）

SpecDataシートで以下を編集：
- グローバル変数（GlobalConfigテーブル）
- プロパティ定義（ItemPropertiesテーブル）
- callAPIのロジック（CallAPIProperties*テーブル）
- グリッド表示マッピング（GridDisplayテーブル）

### 4. マクロを実行

方法1: Excel内で実行
1. 開発タブを有効化（ファイル > オプション > リボンのユーザー設定 > 開発）
2. 開発タブ > マクロ > `GenerateItemPicking` を選択
3. 実行ボタンをクリック

方法2: PowerShellで実行
```powershell
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Open("$(Get-Location)\ItemPikingDetailGenerator.xlsx")
$excel.Run('GenerateItemPicking')
$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
```

### 5. 生成結果の確認

`ItemPicking.cs`がExcelファイルと同じフォルダに生成されます。

## 生成されるコード構造

```
ItemPicking.cs (481行)
├── 1-151行目: 業務固有処理（SpecDataシートから生成）
│   ├── using文
│   ├── namespace宣言
│   ├── グローバル変数
│   ├── cItemクラス定義
│   ├── callAPI()メソッド
│   └── callAPIwithTarget()メソッド
│
└── 152-481行目: 共通処理（CommonTailシートから挿入）
    ├── getData()
    ├── GenerateChangeSetJson()
    ├── GenerateFilter()
    ├── addScanTarget()
    ├── matchGridItems()
    ├── GenerateWebSetJson()
    └── 各種クラス定義
```

## カスタマイズ方法

### 新しいプロパティを追加する場合

1. **ItemPropertiesテーブル**に新しい行を追加
2. **CallAPIPropertiesMode0テーブル**に新しい行を追加
3. **CallAPIPropertiesElse/Keyテーブル**に必要に応じて追加
4. **GridDisplayテーブル**に表示が必要な場合は追加
5. マクロを実行して再生成

### グローバル変数を変更する場合

1. **GlobalConfigテーブル**の該当行を編集
2. マクロを実行して再生成

### 共通処理を修正する場合

1. **CommonTailシート**の該当行を編集
2. マクロを実行して再生成

または、元の`ItemPicking.cs`を修正後、`create_generator_ItemPicking.ps1`を再実行してExcelを再生成

## トラブルシューティング

### マクロが実行できない
- Excelのマクロセキュリティ設定を確認
- ファイル > オプション > セキュリティセンター > セキュリティセンターの設定 > マクロの設定
- 「すべてのマクロを有効にする」または「デジタル署名されたマクロを除き、すべてのマクロを無効にする」に変更

### 生成されたコードがコンパイルエラーになる
- SpecDataシートのテーブルデータを確認
- 特にInitialValueやConversionフィールドの構文を確認
- CommonTailシートのコードに構文エラーがないか確認

### Excelファイルの再生成が必要な場合
```powershell
.\create_generator_ItemPicking.ps1
```
を実行すると、最新の`ItemPicking.cs`から共通処理を読み込んでExcelを再生成します。

## 参照

- 設計書: `ItemPiking.md`
- 参考実装: `ArrivalLocationDetailGenerator.xlsx`
- 生成スクリプト: `create_generator.ps1`（ArrivalLocationDetail用）
