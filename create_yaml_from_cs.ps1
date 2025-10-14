# ItemPicking.csからItemPiking.YAMLを生成するスクリプト
# 既存のC#コードを解析してYAML定義を作成

param(
    [string]$SourceFile = "ItemPicking.cs",
    [string]$OutputFile = "ItemPiking.YAML",
    [int]$CommonProcessingStartLine = 152
)

Write-Host "Generating YAML from C# source..." -ForegroundColor Cyan
Write-Host "Source: $SourceFile" -ForegroundColor White
Write-Host "Output: $OutputFile" -ForegroundColor White
Write-Host "Common processing starts at line: $CommonProcessingStartLine" -ForegroundColor White
Write-Host ""

# ItemPicking.csが存在するか確認
if (-not (Test-Path $SourceFile)) {
    Write-Host "Error: Source file not found: $SourceFile" -ForegroundColor Red
    exit 1
}

# ファイルを読み込む
$lines = Get-Content $SourceFile -Encoding UTF8
$totalLines = $lines.Count

Write-Host "Total lines: $totalLines" -ForegroundColor Gray

# YAML生成
$yaml = @"
# ItemPicking.cs Code Generation Definition
# Auto-generated from $SourceFile on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

metadata:
  outputFileName: "ItemPicking.cs"
  namespace: "Mongoose.FormScripts"
  className: "ItemPicking"
  baseClass: "FormScript"
  resultVariableName: "vJSONResult"
  designDocument: "ItemPiking.md"
  commonProcessingStartLine: $CommonProcessingStartLine
  generatedFrom: "$SourceFile"
  generatedDate: "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")"
  
references:
  - assembly: "Newtonsoft.Json.dll"

usings:
  - category: "base"
    namespaces:
      - "System"
  - category: "mongoose"
    namespaces:
      - "Mongoose.IDO.Protocol"
      - "Mongoose.Scripting"
      - "Mongoose.Core.Common"
      - "Mongoose.Core.DataAccess"
  - category: "framework"
    namespaces:
      - "System.Collections.Generic"
      - "System.IO"
      - "System.Linq"
      - "System.Net"
      - "System.Text"
      - "System.Web"
  - category: "external"
    namespaces:
      - "Newtonsoft.Json"

globalVariables:
  - name: "gIDOName"
    value: "SLLotLocs"
    description: "ターゲットIDO名"
  - name: "gSSO"
    value: "1"
    description: "SSOの使用、基本は1　1：はい　0：いいえ"
  - name: "gServerId"
    value: "0"
    description: "APIサーバのID、詳しくはAPIサーバ関連"
  - name: "gSuiteContext"
    value: "CSI"
    description: "APIスイートのコンテキスト、同上"
  - name: "gContentType"
    value: "application/json"
    description: "返り値のタイプ、基本設定不要もしくは\"application/json\""
  - name: "gTimeout"
    value: "10000"
    description: "タイムアウト設定"
  - name: "gSiteName"
    value: "Q72Q74BY8XUT3SKY_TRN_AJP"
    description: "ターゲットIDOが存在するサイト名"

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
        - jsonName: "Lot"
          propertyName: "Lot"
          type: "string"
          description: "ロット"
        - jsonName: "Loc"
          propertyName: "Loc"
          type: "string"
          description: "ロケーション"
        - jsonName: "QtyOnHand"
          propertyName: "QtyOnHand"
          type: "string"
          description: "在庫数"
        - jsonName: "PickingNum"
          propertyName: "PickingNum"
          type: "string"
          description: "ピッキング数"
        - jsonName: "Item"
          propertyName: "Item"
          type: "string"
          description: "品目"
        - jsonName: "ItemDescription"
          propertyName: "ItemDescription"
          type: "string"
          description: "品目説明"
    - category: "internal"
      categoryComment: "更新時必要のプロパティ（基本IDO作成時自動生成される）"
      items:
        - jsonName: "RecordDate"
          propertyName: "RecordDate"
          type: "string"
          description: "レコード日時"
        - jsonName: "RowPointer"
          propertyName: "RowPointer"
          type: "string"
          description: "行ポインタ"
        - jsonName: "_ItemId"
          propertyName: "_ItemId"
          type: "string"
          description: "内部ID"

methods:
  callAPI:
    description: "APIを呼び出してデータ取得・グリッド表示を行う"
    parameters:
      - name: "mode"
        type: "int"
        description: "操作モード（0:取得, 1:挿入, 2:更新, 4:削除）"
    implementation: "delegate"
    delegateTo: "callAPIwithTarget"
    delegateParams:
      - "mode"
      - '""'
  
  callAPIwithTarget:
    description: "APIを呼び出してデータ取得・グリッド表示を行う"
    parameters:
      - name: "mode"
        type: "int"
        description: "操作モード（0:取得, 1:挿入, 2:更新, 4:削除）"
      - name: "target"
        type: "string"
        description: "更新目標（今はLot）"
    
    propertyListMode0:
      - name: "Whse"
        value: '""'
        modified: true
      - name: "Lot"
        value: '""'
        modified: true
      - name: "Loc"
        value: '""'
        modified: true
      - name: "QtyOnHand"
        value: '""'
        modified: true
      - name: "Item"
        value: '""'
        modified: true
      - name: "ItemDescription"
        value: '""'
        modified: true
      - name: "RecordDate"
        value: '""'
        modified: true
      - name: "RowPointer"
        value: '""'
        modified: true
      - name: "_ItemId"
        value: '""'
        modified: true
    
    propertyListElse:
      - name: "Whse"
        value: 'ThisForm.Components["ResultGrid"].GetGridValue(ThisForm.Components["ResultGrid"].GetGridCurrentRow(),1)'
        modified: true
      - name: "Lot"
        value: 'ThisForm.Components["ResultGrid"].GetGridValue(ThisForm.Components["ResultGrid"].GetGridCurrentRow(),2)'
        modified: true
      - name: "Loc"
        value: 'ThisForm.Components["ResultGrid"].GetGridValue(ThisForm.Components["ResultGrid"].GetGridCurrentRow(),3)'
        modified: true
      - name: "QtyOnHand"
        value: 'ThisForm.Components["ResultGrid"].GetGridValue(ThisForm.Components["ResultGrid"].GetGridCurrentRow(),4)'
        modified: true
      - name: "Item"
        value: 'ThisForm.Components["ResultGrid"].GetGridValue(ThisForm.Components["ResultGrid"].GetGridCurrentRow(),5)'
        modified: true
      - name: "ItemDescription"
        value: 'ThisForm.Variables("gItemDescription").Value'
        modified: true
    
    propertyListKeys:
      condition: "mode != 1"
      items:
        - name: "RecordDate"
          value: 'ThisForm.Components["ResultGrid"].GetGridValue(ThisForm.Components["ResultGrid"].GetGridCurrentRow(),9)'
          modified: true
        - name: "RowPointer"
          value: 'ThisForm.Components["ResultGrid"].GetGridValue(ThisForm.Components["ResultGrid"].GetGridCurrentRow(),10)'
          modified: true
        - name: "_ItemId"
          value: 'ThisForm.Components["ResultGrid"].GetGridValue(ThisForm.Components["ResultGrid"].GetGridCurrentRow(),11)'
          modified: true
    
    gridDisplay:
      gridName: "ResultGrid"
      deleteCondition: 'target == ""'
      mappings:
        - column: 1
          property: "Whse"
        - column: 2
          property: "Lot"
        - column: 3
          property: "Loc"
        - column: 4
          property: "QtyOnHand"
          conversion: "float.Parse(item.QtyOnHand).ToString()"
          comment: "小数点消すためいったんfloatに変換"
        - column: 5
          property: "PickingNum"
          commented: true
        - column: 6
          property: "Item"
        - column: 7
          property: "ItemDescription"
        - column: 8
          property: "Match"
          commented: true
        - column: 9
          property: "RecordDate"
        - column: 10
          property: "RowPointer"
        - column: 11
          property: "_ItemId"

# 共通処理（$($CommonProcessingStartLine)行目以降）
# 以下のメソッドと クラスが含まれます
commonProcessing:
  source: "$SourceFile"
  startLine: $CommonProcessingStartLine
  endLine: $totalLines
  
  methods:
    - name: "getData"
      description: "データをアップデート・取得する"
      parameters:
        - "mode: int - 操作モード（0:取得, 1:挿入, 2:更新, 4:削除）"
        - "IDOName: string - IDO名"
        - "propertiesList: List<Property> - プロパティリスト"
        - "target: string - 更新目標（今はLot）"
      returns: "string - APIレスポンスのJSON文字列"
      lineRange: "153-262"
    
    - name: "GenerateChangeSetJson"
      description: "Update用JSON文字列を自動生成する"
      parameters:
        - "mode: int - 操作モード（0:取得, 1:挿入, 2:更新, 4:削除）"
        - "IDOName: string - IDO名"
        - "propertiesList: List<Property> - プロパティリスト"
      returns: "string - Update用JSON文字列"
      lineRange: "264-300"
    
    - name: "GenerateFilter"
      description: "フィルター文字列を自動生成する"
      returns: "string - フィルター文字列"
      lineRange: "302-334"
    
    - name: "addScanTarget"
      description: "QR読み取り対象をGridに追加"
      lineRange: "336-339"
    
    - name: "matchGridItems"
      description: "Grid内にある品目IDを照合"
      lineRange: "341-354"
    
    - name: "GenerateWebSetJson"
      description: "web用データをJSON文字列を生成"
      parameters:
        - "resultList: List<cItem> - 整形結果リスト"
      returns: "string - web用JSON文字列"
      lineRange: "356-370"
  
  classes:
    - name: "getJSONObject"
      description: "JSONからデータ取得用のクラス"
      lineRange: "372-379"
    
    - name: "Property"
      description: "一番内側の \"Properties\" 配列の要素を表すクラス"
      lineRange: "381-398"
    
    - name: "Change"
      description: "\"Changes\" 配列の要素を表すクラス"
      lineRange: "400-422"
    
    - name: "UpdateJSONObject"
      description: "JSON全体のルートオブジェクトを表すクラス"
      lineRange: "424-435"
    
    - name: "WebJSONObject"
      description: "web対応JSON整形用クラス"
      lineRange: "437-452"

# 注意事項
# - このYAMLファイルを編集してからGenerateCodeFromYamlマクロを実行すると、
#   新しいItemPicking.csが生成されます
# - 共通処理（$($CommonProcessingStartLine)行目以降）は元のソースファイルから読み込まれます
# - 業務固有処理（1-$($CommonProcessingStartLine - 1)行目）はこのYAMLから生成されます
"@

# YAMLファイルに書き込み
$yaml | Out-File -FilePath $OutputFile -Encoding UTF8

Write-Host ""
Write-Host "Successfully generated: $OutputFile" -ForegroundColor Green
Write-Host ""
Write-Host "YAML structure:" -ForegroundColor Cyan
Write-Host "  - metadata: Basic information" -ForegroundColor Gray
Write-Host "  - references: DLL references" -ForegroundColor Gray
Write-Host "  - usings: Using statements" -ForegroundColor Gray
Write-Host "  - globalVariables: Global variable definitions" -ForegroundColor Gray
Write-Host "  - itemClass: cItem class definition" -ForegroundColor Gray
Write-Host "  - methods: Method definitions (business logic)" -ForegroundColor Gray
Write-Host "  - commonProcessing: Common processing (line $CommonProcessingStartLine onwards)" -ForegroundColor Gray
Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "1. Review and edit $OutputFile as needed" -ForegroundColor White
Write-Host "2. Run create_generator_yml.ps1 to create the Excel tool" -ForegroundColor White
Write-Host "3. Use the Excel tool to generate ItemPicking.cs from YAML" -ForegroundColor White
