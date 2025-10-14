# ItemPiking.YAMLからItemPicking.csを生成するPowerShellスクリプト
# Excel不要、PowerShellのみで完結

param(
    [string]$YamlFile = "ItemPiking.YAML",
    [string]$OutputFile = "",
    [switch]$Verbose
)

# エラーハンドリング
$ErrorActionPreference = "Stop"

# カラー出力関数
function Write-Info($message) { Write-Host $message -ForegroundColor Cyan }
function Write-Success($message) { Write-Host $message -ForegroundColor Green }
function Write-Error($message) { Write-Host $message -ForegroundColor Red }
function Write-Warn($message) { Write-Host $message -ForegroundColor Yellow }

# インデント定数
$INDENT1 = "   "
$INDENT2 = "      "
$INDENT3 = "         "
$INDENT4 = "            "

Write-Info "=========================================="
Write-Info "ItemPicking.cs Code Generator from YAML"
Write-Info "=========================================="
Write-Info ""

# powershell-yamlモジュールの確認
Write-Info "Checking powershell-yaml module..."
if (-not (Get-Module -ListAvailable -Name powershell-yaml)) {
    Write-Warn "powershell-yaml module not found. Installing..."
    try {
        Install-Module -Name powershell-yaml -Scope CurrentUser -Force -AllowClobber
        Write-Success "Module installed successfully!"
    }
    catch {
        Write-Error "Failed to install powershell-yaml module: $_"
        exit 1
    }
}

# モジュールのインポート
Import-Module powershell-yaml

# YAMLファイルの存在確認
if (-not (Test-Path $YamlFile)) {
    Write-Error "YAML file not found: $YamlFile"
    exit 1
}

Write-Info "Loading YAML file: $YamlFile"

# YAMLファイルの読み込み
try {
    $yamlContent = Get-Content -Raw -Path $YamlFile -Encoding UTF8
    $yaml = ConvertFrom-Yaml $yamlContent
    Write-Success "YAML loaded successfully!"
}
catch {
    Write-Error "Failed to parse YAML file: $_"
    exit 1
}

# 出力ファイル名の決定
if ([string]::IsNullOrEmpty($OutputFile)) {
    $OutputFile = $yaml.metadata.outputFileName
}

Write-Info "Output file: $OutputFile"
Write-Info ""

# コード生成用のStringBuilder
$code = New-Object System.Collections.Generic.List[string]

# 1. References
Write-Info "Generating references..."
if ($yaml.references) {
    foreach ($ref in $yaml.references) {
        $code.Add("//<ref>$($ref.assembly)</ref>")
    }
    $code.Add("")
}

# 2. Usings
Write-Info "Generating using statements..."
if ($yaml.usings) {
    foreach ($category in $yaml.usings) {
        foreach ($ns in $category.namespaces) {
            $code.Add("using $ns;")
        }
        # カテゴリ間に空行
        $code.Add("")
    }
}

# 3. Namespace & Class declaration
Write-Info "Generating namespace and class declaration..."
$code.Add("namespace $($yaml.metadata.namespace)")
$code.Add("{")
$code.Add("${INDENT1}public class $($yaml.metadata.className) : $($yaml.metadata.baseClass)")
$code.Add("${INDENT1}{")
$code.Add("${INDENT2}/// <summary>")
$code.Add("${INDENT2}/// グローバル変数")
$code.Add("${INDENT2}/// </summary>")

# 4. Global Variables
Write-Info "Generating global variables..."
if ($yaml.globalVariables) {
    foreach ($gvar in $yaml.globalVariables) {
        $line = "${INDENT2}string $($gvar.name) = `"$($gvar.value)`";"
        if ($gvar.description) {
            $line += "  // $($gvar.description)"
        }
        $code.Add($line)
    }
    $code.Add("")
    $code.Add("")
}

# 5. Item Class
Write-Info "Generating cItem class..."
if ($yaml.itemClass) {
    $code.Add("${INDENT2}/// <summary>")
    $code.Add("${INDENT2}/// $($yaml.itemClass.description)")
    $code.Add("${INDENT2}/// </summary>")
    $code.Add("${INDENT2}public class $($yaml.itemClass.name)")
    $code.Add("${INDENT2}{")
    
    foreach ($propCategory in $yaml.itemClass.properties) {
        if ($propCategory.categoryComment) {
            $code.Add("${INDENT3}// $($propCategory.categoryComment)")
        }
        
        foreach ($prop in $propCategory.items) {
            $code.Add("${INDENT3}[JsonProperty(`"$($prop.jsonName)`")]")
            $code.Add("${INDENT3}public $($prop.type) $($prop.propertyName) { get; set; }")
            $code.Add("")
        }
    }
    
    $code.Add("${INDENT2}}")
    $code.Add("")
}

# 6. callAPI method
Write-Info "Generating callAPI method..."
if ($yaml.methods.callAPI) {
    $method = $yaml.methods.callAPI
    $code.Add("${INDENT2}/// <summary>")
    $code.Add("${INDENT2}/// $($method.description)")
    $code.Add("${INDENT2}/// </summary>")
    
    foreach ($param in $method.parameters) {
        $code.Add("${INDENT2}/// <param name=`"$($param.name)`">$($param.description)</param>")
    }
    
    $code.Add("${INDENT2}public void callAPI(int mode){")
    $code.Add("${INDENT3}callAPIwithTarget(mode,`"`");")
    $code.Add("${INDENT2}}")
    $code.Add("${INDENT2}")
}

# 7. callAPIwithTarget method
Write-Info "Generating callAPIwithTarget method..."
if ($yaml.methods.callAPIwithTarget) {
    $method = $yaml.methods.callAPIwithTarget
    $code.Add("${INDENT2}/// <summary>")
    $code.Add("${INDENT2}/// $($method.description)")
    $code.Add("${INDENT2}/// </summary>")
    
    foreach ($param in $method.parameters) {
        $code.Add("${INDENT2}/// <param name=`"$($param.name)`">$($param.description)</param>")
    }
    
    $code.Add("${INDENT2}public void callAPIwithTarget(int mode,string target){")
    $code.Add("")
    $code.Add("${INDENT3}List<Property> propertiesList = new List<Property>();")
    $code.Add("")
    $code.Add("${INDENT3}//使用するプロパティのデータリストを作成")
    $code.Add("${INDENT3}if(mode == 0){")
    
    # Mode 0 properties
    if ($method.propertyListMode0) {
        foreach ($prop in $method.propertyListMode0) {
            $modifiedVal = if ($prop.modified) { "true" } else { "false" }
            $code.Add("${INDENT4}propertiesList.Add(new Property { Name = `"$($prop.name)`",Value = $($prop.value), Modified = $modifiedVal });")
        }
    }
    
    $code.Add("")
    $code.Add("${INDENT3}}")
    $code.Add("${INDENT3}//データ更新用、get時不要")
    $code.Add("${INDENT3}else{")
    
    # Else properties
    if ($method.propertyListElse) {
        foreach ($prop in $method.propertyListElse) {
            $modifiedVal = if ($prop.modified) { "true" } else { "false" }
            $code.Add("${INDENT4}propertiesList.Add(new Property { Name = `"$($prop.name)`", Value = $($prop.value), Modified = $modifiedVal});")
        }
    }
    
    $code.Add("${INDENT4}//データ特定用、insert時は不要")
    
    # Key properties
    if ($method.propertyListKeys) {
        $code.Add("${INDENT4}if($($method.propertyListKeys.condition)){")
        foreach ($prop in $method.propertyListKeys.items) {
            $modifiedVal = if ($prop.modified) { "true" } else { "false" }
            $code.Add("${INDENT4}   propertiesList.Add(new Property { Name = `"$($prop.name)`", Value = $($prop.value), Modified = $modifiedVal });")
        }
        $code.Add("${INDENT4}}")
    }
    
    $code.Add("${INDENT3}}")
    $code.Add("")
    $code.Add("${INDENT3}//データを取得")
    $code.Add("${INDENT3}string JSONResponse = getData(mode,gIDOName,propertiesList,target);")
    $code.Add("${INDENT3}")
    $code.Add("${INDENT3}//取得した結果はJSONなので、bodyを処理しデータを取得")
    $code.Add("${INDENT3}List<cItem> itemsList = JsonConvert.DeserializeObject<getJSONObject>(JSONResponse).Items;")
    $code.Add("")
    $code.Add("${INDENT3}//処理結果を変数に格納")
    $code.Add("${INDENT3}ThisForm.Variables(`"$($yaml.metadata.resultVariableName)`").Value = GenerateWebSetJson(itemsList);")
    $code.Add("")
    
    # Grid Display
    if ($method.gridDisplay) {
        $grid = $method.gridDisplay
        $code.Add("${INDENT3}//グリッドの初期化")
        $code.Add("${INDENT3}int count = 1;")
        $code.Add("${INDENT3}if(ThisForm.Components[`"$($grid.gridName)`"].GetGridRowCount() > 0 && $($grid.deleteCondition)){")
        $code.Add("${INDENT4}ThisForm.Components[`"$($grid.gridName)`"].DeleteGridRows(1,ThisForm.Components[`"$($grid.gridName)`"].GetGridRowCount());")
        $code.Add("${INDENT3}}")
        $code.Add("")
        $code.Add("${INDENT3}//データを表示")
        $code.Add("${INDENT3}foreach (var item in itemsList)")
        $code.Add("${INDENT3}{")
        $code.Add("${INDENT4}ThisForm.Components[`"$($grid.gridName)`"].InsertGridRows(count,1);")
        
        foreach ($mapping in $grid.mappings) {
            if ($mapping.commented) {
                $code.Add("${INDENT4}//ThisForm.Components[`"$($grid.gridName)`"].SetGridValue(count,$($mapping.column),item.$($mapping.property));")
            }
            elseif ($mapping.conversion) {
                $line = "${INDENT4}ThisForm.Components[`"$($grid.gridName)`"].SetGridValue(count,$($mapping.column),$($mapping.conversion));"
                if ($mapping.comment) {
                    $line += "//$($mapping.comment)"
                }
                $code.Add($line)
            }
            else {
                $code.Add("${INDENT4}ThisForm.Components[`"$($grid.gridName)`"].SetGridValue(count,$($mapping.column),item.$($mapping.property));")
            }
        }
        
        $code.Add("${INDENT4}")
        $code.Add("${INDENT4}count++;")
        $code.Add("${INDENT3}} ")
    }
    
    $code.Add("${INDENT2}}")
    $code.Add("")
}

# 8. Common Processing (from source file)
Write-Info "Loading common processing from source file..."
if ($yaml.commonProcessing) {
    $commonProc = $yaml.commonProcessing
    $sourceFile = $commonProc.source
    
    if (-not (Test-Path $sourceFile)) {
        Write-Warn "Common processing source file not found: $sourceFile"
        Write-Warn "Skipping common processing section..."
    }
    else {
        $startLine = $commonProc.startLine
        $endLine = $commonProc.endLine
        
        Write-Info "Reading lines $startLine to $endLine from $sourceFile"
        
        $allLines = Get-Content -Path $sourceFile -Encoding UTF8
        
        # 行番号は1ベースなので、配列インデックスは0ベース
        $startIndex = $startLine - 1
        $endIndex = $endLine - 1
        
        for ($i = $startIndex; $i -le $endIndex; $i++) {
            if ($i -lt $allLines.Count) {
                $code.Add($allLines[$i])
            }
        }
        
        Write-Success "Common processing loaded: $($endIndex - $startIndex + 1) lines"
    }
}

# 9. ファイルに書き込み
Write-Info ""
Write-Info "Writing to file: $OutputFile"

try {
    # UTF-8 without BOMで出力
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllLines($OutputFile, $code, $utf8NoBom)
    Write-Success "File written successfully!"
}
catch {
    Write-Error "Failed to write file: $_"
    exit 1
}

# 統計情報
Write-Info ""
Write-Info "=========================================="
Write-Success "Code generation completed!"
Write-Info "=========================================="
Write-Info "Total lines: $($code.Count)"
Write-Info "Output file: $OutputFile"
Write-Info "File size: $([math]::Round((Get-Item $OutputFile).Length / 1KB, 2)) KB"

if ($Verbose) {
    Write-Info ""
    Write-Info "Generated sections:"
    Write-Info "  - References: $($yaml.references.Count) items"
    Write-Info "  - Usings: $($yaml.usings.Count) categories"
    Write-Info "  - Global variables: $($yaml.globalVariables.Count) items"
    Write-Info "  - Item properties: $(($yaml.itemClass.properties | ForEach-Object { $_.items.Count } | Measure-Object -Sum).Sum) items"
    Write-Info "  - Methods: callAPI, callAPIwithTarget"
    if ($yaml.commonProcessing) {
        $commonLines = $yaml.commonProcessing.endLine - $yaml.commonProcessing.startLine + 1
        Write-Info "  - Common processing: $commonLines lines"
    }
}

Write-Info ""
Write-Success "Done!"
