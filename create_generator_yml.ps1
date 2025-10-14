# ItemPikingDetailGeneratorYML.xlsm 生成スクリプト
# YAMLファイルからItemPicking.csを生成するExcelマクロツールを作成

$macroCode = @'
Option Explicit

Private Const INDENT1 As String = "   "
Private Const INDENT2 As String = "      "
Private Const INDENT3 As String = "         "
Private Const INDENT4 As String = "            "

' YAMLパーサー用の簡易関数群
Private Function ReadYamlFile(filePath As String) As Object
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(filePath) Then
        MsgBox "YAML file not found: " & filePath, vbExclamation
        Set ReadYamlFile = Nothing
        Exit Function
    End If
    
    Dim stream As Object
    Set stream = fso.OpenTextFile(filePath, 1, False, 0)
    Dim content As String
    content = stream.ReadAll()
    stream.Close
    
    ' 簡易YAMLパーサー（PowerShell経由）
    Dim yaml As Object
    Set yaml = ParseYamlContent(content)
    Set ReadYamlFile = yaml
End Function

Private Function ParseYamlContent(content As String) As Object
    ' PowerShellを使用してYAMLをパース
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    
    ' 一時ファイルにYAMLを保存
    Dim tempYaml As String
    tempYaml = Environ("TEMP") & "\temp_yaml_" & Format(Now, "yyyymmddhhnnss") & ".yml"
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ts As Object
    Set ts = fso.CreateTextFile(tempYaml, True, False)
    ts.Write content
    ts.Close
    
    ' PowerShellでYAMLをJSONに変換
    Dim tempJson As String
    tempJson = Environ("TEMP") & "\temp_json_" & Format(Now, "yyyymmddhhnnss") & ".json"
    
    Dim psCommand As String
    psCommand = "powershell -NoProfile -Command ""$yaml = Get-Content -Raw -Path '" & tempYaml & "' | ConvertFrom-Yaml | ConvertTo-Json -Depth 10; $yaml | Out-File -FilePath '" & tempJson & "' -Encoding UTF8"""
    
    shell.Run psCommand, 0, True
    
    ' JSONを読み込んでオブジェクト化
    Dim jsonContent As String
    Set ts = fso.OpenTextFile(tempJson, 1, False, -1)
    jsonContent = ts.ReadAll()
    ts.Close
    
    ' クリーンアップ
    If fso.FileExists(tempYaml) Then fso.DeleteFile tempYaml
    If fso.FileExists(tempJson) Then fso.DeleteFile tempJson
    
    ' JSONScriptingライブラリを使用してパース
    Dim json As Object
    Set json = JsonConverter.ParseJson(jsonContent)
    
    Set ParseYamlContent = json
End Function

Public Sub GenerateCodeFromYaml()
    Dim configSheet As Worksheet
    Set configSheet = ThisWorkbook.Worksheets("Config")
    
    Dim yamlPath As String
    yamlPath = Trim$(configSheet.Range("YamlPath").Value)
    
    If Len(yamlPath) = 0 Then
        MsgBox "YAML path (Config!YamlPath) is empty.", vbExclamation
        Exit Sub
    End If
    
    ' 相対パスを解決
    yamlPath = ResolveFilePath(yamlPath)
    
    If Not FileExists(yamlPath) Then
        MsgBox "YAML file not found: " & yamlPath, vbExclamation
        Exit Sub
    End If
    
    ' YAMLファイルを読み込む（PowerShell経由）
    Dim yamlData As Object
    Set yamlData = LoadYamlViaPowerShell(yamlPath)
    
    If yamlData Is Nothing Then
        MsgBox "Failed to load YAML file.", vbCritical
        Exit Sub
    End If
    
    ' コード生成
    Dim lines As Collection
    Set lines = New Collection
    
    GenerateCodeFromYamlData yamlData, lines
    
    ' 出力ファイルに書き込み
    Dim outputPath As String
    outputPath = yamlData("metadata")("outputFileName")
    outputPath = ResolveFilePath(outputPath)
    
    WriteLines outputPath, lines
    
    If Application.Interactive Then
        MsgBox "Generated: " & outputPath, vbInformation
    Else
        Application.StatusBar = "Generated: " & outputPath
    End If
End Sub

Private Function LoadYamlViaPowerShell(yamlPath As String) As Object
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    
    Dim tempJson As String
    tempJson = Environ("TEMP") & "\yaml_to_json_" & Format(Now, "yyyymmddhhnnss") & ".json"
    
    ' PowerShellスクリプトを作成
    Dim psScript As String
    psScript = "try { " & _
               "if (-not (Get-Module -ListAvailable -Name powershell-yaml)) { " & _
               "  Install-Module -Name powershell-yaml -Force -Scope CurrentUser -AllowClobber; " & _
               "} " & _
               "Import-Module powershell-yaml; " & _
               "$yaml = Get-Content -Raw -Path '" & yamlPath & "' | ConvertFrom-Yaml; " & _
               "$yaml | ConvertTo-Json -Depth 20 | Out-File -FilePath '" & tempJson & "' -Encoding UTF8; " & _
               "} catch { Write-Error $_.Exception.Message; exit 1; }"
    
    Dim psCommand As String
    psCommand = "powershell -NoProfile -ExecutionPolicy Bypass -Command """ & psScript & """"
    
    Dim result As Long
    result = shell.Run(psCommand, 0, True)
    
    If result <> 0 Then
        MsgBox "Failed to convert YAML to JSON via PowerShell.", vbCritical
        Set LoadYamlViaPowerShell = Nothing
        Exit Function
    End If
    
    ' JSONファイルを読み込む
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(tempJson) Then
        MsgBox "JSON file not created: " & tempJson, vbCritical
        Set LoadYamlViaPowerShell = Nothing
        Exit Function
    End If
    
    Dim ts As Object
    Set ts = fso.OpenTextFile(tempJson, 1, False, -1)
    Dim jsonContent As String
    jsonContent = ts.ReadAll()
    ts.Close
    
    ' クリーンアップ
    fso.DeleteFile tempJson
    
    ' JSONをパース
    On Error Resume Next
    Dim json As Object
    Set json = JsonConverter.ParseJson(jsonContent)
    
    If Err.Number <> 0 Then
        MsgBox "Failed to parse JSON: " & Err.Description, vbCritical
        Set LoadYamlViaPowerShell = Nothing
        Exit Function
    End If
    On Error GoTo 0
    
    Set LoadYamlViaPowerShell = json
End Function

Private Sub GenerateCodeFromYamlData(yamlData As Object, lines As Collection)
    Dim metadata As Object
    Set metadata = yamlData("metadata")
    
    ' References
    If yamlData.Exists("references") Then
        Dim refs As Object
        Set refs = yamlData("references")
        Dim i As Long
        For i = 1 To refs.Count
            lines.Add "//<ref>" & refs(i)("assembly") & "</ref>"
        Next i
        lines.Add ""
    End If
    
    ' Usings
    If yamlData.Exists("usings") Then
        Dim usings As Object
        Set usings = yamlData("usings")
        For i = 1 To usings.Count
            Dim category As Object
            Set category = usings(i)
            Dim namespaces As Object
            Set namespaces = category("namespaces")
            Dim j As Long
            For j = 1 To namespaces.Count
                lines.Add "using " & namespaces(j) & ";"
            Next j
            If i < usings.Count Then lines.Add ""
        Next i
        lines.Add ""
    End If
    
    ' Namespace header
    lines.Add "namespace " & metadata("namespace")
    lines.Add "{"
    lines.Add INDENT1 & "public class " & metadata("className") & " : " & metadata("baseClass")
    lines.Add INDENT1 & "{"
    lines.Add INDENT2 & "/// <summary>"
    lines.Add INDENT2 & "/// グローバル変数"
    lines.Add INDENT2 & "/// </summary>"
    
    ' Global variables
    If yamlData.Exists("globalVariables") Then
        Dim globals As Object
        Set globals = yamlData("globalVariables")
        For i = 1 To globals.Count
            Dim gVar As Object
            Set gVar = globals(i)
            Dim line As String
            line = INDENT2 & "string " & gVar("name") & " = """ & gVar("value") & """;"
            If gVar.Exists("description") Then
                line = line & "  // " & gVar("description")
            End If
            lines.Add line
        Next i
        lines.Add ""
        lines.Add ""
    End If
    
    ' Item class
    If yamlData.Exists("itemClass") Then
        GenerateItemClass yamlData("itemClass"), lines
    End If
    
    ' Methods
    If yamlData.Exists("methods") Then
        GenerateMethods yamlData("methods"), metadata, lines
    End If
    
    ' Common processing (from file)
    If yamlData.Exists("commonProcessing") Then
        Dim commonProc As Object
        Set commonProc = yamlData("commonProcessing")
        Dim sourceFile As String
        sourceFile = ResolveFilePath(commonProc("source"))
        
        If FileExists(sourceFile) Then
            Dim startLine As Long, endLine As Long
            startLine = commonProc("startLine")
            endLine = commonProc("endLine")
            
            Dim tailLines As Collection
            Set tailLines = ReadLinesFromFile(sourceFile, startLine, endLine)
            
            Dim line As Variant
            For Each line In tailLines
                lines.Add line
            Next line
        End If
    End If
End Sub

Private Sub GenerateItemClass(itemClass As Object, lines As Collection)
    lines.Add INDENT2 & "/// <summary>"
    lines.Add INDENT2 & "/// " & itemClass("description")
    lines.Add INDENT2 & "/// </summary>"
    lines.Add INDENT2 & "public class " & itemClass("name")
    lines.Add INDENT2 & "{"
    
    Dim properties As Object
    Set properties = itemClass("properties")
    
    Dim i As Long
    For i = 1 To properties.Count
        Dim propCategory As Object
        Set propCategory = properties(i)
        
        If propCategory.Exists("categoryComment") Then
            lines.Add INDENT3 & "// " & propCategory("categoryComment")
        End If
        
        Dim items As Object
        Set items = propCategory("items")
        
        Dim j As Long
        For j = 1 To items.Count
            Dim prop As Object
            Set prop = items(j)
            
            lines.Add INDENT3 & "[JsonProperty(""" & prop("jsonName") & """)]"
            lines.Add INDENT3 & "public " & prop("type") & " " & prop("propertyName") & " { get; set; }"
            If j < items.Count Then lines.Add ""
        Next j
        
        If i < properties.Count Then lines.Add ""
    Next i
    
    lines.Add ""
    lines.Add INDENT2 & "}"
    lines.Add ""
End Sub

Private Sub GenerateMethods(methods As Object, metadata As Object, lines As Collection)
    ' callAPI method
    If methods.Exists("callAPI") Then
        Dim callAPI As Object
        Set callAPI = methods("callAPI")
        
        lines.Add INDENT2 & "/// <summary>"
        lines.Add INDENT2 & "/// " & callAPI("description")
        lines.Add INDENT2 & "/// </summary>"
        
        Dim params As Object
        Set params = callAPI("parameters")
        Dim i As Long
        For i = 1 To params.Count
            Dim param As Object
            Set param = params(i)
            lines.Add INDENT2 & "/// <param name=""" & param("name") & """>" & param("description") & "</param>"
        Next i
        
        lines.Add INDENT2 & "public void callAPI(int mode){"
        lines.Add INDENT3 & "callAPIwithTarget(mode,"""");"
        lines.Add INDENT2 & "}"
        lines.Add INDENT2 & ""
    End If
    
    ' callAPIwithTarget method
    If methods.Exists("callAPIwithTarget") Then
        GenerateCallAPIWithTarget methods("callAPIwithTarget"), metadata, lines
    End If
End Sub

Private Sub GenerateCallAPIWithTarget(method As Object, metadata As Object, lines As Collection)
    lines.Add INDENT2 & "/// <summary>"
    lines.Add INDENT2 & "/// " & method("description")
    lines.Add INDENT2 & "/// </summary>"
    
    Dim params As Object
    Set params = method("parameters")
    Dim i As Long
    For i = 1 To params.Count
        Dim param As Object
        Set param = params(i)
        lines.Add INDENT2 & "/// <param name=""" & param("name") & """>" & param("description") & "</param>"
    Next i
    
    lines.Add INDENT2 & "public void callAPIwithTarget(int mode,string target){"
    lines.Add ""
    lines.Add INDENT3 & "List<Property> propertiesList = new List<Property>();"
    lines.Add ""
    lines.Add INDENT3 & "//使用するプロパティのデータリストを作成"
    lines.Add INDENT3 & "if(mode == 0){"
    
    ' Mode 0 properties
    If method.Exists("propertyListMode0") Then
        Dim props As Object
        Set props = method("propertyListMode0")
        For i = 1 To props.Count
            Dim prop As Object
            Set prop = props(i)
            Dim modifiedVal As String
            modifiedVal = IIf(prop("modified"), "true", "false")
            lines.Add INDENT4 & "propertiesList.Add(new Property { Name = """ & prop("name") & """,Value = " & prop("value") & ", Modified = " & modifiedVal & " });"
        Next i
    End If
    
    lines.Add ""
    lines.Add INDENT3 & "}"
    lines.Add INDENT3 & "//データ更新用、get時不要"
    lines.Add INDENT3 & "else{"
    
    ' Else properties
    If method.Exists("propertyListElse") Then
        Set props = method("propertyListElse")
        For i = 1 To props.Count
            Set prop = props(i)
            modifiedVal = IIf(prop("modified"), "true", "false")
            lines.Add INDENT4 & "propertiesList.Add(new Property { Name = """ & prop("name") & """, Value = " & prop("value") & ", Modified = " & modifiedVal & "});"
        Next i
    End If
    
    lines.Add INDENT4 & "//データ特定用、insert時は不要"
    
    ' Key properties
    If method.Exists("propertyListKeys") Then
        Dim keys As Object
        Set keys = method("propertyListKeys")
        lines.Add INDENT4 & "if(" & keys("condition") & "){"
        
        Dim keyItems As Object
        Set keyItems = keys("items")
        For i = 1 To keyItems.Count
            Set prop = keyItems(i)
            modifiedVal = IIf(prop("modified"), "true", "false")
            lines.Add INDENT4 & "   propertiesList.Add(new Property { Name = """ & prop("name") & """, Value = " & prop("value") & ", Modified = " & modifiedVal & " });"
        Next i
        
        lines.Add INDENT4 & "}        "
    End If
    
    lines.Add INDENT3 & "}"
    lines.Add ""
    lines.Add INDENT3 & "//データを取得"
    lines.Add INDENT3 & "string JSONResponse = getData(mode,gIDOName,propertiesList,target);"
    lines.Add INDENT3 & ""
    lines.Add INDENT3 & "//取得した結果はJSONなので、bodyを処理しデータを取得"
    lines.Add INDENT3 & "List<cItem> itemsList = JsonConvert.DeserializeObject<getJSONObject>(JSONResponse).Items;"
    lines.Add ""
    lines.Add INDENT3 & "//処理結果を変数に格納"
    lines.Add INDENT3 & "ThisForm.Variables(""" & metadata("resultVariableName") & """).Value = GenerateWebSetJson(itemsList);"
    lines.Add ""
    
    ' Grid display
    If method.Exists("gridDisplay") Then
        GenerateGridDisplay method("gridDisplay"), lines
    End If
    
    lines.Add INDENT2 & "}"
    lines.Add ""
End Sub

Private Sub GenerateGridDisplay(gridDisplay As Object, lines As Collection)
    lines.Add INDENT3 & "//グリッドの初期化"
    lines.Add INDENT3 & "int count = 1;"
    lines.Add INDENT3 & "if(ThisForm.Components[""" & gridDisplay("gridName") & """].GetGridRowCount() > 0 && " & gridDisplay("deleteCondition") & "){"
    lines.Add INDENT4 & "ThisForm.Components[""" & gridDisplay("gridName") & """].DeleteGridRows(1,ThisForm.Components[""" & gridDisplay("gridName") & """].GetGridRowCount());"
    lines.Add INDENT3 & "}"
    lines.Add ""
    lines.Add INDENT3 & "//データを表示"
    lines.Add INDENT3 & "foreach (var item in itemsList)"
    lines.Add INDENT3 & "{"
    lines.Add INDENT4 & "ThisForm.Components[""" & gridDisplay("gridName") & """].InsertGridRows(count,1);"
    
    Dim mappings As Object
    Set mappings = gridDisplay("mappings")
    Dim i As Long
    For i = 1 To mappings.Count
        Dim mapping As Object
        Set mapping = mappings(i)
        
        If Not mapping.Exists("commented") Or Not mapping("commented") Then
            Dim line As String
            If mapping.Exists("conversion") And Len(mapping("conversion")) > 0 Then
                line = INDENT4 & "ThisForm.Components[""" & gridDisplay("gridName") & """].SetGridValue(count," & mapping("column") & "," & mapping("conversion") & ");"
            Else
                line = INDENT4 & "ThisForm.Components[""" & gridDisplay("gridName") & """].SetGridValue(count," & mapping("column") & ",item." & mapping("property") & ");"
            End If
            
            If mapping.Exists("comment") And Len(mapping("comment")) > 0 Then
                line = line & "//" & mapping("comment")
            End If
            
            lines.Add line
        Else
            lines.Add INDENT4 & "//ThisForm.Components[""" & gridDisplay("gridName") & """].SetGridValue(count," & mapping("column") & ",item." & mapping("property") & ");"
        End If
    Next i
    
    lines.Add INDENT4 & ""
    lines.Add INDENT4 & "count++;"
    lines.Add INDENT3 & "} "
End Sub

Private Function ResolveFilePath(filePath As String) As String
    If Len(filePath) >= 2 Then
        If Mid$(filePath, 2, 1) = ":" Or Left$(filePath, 2) = "\\" Then
            ResolveFilePath = filePath
            Exit Function
        End If
    End If
    ResolveFilePath = ThisWorkbook.Path & Application.PathSeparator & filePath
End Function

Private Function FileExists(filePath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(filePath)
End Function

Private Sub WriteLines(targetPath As String, lines As Collection)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim stream As Object
    Set stream = fso.CreateTextFile(targetPath, True, False)
    Dim i As Long
    For i = 1 To lines.Count
        stream.WriteLine lines(i)
    Next i
    stream.Close
End Sub

Private Function ReadLinesFromFile(filePath As String, startLine As Long, endLine As Long) As Collection
    Dim lines As New Collection
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim ts As Object
    Set ts = fso.OpenTextFile(filePath, 1, False, 0)
    
    Dim currentLine As Long
    currentLine = 1
    
    Do While Not ts.AtEndOfStream
        Dim line As String
        line = ts.ReadLine
        
        If currentLine >= startLine And currentLine <= endLine Then
            lines.Add line
        End If
        
        currentLine = currentLine + 1
        
        If currentLine > endLine Then Exit Do
    Loop
    
    ts.Close
    Set ReadLinesFromFile = lines
End Function
'@

Write-Host "Creating ItemPikingDetailGeneratorYML.xlsm..." -ForegroundColor Cyan

$generatorPath = Join-Path (Get-Location) 'ItemPikingDetailGeneratorYML.xlsm'
if(Test-Path $generatorPath){ Remove-Item $generatorPath -Force }

$excel = $null
$wb = $null
$configSheet = $null
$readmeSheet = $null
$module = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    # マクロ有効ブックとして作成
    $wb = $excel.Workbooks.Add()
    
    Write-Host "Creating Config sheet..." -ForegroundColor Cyan
    $configSheet = $wb.Worksheets.Item(1)
    $configSheet.Name = 'Config'

    $configSheet.Cells.Item(1,1).Value2 = 'YamlPath'
    $configSheet.Cells.Item(1,2).Value2 = 'ItemPiking.YAML'
    $configSheet.Cells.Item(1,3).Value2 = 'YAML definition file (relative or absolute path)'
    $configSheet.Range('B1').Name = 'YamlPath'
    
    $configSheet.Cells.Item(3,1).Value2 = '【設定】'
    $configSheet.Cells.Item(4,1).Value2 = 'YamlPath: コード生成定義YAMLファイルのパス'
    $configSheet.Cells.Item(5,1).Value2 = '相対パスの場合、このブックと同じフォルダからの相対パス'
    $configSheet.Cells.Item(7,1).Value2 = '【使い方】'
    $configSheet.Cells.Item(8,1).Value2 = '1. YamlPathに正しいYAMLファイルパスを設定'
    $configSheet.Cells.Item(9,1).Value2 = '2. 開発タブ > マクロ > GenerateCodeFromYaml を実行'
    $configSheet.Cells.Item(10,1).Value2 = '3. YAMLで指定されたファイル名でコードが生成されます'
    
    $configSheet.Cells.Item(12,1).Value2 = '【必要な環境】'
    $configSheet.Cells.Item(13,1).Value2 = '・PowerShell 5.1以上'
    $configSheet.Cells.Item(14,1).Value2 = '・powershell-yaml モジュール（初回実行時に自動インストール）'
    $configSheet.Cells.Item(15,1).Value2 = '  Install-Module -Name powershell-yaml -Scope CurrentUser'
    
    $configSheet.Columns('A:C').AutoFit()

    Write-Host "Creating Readme sheet..." -ForegroundColor Cyan
    $readmeSheet = $wb.Worksheets.Add($configSheet)
    $readmeSheet.Name = 'Readme'
    
    $readmeSheet.Cells.Item(1,1).Value2 = 'ItemPikingDetailGeneratorYML - YAMLベースのコード生成ツール'
    $readmeSheet.Cells.Item(3,1).Value2 = '【概要】'
    $readmeSheet.Cells.Item(4,1).Value2 = 'YAML定義ファイルからItemPicking.csを生成するExcelマクロツール'
    $readmeSheet.Cells.Item(6,1).Value2 = '【ワークフロー】'
    $readmeSheet.Cells.Item(7,1).Value2 = '1. ItemPiking.md (設計書) から ItemPiking.YAML を作成'
    $readmeSheet.Cells.Item(8,1).Value2 = '2. このExcelマクロでYAMLからItemPicking.csを生成'
    $readmeSheet.Cells.Item(10,1).Value2 = '【YAMLファイルの構造】'
    $readmeSheet.Cells.Item(11,1).Value2 = 'metadata: 基本情報（クラス名、名前空間など）'
    $readmeSheet.Cells.Item(12,1).Value2 = 'references: DLL参照'
    $readmeSheet.Cells.Item(13,1).Value2 = 'usings: using文'
    $readmeSheet.Cells.Item(14,1).Value2 = 'globalVariables: グローバル変数定義'
    $readmeSheet.Cells.Item(15,1).Value2 = 'itemClass: cItemクラス定義'
    $readmeSheet.Cells.Item(16,1).Value2 = 'methods: メソッド定義（callAPI, callAPIwithTarget）'
    $readmeSheet.Cells.Item(17,1).Value2 = 'commonProcessing: 共通処理（152行目以降）'
    $readmeSheet.Cells.Item(19,1).Value2 = '【利点】'
    $readmeSheet.Cells.Item(20,1).Value2 = '・YAMLで業務ロジックを宣言的に定義可能'
    $readmeSheet.Cells.Item(21,1).Value2 = '・バージョン管理しやすい（テキストファイル）'
    $readmeSheet.Cells.Item(22,1).Value2 = '・他のプロジェクトへの適用が容易'
    $readmeSheet.Cells.Item(23,1).Value2 = '・設計書から自動生成可能'
    
    $readmeSheet.Columns('A').ColumnWidth = 80

    Write-Host "Adding VBA module with JsonConverter..." -ForegroundColor Cyan
    
    # JsonConverter.basをVBAProjectに追加
    # 注: JsonConverter.basは別途必要（VBA-JSONライブラリ）
    $module = $wb.VBProject.VBComponents.Add(1)
    $module.Name = 'CodeGenerator'
    $module.CodeModule.AddFromString($macroCode)

    Write-Host "Saving as macro-enabled workbook..." -ForegroundColor Cyan
    # 52 = xlOpenXMLWorkbookMacroEnabled (.xlsm)
    $wb.SaveAs($generatorPath, 52)
    $wb.Close($true)
    
    Write-Host "Successfully created: $generatorPath" -ForegroundColor Green
    Write-Host ""
    Write-Host "Note: JsonConverter.bas (VBA-JSON library) needs to be added manually." -ForegroundColor Yellow
    Write-Host "Download from: https://github.com/VBA-tools/VBA-JSON" -ForegroundColor Yellow
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
    throw
}
finally {
    if($module){ [System.Runtime.InteropServices.Marshal]::ReleaseComObject($module) | Out-Null }
    if($readmeSheet){ [System.Runtime.InteropServices.Marshal]::ReleaseComObject($readmeSheet) | Out-Null }
    if($configSheet){ [System.Runtime.InteropServices.Marshal]::ReleaseComObject($configSheet) | Out-Null }
    if($wb){ [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    if($excel){ $excel.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

Write-Host ""
Write-Host "Next steps:" -ForegroundColor Cyan
Write-Host "1. Install powershell-yaml module if not already installed:" -ForegroundColor White
Write-Host "   Install-Module -Name powershell-yaml -Scope CurrentUser" -ForegroundColor Gray
Write-Host "2. Open ItemPikingDetailGeneratorYML.xlsm" -ForegroundColor White
Write-Host "3. Run macro: GenerateCodeFromYaml" -ForegroundColor White
