function Add-ExcelTable {
    param(
        [Parameter(Mandatory=$true)]$Sheet,
        [Parameter(Mandatory=$true)]$StartRow,
        [Parameter(Mandatory=$true)]$Columns,
        [Parameter(Mandatory=$true)]$Data,
        [Parameter(Mandatory=$true)]$Name,
        [int]$StartColumn = 1
    )

    $headerRow = $StartRow
    for($colIndex = 0; $colIndex -lt $Columns.Count; $colIndex++){
        $columnDef = $Columns[$colIndex]
        if($columnDef -is [System.Collections.IDictionary]){
            $header = $columnDef['Header']
        } else {
            $header = $columnDef.Header
        }
        if($header -isnot [string]){
            $header = $header.ToString()
        }
        try {
            $Sheet.Cells.Item($headerRow, $StartColumn + $colIndex).Value2 = $header
        } catch {
            Write-Host "Failed to set header for column $colIndex in table $Name (value=$header)" -ForegroundColor Yellow
            throw
        }
    }

    $row = $headerRow + 1
    foreach($item in $Data){
        for($colIndex = 0; $colIndex -lt $Columns.Count; $colIndex++){
            $columnDef = $Columns[$colIndex]
            if($columnDef -is [System.Collections.IDictionary]){
                $keyName = $columnDef['Key']
            } else {
                $keyName = $columnDef.Key
            }

            if($item -is [System.Collections.IDictionary]){
                $value = $item[$keyName]
            } else {
                $prop = $item.PSObject.Properties[$keyName]
                $value = if($prop){ $prop.Value } else { $null }
            }
            if($null -eq $value){ $value = "" }
            if($value -isnot [string]){
                $value = $value.ToString()
            }
            try {
                $Sheet.Cells.Item($row, $StartColumn + $colIndex).Value2 = $value
            } catch {
                Write-Host "Failed to set value for table $Name row $row column $colIndex (key=$keyName value=$value)" -ForegroundColor Yellow
                throw
            }
        }
        $row++
    }

    $endRow = $row - 1
    $endColumn = $StartColumn + $Columns.Count - 1
    $range = $Sheet.Range($Sheet.Cells.Item($headerRow, $StartColumn), $Sheet.Cells.Item($endRow, $endColumn))
    $table = $Sheet.ListObjects.Add(1, $range, $null, 1)
    $table.Name = $Name
    $table.TableStyle = "TableStyleLight9"
    return $endRow + 2
}

$macroCode = @'
Option Explicit

Private Const INDENT1 As String = "   "
Private Const INDENT2 As String = "      "
Private Const INDENT3 As String = "         "
Private Const INDENT4 As String = "            "

Public Sub GenerateItemPicking()
    Dim spec As Worksheet
    Dim tail As Worksheet
    Set spec = ThisWorkbook.Worksheets("SpecData")
    Set tail = ThisWorkbook.Worksheets("CommonTail")
    
    Dim outputPath As String
    outputPath = Trim$(spec.Range("OutputPath").Value)
    If Len(outputPath) = 0 Then
        MsgBox "Output path (SpecData!OutputPath) is empty.", vbExclamation
        Exit Sub
    End If
    
    Dim lines As Collection
    Set lines = New Collection
    
    ' 業務固有処理（1-151行目）
    AppendReference lines, spec.Range("ReferenceLine").Value
    AppendUsings lines, spec.ListObjects("Usings")
    AppendNamespaceHeader lines, spec
    AppendGlobalConfig lines, spec.ListObjects("GlobalConfig")
    AppendItemClass lines, spec.ListObjects("ItemProperties")
    AppendCallAPI lines, spec
    AppendCallAPIWithTarget lines, spec
    
    ' 共通処理（152行目以降）
    AppendTail lines, tail.ListObjects("TailLines")
    
    Dim targetPath As String
    targetPath = ResolveOutputPath(outputPath)
    WriteLines targetPath, lines
    
    If Application.Interactive Then
        MsgBox "Generated: " & targetPath, vbInformation
    Else
        Application.StatusBar = "Generated: " & targetPath
    End If
End Sub

Private Sub AppendReference(lines As Collection, refLine As String)
    If Len(Trim$(refLine)) > 0 Then
        lines.Add refLine
    End If
    lines.Add ""
End Sub

Private Sub AppendUsings(lines As Collection, tbl As ListObject)
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    Dim data As Variant
    data = tbl.DataBodyRange.Value
    Dim currentCategory As String
    currentCategory = vbNullString
    Dim i As Long
    For i = 1 To TableRowCount(data)
        If currentCategory <> vbNullString And TableValue(data, i, 3) <> currentCategory Then
            lines.Add ""
        End If
        lines.Add "using " & TableValue(data, i, 2) & ";"
        currentCategory = TableValue(data, i, 3)
    Next i
    lines.Add ""
End Sub

Private Sub AppendNamespaceHeader(lines As Collection, spec As Worksheet)
    lines.Add "namespace " & spec.Range("NamespaceName").Value
    lines.Add "{"
    lines.Add INDENT1 & "public class " & spec.Range("ClassName").Value & " : FormScript"
    lines.Add INDENT1 & "{"
    lines.Add INDENT2 & "/// <summary>"
    lines.Add INDENT2 & "/// グローバル変数"
    lines.Add INDENT2 & "/// </summary>"
End Sub

Private Sub AppendGlobalConfig(lines As Collection, tbl As ListObject)
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    Dim data As Variant
    data = tbl.DataBodyRange.Value
    Dim i As Long
    For i = 1 To TableRowCount(data)
        Dim line As String
        line = INDENT2 & "string " & TableValue(data, i, 1) & " = " & QuoteLiteral(TableValue(data, i, 2)) & ";"
        If Len(Trim$(TableValue(data, i, 3))) > 0 Then
            line = line & "  // " & TableValue(data, i, 3)
        End If
        lines.Add line
    Next i
    lines.Add ""
    lines.Add ""
End Sub

Private Sub AppendItemClass(lines As Collection, tbl As ListObject)
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    Dim data As Variant
    data = tbl.DataBodyRange.Value
    lines.Add INDENT2 & "/// <summary>"
    lines.Add INDENT2 & "/// 取得データ、更新データ作成のクラス"
    lines.Add INDENT2 & "/// </summary>"
    lines.Add INDENT2 & "public class cItem"
    lines.Add INDENT2 & "{"
    
    Dim lastCategory As String
    lastCategory = ""
    Dim i As Long
    For i = 1 To TableRowCount(data)
        Dim currentDesc As String
        currentDesc = Trim$(TableValue(data, i, 4))
        
        ' カテゴリコメントの処理
        If i = 1 Then
            lines.Add INDENT3 & "// 取得予定のプロパティ"
        ElseIf InStr(currentDesc, "更新") > 0 Or InStr(currentDesc, "IDO") > 0 Or InStr(currentDesc, "自動生成") > 0 Then
            If lastCategory <> "internal" Then
                lines.Add ""
                lines.Add INDENT3 & "// 更新時必要のプロパティ（基本IDO作成時自動生成される）"
                lastCategory = "internal"
            End If
        End If
        
        Dim attributeLine As String
        attributeLine = INDENT3 & "[JsonProperty(""" & TableValue(data, i, 1) & """)]"
        lines.Add attributeLine
        lines.Add INDENT3 & "public " & TableValue(data, i, 3) & " " & TableValue(data, i, 2) & " { get; set; }"
        If i < TableRowCount(data) Then
            lines.Add ""
        End If
    Next i
    lines.Add ""
    lines.Add INDENT2 & "}"
    lines.Add ""
End Sub

Private Sub AppendCallAPI(lines As Collection, spec As Worksheet)
    lines.Add INDENT2 & "/// <summary>"
    lines.Add INDENT2 & "/// APIを呼び出してデータ取得・グリッド表示を行う"
    lines.Add INDENT2 & "/// </summary>"
    lines.Add INDENT2 & "/// <param name=""mode"">操作モード（0:取得, 1:挿入, 2:更新, 4:削除）</param>"
    lines.Add INDENT2 & "public void callAPI(int mode){"
    lines.Add INDENT3 & "callAPIwithTarget(mode,"""");"
    lines.Add INDENT2 & "}"
    lines.Add INDENT2 & ""
End Sub

Private Sub AppendCallAPIWithTarget(lines As Collection, spec As Worksheet)
    lines.Add INDENT2 & "/// <summary>"
    lines.Add INDENT2 & "/// APIを呼び出してデータ取得・グリッド表示を行う"
    lines.Add INDENT2 & "/// </summary>"
    lines.Add INDENT2 & "/// <param name=""mode"">操作モード（0:取得, 1:挿入, 2:更新, 4:削除）</param>"
    lines.Add INDENT2 & "/// <param name=""target"">更新目標（今はLot）</param>"
    lines.Add INDENT2 & "public void callAPIwithTarget(int mode,string target){"
    lines.Add ""
    lines.Add INDENT3 & "List<Property> propertiesList = new List<Property>();"
    lines.Add ""
    lines.Add INDENT3 & "//使用するプロパティのデータリストを作成"
    lines.Add INDENT3 & "if(mode == 0){"
    
    ' Mode 0のプロパティリストを追加
    Dim tbl As ListObject
    Set tbl = spec.ListObjects("CallAPIPropertiesMode0")
    If Not tbl.DataBodyRange Is Nothing Then
        Dim data As Variant
        data = tbl.DataBodyRange.Value
        Dim i As Long
        For i = 1 To TableRowCount(data)
            Dim entry As String
            entry = INDENT4 & "propertiesList.Add(new Property { Name = """ & TableValue(data, i, 1) & """,Value = " & TableValue(data, i, 2) & ", Modified = " & TableValue(data, i, 3) & " });"
            If Len(Trim$(TableValue(data, i, 4))) > 0 Then
                entry = entry & " //" & TableValue(data, i, 4)
            End If
            lines.Add entry
        Next i
    End If
    
    lines.Add ""
    lines.Add INDENT3 & "}"
    lines.Add INDENT3 & "//データ更新用、get時不要"
    lines.Add INDENT3 & "else{"
    
    ' Mode else(更新)のプロパティリストを追加
    Set tbl = spec.ListObjects("CallAPIPropertiesElse")
    If Not tbl.DataBodyRange Is Nothing Then
        data = tbl.DataBodyRange.Value
        For i = 1 To TableRowCount(data)
            entry = INDENT4 & "propertiesList.Add(new Property { Name = """ & TableValue(data, i, 1) & """, Value = " & TableValue(data, i, 2) & ", Modified = " & TableValue(data, i, 3) & "});"
            If Len(Trim$(TableValue(data, i, 4))) > 0 Then
                entry = entry & " //" & TableValue(data, i, 4)
            End If
            lines.Add entry
        Next i
    End If
    
    lines.Add INDENT4 & "//データ特定用、insert時は不要"
    lines.Add INDENT4 & "if(mode != 1){"
    
    ' データ特定用プロパティ
    Set tbl = spec.ListObjects("CallAPIPropertiesKey")
    If Not tbl.DataBodyRange Is Nothing Then
        data = tbl.DataBodyRange.Value
        For i = 1 To TableRowCount(data)
            entry = INDENT4 & "   propertiesList.Add(new Property { Name = """ & TableValue(data, i, 1) & """, Value = " & TableValue(data, i, 2) & ", Modified = " & TableValue(data, i, 3) & " });"
            If Len(Trim$(TableValue(data, i, 4))) > 0 Then
                entry = entry & " //" & TableValue(data, i, 4)
            End If
            lines.Add entry
        Next i
    End If
    
    lines.Add INDENT4 & "}        "
    lines.Add INDENT3 & "}"
    lines.Add ""
    lines.Add INDENT3 & "//データを取得"
    lines.Add INDENT3 & "string JSONResponse = getData(mode,gIDOName,propertiesList,target);"
    lines.Add INDENT3 & ""
    lines.Add INDENT3 & "//取得した結果はJSONなので、bodyを処理しデータを取得"
    lines.Add INDENT3 & "List<cItem> itemsList = JsonConvert.DeserializeObject<getJSONObject>(JSONResponse).Items;"
    lines.Add ""
    lines.Add INDENT3 & "//処理結果を変数に格納"
    lines.Add INDENT3 & "ThisForm.Variables(""" & spec.Range("ResultVariableName").Value & """).Value = GenerateWebSetJson(itemsList);"
    lines.Add ""
    lines.Add INDENT3 & "//グリッドの初期化"
    lines.Add INDENT3 & "int count = 1;"
    lines.Add INDENT3 & "if(ThisForm.Components[""ResultGrid""].GetGridRowCount() > 0 && target == """"){"
    lines.Add INDENT4 & "ThisForm.Components[""ResultGrid""].DeleteGridRows(1,ThisForm.Components[""ResultGrid""].GetGridRowCount());"
    lines.Add INDENT3 & "}"
    lines.Add ""
    lines.Add INDENT3 & "//データを表示"
    lines.Add INDENT3 & "foreach (var item in itemsList)"
    lines.Add INDENT3 & "{"
    
    ' グリッド表示処理
    Set tbl = spec.ListObjects("GridDisplay")
    If Not tbl.DataBodyRange Is Nothing Then
        data = tbl.DataBodyRange.Value
        lines.Add INDENT4 & "ThisForm.Components[""ResultGrid""].InsertGridRows(count,1);"
        For i = 1 To TableRowCount(data)
            Dim colIdx As String
            colIdx = TableValue(data, i, 1)
            Dim propName As String
            propName = TableValue(data, i, 2)
            Dim conversion As String
            conversion = Trim$(TableValue(data, i, 3))
            Dim comment As String
            comment = Trim$(TableValue(data, i, 4))
            
            If Len(conversion) > 0 Then
                entry = INDENT4 & "ThisForm.Components[""ResultGrid""].SetGridValue(count," & colIdx & "," & conversion & ");"
            Else
                entry = INDENT4 & "ThisForm.Components[""ResultGrid""].SetGridValue(count," & colIdx & ",item." & propName & ");"
            End If
            If Len(comment) > 0 Then
                entry = entry & "//" & comment
            End If
            lines.Add entry
        Next i
    End If
    
    lines.Add INDENT4 & ""
    lines.Add INDENT4 & "count++;"
    lines.Add INDENT3 & "} "
    lines.Add INDENT2 & "}"
    lines.Add ""
End Sub

Private Sub AppendTail(lines As Collection, tbl As ListObject)
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    Dim data As Variant
    data = tbl.DataBodyRange.Value
    Dim i As Long
    For i = 1 To TableRowCount(data)
        lines.Add TableValue(data, i, 1)
    Next i
End Sub

Private Function ResolveOutputPath(outputPath As String) As String
    If Len(outputPath) >= 2 Then
        If Mid$(outputPath, 2, 1) = ":" Or Left$(outputPath, 2) = "\\" Then
            ResolveOutputPath = outputPath
            Exit Function
        End If
    End If
    ResolveOutputPath = ThisWorkbook.Path & Application.PathSeparator & outputPath
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

Private Function QuoteLiteral(value As String) As String
    QuoteLiteral = """" & Replace(value, """", """""") & """"
End Function

Private Function TableRowCount(data As Variant) As Long
    If IsArray(data) Then
        TableRowCount = UBound(data, 1)
    Else
        TableRowCount = 1
    End If
End Function

Private Function TableValue(data As Variant, rowIndex As Long, columnIndex As Long) As String
    If IsArray(data) Then
        TableValue = CStr(data(rowIndex, columnIndex))
    Else
        If rowIndex = 1 And columnIndex = 1 Then
            TableValue = CStr(data)
        Else
            TableValue = ""
        End If
    End If
End Function
'@

# ItemPicking用データ定義
$usings = @(
    [pscustomobject]@{Order=1; Namespace='System'; Category='base'},
    [pscustomobject]@{Order=2; Namespace='Mongoose.IDO.Protocol'; Category='mongoose'},
    [pscustomobject]@{Order=3; Namespace='Mongoose.Scripting'; Category='mongoose'},
    [pscustomobject]@{Order=4; Namespace='Mongoose.Core.Common'; Category='mongoose'},
    [pscustomobject]@{Order=5; Namespace='Mongoose.Core.DataAccess'; Category='mongoose'},
    [pscustomobject]@{Order=6; Namespace='System.Collections.Generic'; Category='framework'},
    [pscustomobject]@{Order=7; Namespace='System.IO'; Category='framework'},
    [pscustomobject]@{Order=8; Namespace='System.Linq'; Category='framework'},
    [pscustomobject]@{Order=9; Namespace='System.Net'; Category='framework'},
    [pscustomobject]@{Order=10; Namespace='System.Text'; Category='framework'},
    [pscustomobject]@{Order=11; Namespace='System.Web'; Category='framework'},
    [pscustomobject]@{Order=12; Namespace='Newtonsoft.Json'; Category='external'}
)

$globalConfig = @(
    [pscustomobject]@{Name='gIDOName'; Value='SLLotLocs'; Description='ターゲットIDO名'},
    [pscustomobject]@{Name='gSSO'; Value='1'; Description='SSOの使用、基本は1　1：はい　0：いいえ'},
    [pscustomobject]@{Name='gServerId'; Value='0'; Description='APIサーバのID、詳しくはAPIサーバ関連'},
    [pscustomobject]@{Name='gSuiteContext'; Value='CSI'; Description='APIスイートのコンテキスト、同上'},
    [pscustomobject]@{Name='gContentType'; Value='application/json'; Description='返り値のタイプ、基本設定不要もしくは"application/json"'},
    [pscustomobject]@{Name='gTimeout'; Value='10000'; Description='タイムアウト設定'},
    [pscustomobject]@{Name='gSiteName'; Value='Q72Q74BY8XUT3SKY_TRN_AJP'; Description='ターゲットIDOが存在するサイト名'}
)

$itemProperties = @(
    [pscustomobject]@{JsonName='Whse'; PropertyName='Whse'; Type='string'; Description='倉庫'},
    [pscustomobject]@{JsonName='Lot'; PropertyName='Lot'; Type='string'; Description='ロット'},
    [pscustomobject]@{JsonName='Loc'; PropertyName='Loc'; Type='string'; Description='ロケーション'},
    [pscustomobject]@{JsonName='QtyOnHand'; PropertyName='QtyOnHand'; Type='string'; Description='在庫数'},
    [pscustomobject]@{JsonName='PickingNum'; PropertyName='PickingNum'; Type='string'; Description='ピッキング数'},
    [pscustomobject]@{JsonName='Item'; PropertyName='Item'; Type='string'; Description='品目'},
    [pscustomobject]@{JsonName='ItemDescription'; PropertyName='ItemDescription'; Type='string'; Description='品目説明'},
    [pscustomobject]@{JsonName='RecordDate'; PropertyName='RecordDate'; Type='string'; Description='更新時必要のプロパティ（基本IDO作成時自動生成される）'},
    [pscustomobject]@{JsonName='RowPointer'; PropertyName='RowPointer'; Type='string'; Description='更新時必要のプロパティ（基本IDO作成時自動生成される）'},
    [pscustomobject]@{JsonName='_ItemId'; PropertyName='_ItemId'; Type='string'; Description='更新時必要のプロパティ（基本IDO作成時自動生成される）'}
)

$callApiPropertiesMode0 = @(
    [pscustomobject]@{Name='Whse'; InitialValue='""'; Modified='true'; Comment=''},
    [pscustomobject]@{Name='Lot'; InitialValue='""'; Modified='true'; Comment=''},
    [pscustomobject]@{Name='Loc'; InitialValue='""'; Modified='true'; Comment=''},
    [pscustomobject]@{Name='QtyOnHand'; InitialValue='""'; Modified='true'; Comment=''},
    [pscustomobject]@{Name='Item'; InitialValue='""'; Modified='true'; Comment=''},
    [pscustomobject]@{Name='ItemDescription'; InitialValue='""'; Modified='true'; Comment=''},
    [pscustomobject]@{Name='RecordDate'; InitialValue='""'; Modified='true'; Comment=''},
    [pscustomobject]@{Name='RowPointer'; InitialValue='""'; Modified='true'; Comment=''},
    [pscustomobject]@{Name='_ItemId'; InitialValue='""'; Modified='true'; Comment=''}
)

$callApiPropertiesElse = @(
    [pscustomobject]@{Name='Whse'; InitialValue='ThisForm.Components["ResultGrid"].GetGridValue(ThisForm.Components["ResultGrid"].GetGridCurrentRow(),1)'; Modified='true'; Comment=''},
    [pscustomobject]@{Name='Lot'; InitialValue='ThisForm.Components["ResultGrid"].GetGridValue(ThisForm.Components["ResultGrid"].GetGridCurrentRow(),2)'; Modified='true'; Comment=''},
    [pscustomobject]@{Name='Loc'; InitialValue='ThisForm.Components["ResultGrid"].GetGridValue(ThisForm.Components["ResultGrid"].GetGridCurrentRow(),3)'; Modified='true'; Comment=''},
    [pscustomobject]@{Name='QtyOnHand'; InitialValue='ThisForm.Components["ResultGrid"].GetGridValue(ThisForm.Components["ResultGrid"].GetGridCurrentRow(),4)'; Modified='true'; Comment=''},
    [pscustomobject]@{Name='Item'; InitialValue='ThisForm.Components["ResultGrid"].GetGridValue(ThisForm.Components["ResultGrid"].GetGridCurrentRow(),5)'; Modified='true'; Comment=''},
    [pscustomobject]@{Name='ItemDescription'; InitialValue='ThisForm.Variables("gItemDescription").Value'; Modified='true'; Comment=''}
)

$callApiPropertiesKey = @(
    [pscustomobject]@{Name='RecordDate'; InitialValue='ThisForm.Components["ResultGrid"].GetGridValue(ThisForm.Components["ResultGrid"].GetGridCurrentRow(),9)'; Modified='true'; Comment=''},
    [pscustomobject]@{Name='RowPointer'; InitialValue='ThisForm.Components["ResultGrid"].GetGridValue(ThisForm.Components["ResultGrid"].GetGridCurrentRow(),10)'; Modified='true'; Comment=''},
    [pscustomobject]@{Name='_ItemId'; InitialValue='ThisForm.Components["ResultGrid"].GetGridValue(ThisForm.Components["ResultGrid"].GetGridCurrentRow(),11)'; Modified='true'; Comment=''}
)

$gridDisplay = @(
    [pscustomobject]@{Column='1'; Property='Whse'; Conversion=''; Comment=''},
    [pscustomobject]@{Column='2'; Property='Lot'; Conversion=''; Comment=''},
    [pscustomobject]@{Column='3'; Property='Loc'; Conversion=''; Comment=''},
    [pscustomobject]@{Column='4'; Property='QtyOnHand'; Conversion='float.Parse(item.QtyOnHand).ToString()'; Comment='小数点消すためいったんfloatに変換'},
    [pscustomobject]@{Column='5'; Property='PickingNum'; Conversion=''; Comment='コメント済み'},
    [pscustomobject]@{Column='6'; Property='Item'; Conversion=''; Comment=''},
    [pscustomobject]@{Column='7'; Property='ItemDescription'; Conversion=''; Comment=''},
    [pscustomobject]@{Column='8'; Property='Match'; Conversion=''; Comment='コメント済み'},
    [pscustomobject]@{Column='9'; Property='RecordDate'; Conversion=''; Comment=''},
    [pscustomobject]@{Column='10'; Property='RowPointer'; Conversion=''; Comment=''},
    [pscustomobject]@{Column='11'; Property='_ItemId'; Conversion=''; Comment=''}
)

# ItemPicking.csの152行目以降を読み込み
$tailLines = Get-Content -Path 'ItemPicking.cs' -Encoding UTF8 | Select-Object -Skip 151

$generatorPath = Join-Path (Get-Location) 'ItemPikingDetailGenerator.xlsx'
if(Test-Path $generatorPath){ Remove-Item $generatorPath -Force }

$excel = $null
$wb = $null
$spec = $null
$tailSheet = $null
$module = $null
$readme = $null

try {
    Write-Host "Creating Excel workbook..." -ForegroundColor Cyan
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Add()
    $spec = $wb.Worksheets.Item(1)
    $spec.Name = 'SpecData'

    Write-Host "Setting up named ranges..." -ForegroundColor Cyan
    $spec.Cells.Item(1,1).Value2 = 'OutputPath'
    $spec.Cells.Item(1,2).Value2 = 'ItemPicking.cs'
    $spec.Cells.Item(1,3).Value2 = 'Output file (relative to workbook)'
    $spec.Range('B1').Name = 'OutputPath'

    $spec.Cells.Item(2,1).Value2 = 'ReferenceLine'
    $spec.Cells.Item(2,2).Value2 = '//<ref>Newtonsoft.Json.dll</ref>'
    $spec.Cells.Item(2,3).Value2 = 'Reference assembly directive'
    $spec.Range('B2').Name = 'ReferenceLine'

    $spec.Cells.Item(3,1).Value2 = 'NamespaceName'
    $spec.Cells.Item(3,2).Value2 = 'Mongoose.FormScripts'
    $spec.Cells.Item(3,3).Value2 = 'Namespace'
    $spec.Range('B3').Name = 'NamespaceName'

    $spec.Cells.Item(4,1).Value2 = 'ClassName'
    $spec.Cells.Item(4,2).Value2 = 'ItemPicking'
    $spec.Cells.Item(4,3).Value2 = 'FormScript class name'
    $spec.Range('B4').Name = 'ClassName'

    $spec.Cells.Item(5,1).Value2 = 'ResultVariableName'
    $spec.Cells.Item(5,2).Value2 = 'vJSONResult'
    $spec.Cells.Item(5,3).Value2 = 'WSForm variable for JSON payload'
    $spec.Range('B5').Name = 'ResultVariableName'

    Write-Host "Adding tables..." -ForegroundColor Cyan
    $row = 9
    $spec.Cells.Item($row,1).Value2 = 'Usings'
    $row += 1
    $columns = @(
        [pscustomobject]@{Header='Order'; Key='Order'},
        [pscustomobject]@{Header='Namespace'; Key='Namespace'},
        [pscustomobject]@{Header='Category'; Key='Category'}
    )
    $row = Add-ExcelTable -Sheet $spec -StartRow $row -Columns $columns -Data ($usings | Sort-Object Order) -Name 'Usings'

    $spec.Cells.Item($row,1).Value2 = 'Global Configuration'
    $row += 1
    $columns = @(
        [pscustomobject]@{Header='Name'; Key='Name'},
        [pscustomobject]@{Header='Value'; Key='Value'},
        [pscustomobject]@{Header='Description'; Key='Description'}
    )
    $row = Add-ExcelTable -Sheet $spec -StartRow $row -Columns $columns -Data $globalConfig -Name 'GlobalConfig'

    $spec.Cells.Item($row,1).Value2 = 'IDO Item Properties'
    $row += 1
    $columns = @(
        [pscustomobject]@{Header='JsonName'; Key='JsonName'},
        [pscustomobject]@{Header='PropertyName'; Key='PropertyName'},
        [pscustomobject]@{Header='Type'; Key='Type'},
        [pscustomobject]@{Header='Description'; Key='Description'}
    )
    $row = Add-ExcelTable -Sheet $spec -StartRow $row -Columns $columns -Data $itemProperties -Name 'ItemProperties'

    $spec.Cells.Item($row,1).Value2 = 'callAPI Property List (Mode=0)'
    $row += 1
    $columns = @(
        [pscustomobject]@{Header='Name'; Key='Name'},
        [pscustomobject]@{Header='InitialValue'; Key='InitialValue'},
        [pscustomobject]@{Header='Modified'; Key='Modified'},
        [pscustomobject]@{Header='Comment'; Key='Comment'}
    )
    $row = Add-ExcelTable -Sheet $spec -StartRow $row -Columns $columns -Data $callApiPropertiesMode0 -Name 'CallAPIPropertiesMode0'

    $spec.Cells.Item($row,1).Value2 = 'callAPI Property List (Else=Update)'
    $row += 1
    $columns = @(
        [pscustomobject]@{Header='Name'; Key='Name'},
        [pscustomobject]@{Header='InitialValue'; Key='InitialValue'},
        [pscustomobject]@{Header='Modified'; Key='Modified'},
        [pscustomobject]@{Header='Comment'; Key='Comment'}
    )
    $row = Add-ExcelTable -Sheet $spec -StartRow $row -Columns $columns -Data $callApiPropertiesElse -Name 'CallAPIPropertiesElse'

    $spec.Cells.Item($row,1).Value2 = 'callAPI Property List (Key Properties)'
    $row += 1
    $columns = @(
        [pscustomobject]@{Header='Name'; Key='Name'},
        [pscustomobject]@{Header='InitialValue'; Key='InitialValue'},
        [pscustomobject]@{Header='Modified'; Key='Modified'},
        [pscustomobject]@{Header='Comment'; Key='Comment'}
    )
    $row = Add-ExcelTable -Sheet $spec -StartRow $row -Columns $columns -Data $callApiPropertiesKey -Name 'CallAPIPropertiesKey'

    $spec.Cells.Item($row,1).Value2 = 'Grid Display Mapping'
    $row += 1
    $columns = @(
        [pscustomobject]@{Header='Column'; Key='Column'},
        [pscustomobject]@{Header='Property'; Key='Property'},
        [pscustomobject]@{Header='Conversion'; Key='Conversion'},
        [pscustomobject]@{Header='Comment'; Key='Comment'}
    )
    $row = Add-ExcelTable -Sheet $spec -StartRow $row -Columns $columns -Data $gridDisplay -Name 'GridDisplay'

    $spec.Columns('A:D').AutoFit()

    Write-Host "Creating CommonTail sheet..." -ForegroundColor Cyan
    $tailSheet = $wb.Worksheets.Add($spec)
    $tailSheet.Name = 'CommonTail'
    $tailSheet.Cells.Item(1,1).Value2 = 'Code'
    $lineRow = 2
    foreach($line in $tailLines){
        $tailSheet.Cells.Item($lineRow,1).Value2 = $line
        $lineRow++
    }
    $endRow = $lineRow - 1
    $tailRange = $tailSheet.Range($tailSheet.Cells.Item(1,1), $tailSheet.Cells.Item($endRow,1))
    $tailTable = $tailSheet.ListObjects.Add(1, $tailRange, $null, 1)
    $tailTable.Name = 'TailLines'
    $tailTable.TableStyle = 'TableStyleLight9'
    $tailSheet.Columns('A').ColumnWidth = 120

    Write-Host "Creating Readme sheet..." -ForegroundColor Cyan
    $readme = $wb.Worksheets.Add($tailSheet)
    $readme.Name = 'Readme'
    $readme.Cells.Item(1,1).Value2 = 'ItemPicking.cs Generator'
    $readme.Cells.Item(2,1).Value2 = '【使い方】'
    $readme.Cells.Item(3,1).Value2 = '1. SpecDataシートで業務固有の設定を確認・編集（1-151行目の内容）'
    $readme.Cells.Item(4,1).Value2 = '2. 開発タブ > マクロ > GenerateItemPicking を実行'
    $readme.Cells.Item(5,1).Value2 = '3. ItemPicking.csがこのブックと同じフォルダに生成されます'
    $readme.Cells.Item(6,1).Value2 = ''
    $readme.Cells.Item(7,1).Value2 = '【構造】'
    $readme.Cells.Item(8,1).Value2 = '・SpecData: 業務固有処理（1-151行目）の定義'
    $readme.Cells.Item(9,1).Value2 = '・CommonTail: 共通処理（152行目以降）のコード'
    $readme.Cells.Item(10,1).Value2 = ''
    $readme.Cells.Item(11,1).Value2 = '設計書: ItemPiking.md 参照'
    $readme.Columns('A').ColumnWidth = 80

    Write-Host "Adding VBA module..." -ForegroundColor Cyan
    $module = $wb.VBProject.VBComponents.Add(1)
    $module.Name = 'Generator'
    $module.CodeModule.AddFromString($macroCode)

    Write-Host "Saving workbook..." -ForegroundColor Cyan
    $wb.SaveAs($generatorPath, 51)
    $wb.Close($true)
    
    Write-Host "Successfully created: $generatorPath" -ForegroundColor Green
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
    throw
}
finally {
    if($module){ [System.Runtime.InteropServices.Marshal]::ReleaseComObject($module) | Out-Null }
    if($readme){ [System.Runtime.InteropServices.Marshal]::ReleaseComObject($readme) | Out-Null }
    if($tailSheet){ [System.Runtime.InteropServices.Marshal]::ReleaseComObject($tailSheet) | Out-Null }
    if($spec){ [System.Runtime.InteropServices.Marshal]::ReleaseComObject($spec) | Out-Null }
    if($wb){ [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null }
    if($excel){ $excel.Quit(); [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
