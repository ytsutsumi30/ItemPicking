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

Public Sub GenerateArrivalLocationDetail()
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
    
    AppendReference lines, spec.Range("ReferenceLine").Value
    AppendUsings lines, spec.ListObjects("Usings")
    AppendNamespaceHeader lines, spec
    AppendGlobalConfig lines, spec.ListObjects("GlobalConfig")
    AppendItemClass lines, spec.ListObjects("ItemProperties")
    AppendCallAPI lines, spec
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
    Dim i As Long
    Dim lastCategory As String
    lastCategory = ""
    For i = 1 To TableRowCount(data)
        Dim currentDesc As String
        currentDesc = Trim$(TableValue(data, i, 4))
        
        ' カテゴリコメントの処理
        If i = 1 Then
            lines.Add INDENT3 & "// 取得予定のプロパティ"
        ElseIf InStr(currentDesc, "Update") > 0 Or InStr(currentDesc, "更新") > 0 Then
            If lastCategory <> "update" Then
                lines.Add ""
                lines.Add INDENT3 & "// 更新情報"
                lastCategory = "update"
            End If
        ElseIf InStr(currentDesc, "key") > 0 Or InStr(currentDesc, "キー") > 0 Or InStr(currentDesc, "line") > 0 Or InStr(currentDesc, "release") > 0 Then
            If lastCategory <> "key" Then
                lines.Add ""
                lines.Add INDENT3 & "// キー"
                lastCategory = "key"
            End If
        ElseIf InStr(currentDesc, "timestamp") > 0 Or InStr(currentDesc, "pointer") > 0 Or InStr(currentDesc, "internal") > 0 Then
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
    lines.Add ""
    lines.Add INDENT3 & "List<Property> propertiesList = new List<Property>();"
    lines.Add ""
    lines.Add INDENT3 & "//使用するプロパティのデータリストを作成"
    lines.Add INDENT3 & "if(mode == 0){"
    AppendPropertyListForMode0 lines, spec.ListObjects("CallAPIProperties")
    AppendExtraPropertiesForMode0 lines, spec.ListObjects("CallAPIExtra")
    lines.Add ""
    lines.Add INDENT3 & "}"
    lines.Add INDENT3 & "//データ更新用、get時不要"
    lines.Add INDENT3 & "else{"
    lines.Add INDENT4 & "// 時間設定"
    lines.Add INDENT4 & "TimeZoneInfo jstZone = TimeZoneInfo.FindSystemTimeZoneById(""Tokyo Standard Time"");"
    lines.Add INDENT4 & "DateTime utcNow = DateTime.UtcNow;"
    lines.Add INDENT4 & "DateTime jstNow = TimeZoneInfo.ConvertTimeFromUtc(utcNow, jstZone);"
    lines.Add ""
    lines.Add INDENT4 & "propertiesList.Add(new Property { Name = ""ue_Uf_update_time_from_mongoose"", Value = jstNow.ToString(), Modified = true });"
    lines.Add INDENT4 & "propertiesList.Add(new Property { Name = ""ue_Uf_updator_from_mongoose"", Value =  ThisForm.Variables(""User"").Value, Modified = true });"
    lines.Add ""
    lines.Add INDENT4 & "// キー"
    lines.Add INDENT4 & "propertiesList.Add(new Property { Name = ""CoNum"", Value = ThisForm.Variables(""selectCoNum"").Value, Modified = (mode==1?true:false)});"
    lines.Add INDENT4 & "propertiesList.Add(new Property { Name = ""CoLine"", Value = ThisForm.Variables(""selectCoLine"").Value, Modified = (mode==1?true:false)});"
    lines.Add INDENT4 & "propertiesList.Add(new Property { Name = ""CoRelease"", Value = ThisForm.Variables(""selectCoRelease"").Value, Modified = (mode==1?true:false)});"
    lines.Add INDENT4 & ""
    lines.Add INDENT4 & "//データ特定用、insert時は不要"
    lines.Add INDENT4 & "if(mode != 1){"
    lines.Add INDENT4 & "   propertiesList.Add(new Property { Name = ""RecordDate"", Value = ThisForm.Variables(""selectRecordDate"").Value, Modified = true });"
    lines.Add INDENT4 & "   propertiesList.Add(new Property { Name = ""RowPointer"", Value = ThisForm.Variables(""selectRowPointer"").Value, Modified = true });"
    lines.Add INDENT4 & "   propertiesList.Add(new Property { Name = ""_ItemId"", Value = ThisForm.Variables(""selectItemId"").Value, Modified = true });"
    lines.Add INDENT4 & "}        "
    lines.Add INDENT3 & "}"
    lines.Add ""
    lines.Add INDENT3 & "//データを取得"
    lines.Add INDENT3 & "string JSONResponse = getData(mode,gIDOName,propertiesList);"
    lines.Add INDENT3 & "// webに移行、grid更新する必要がなくなった"
    lines.Add INDENT3 & "if(mode != 0)return;"
    lines.Add ""
    lines.Add INDENT3 & "//取得した結果はJSONなので、bodyを処理しデータを取得"
    lines.Add INDENT3 & "List<cItem> itemsList = JsonConvert.DeserializeObject<getJSONObject>(JSONResponse).Items;"
    lines.Add ""
    lines.Add INDENT3 & "//処理結果を変数に格納"
    lines.Add INDENT3 & "ThisForm.Variables(""" & spec.Range("ResultVariableName").Value & """).Value = GenerateWebSetJson(itemsList);"
    lines.Add ""
    lines.Add INDENT3 & "// webに移行、grid更新する必要がなくなった"
    lines.Add INDENT3 & "// //グリッドの初期化"
    lines.Add INDENT3 & "// int count = 1;"
    lines.Add INDENT3 & "// if(ThisForm.Components[""ResultGrid""].GetGridRowCount() > 0){"
    lines.Add INDENT3 & "//    ThisForm.Components[""ResultGrid""].DeleteGridRows(1,ThisForm.Components[""ResultGrid""].GetGridRowCount());"
    lines.Add INDENT3 & "// }"
    lines.Add INDENT3 & "//"
    lines.Add INDENT3 & "// //データを表示"
    lines.Add INDENT3 & "// foreach (var item in itemsList)"
    lines.Add INDENT3 & "// {"
    lines.Add INDENT3 & "//    ThisForm.Components[""ResultGrid""].InsertGridRows(count,1);"
    lines.Add INDENT3 & "//    ThisForm.Components[""ResultGrid""].SetGridValue(count,1,item.Item);"
    lines.Add INDENT3 & "//    ThisForm.Components[""ResultGrid""].SetGridValue(count,2,item.Description);"
    lines.Add INDENT3 & "//    // 未出荷数"
    lines.Add INDENT3 & "//    //ThisForm.Components[""ResultGrid""].SetGridValue(count,3,(float.Parse(item.QtyOrderedConv) - float.Parse(item.QtyShipped)).ToString());//小数点消すためいったんfloatに変換"
    lines.Add INDENT3 & "//    // オーダ数"
    lines.Add INDENT3 & "//    ThisForm.Components[""ResultGrid""].SetGridValue(count,3,float.Parse(item.QtyOrderedConv).ToString());//小数点消すためいったんfloatに変換"
    lines.Add INDENT3 & "//    ThisForm.Components[""ResultGrid""].SetGridValue(count,4,item.CoNum);"
    lines.Add INDENT3 & "//    ThisForm.Components[""ResultGrid""].SetGridValue(count,5,"""");"
    lines.Add INDENT3 & "//    ThisForm.Components[""ResultGrid""].SetGridValue(count,6,item.ue_Uf_update_time_from_mongoose);"
    lines.Add INDENT3 & "//    ThisForm.Components[""ResultGrid""].SetGridValue(count,7,item.ue_Uf_updator_from_mongoose);"
    lines.Add INDENT3 & "//    ThisForm.Components[""ResultGrid""].SetGridValue(count,8,item.RecordDate);"
    lines.Add INDENT3 & "//    ThisForm.Components[""ResultGrid""].SetGridValue(count,9,item.RowPointer);"
    lines.Add INDENT3 & "//    ThisForm.Components[""ResultGrid""].SetGridValue(count,10,item._ItemId);"
    lines.Add INDENT3 & "//    ThisForm.Components[""ResultGrid""].SetGridValue(count,11,item.CoLine);"
    lines.Add INDENT3 & "//    ThisForm.Components[""ResultGrid""].SetGridValue(count,12,item.CoRelease);"
    lines.Add INDENT3 & "//       "
    lines.Add INDENT3 & "//    count++;"
    lines.Add INDENT3 & "// } "
    lines.Add INDENT2 & "}"
    lines.Add ""
End Sub

Private Sub AppendPropertyListForMode0(lines As Collection, tbl As ListObject)
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    Dim data As Variant
    data = tbl.DataBodyRange.Value
    Dim i As Long
    For i = 1 To TableRowCount(data)
        Dim entry As String
        entry = INDENT4 & "propertiesList.Add(new Property { Name = """ & TableValue(data, i, 1) & """, Value = " & TableValue(data, i, 2) & ", Modified = " & TableValue(data, i, 3) & " });"
        If Len(Trim$(TableValue(data, i, 4))) > 0 Then
            entry = entry & " //" & TableValue(data, i, 4)
        End If
        lines.Add entry
    Next i
End Sub

Private Sub AppendExtraPropertiesForMode0(lines As Collection, tbl As ListObject)
    If tbl.DataBodyRange Is Nothing Then Exit Sub
    Dim data As Variant
    data = tbl.DataBodyRange.Value
    lines.Add ""
    lines.Add INDENT4 & "// データ特定用"
    Dim i As Long
    For i = 1 To TableRowCount(data)
        Dim entry As String
        entry = INDENT4 & "propertiesList.Add(new Property { Name = """ & TableValue(data, i, 1) & """, Value = """", Modified = true });"
        If Len(Trim$(TableValue(data, i, 3))) > 0 Then
            entry = entry & " // " & TableValue(data, i, 3)
        End If
        lines.Add entry
    Next i
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

$usings = @(
    [pscustomobject]@{Order=1; Namespace='System'; Category='base'; Source='MD 3.2 overview'},
    [pscustomobject]@{Order=2; Namespace='Mongoose.IDO.Protocol'; Category='mongoose'; Source='MD 3.2 API integration'},
    [pscustomobject]@{Order=3; Namespace='Mongoose.Scripting'; Category='mongoose'; Source='MD 3.2 API integration'},
    [pscustomobject]@{Order=4; Namespace='Mongoose.Core.Common'; Category='mongoose'; Source='MD 3.2 API integration'},
    [pscustomobject]@{Order=5; Namespace='Mongoose.Core.DataAccess'; Category='mongoose'; Source='MD 3.2 API integration'},
    [pscustomobject]@{Order=6; Namespace='System.Collections.Generic'; Category='framework'; Source='MD 4.2 data handling'},
    [pscustomobject]@{Order=7; Namespace='System.IO'; Category='framework'; Source='MD 4.2 data handling'},
    [pscustomobject]@{Order=8; Namespace='System.Linq'; Category='framework'; Source='MD 4.2 aggregation'},
    [pscustomobject]@{Order=9; Namespace='System.Net'; Category='framework'; Source='MD 3.2 API transport'},
    [pscustomobject]@{Order=10; Namespace='System.Text'; Category='framework'; Source='MD 3.2 API transport'},
    [pscustomobject]@{Order=11; Namespace='System.Web'; Category='framework'; Source='MD 3.2 API transport'},
    [pscustomobject]@{Order=12; Namespace='Newtonsoft.Json'; Category='external'; Source='MD 4.2 JSON schema'}
)

$globalConfig = @(
    [pscustomobject]@{Name='gIDOName'; Value='ue_ADV_SLCoitems'; Description='Target IDO name'; Source='ArrivalLocationDetail.cs line 23'},
    [pscustomobject]@{Name='gSSO'; Value='1'; Description='Use SSO flag'; Source='ArrivalLocationDetail.cs line 24'},
    [pscustomobject]@{Name='gServerId'; Value='0'; Description='API server id'; Source='ArrivalLocationDetail.cs line 25'},
    [pscustomobject]@{Name='gSuiteContext'; Value='CSI'; Description='Suite context'; Source='ArrivalLocationDetail.cs line 26'},
    [pscustomobject]@{Name='gContentType'; Value='application/json'; Description='REST content type'; Source='ArrivalLocationDetail.cs line 27'},
    [pscustomobject]@{Name='gTimeout'; Value='10000'; Description='Timeout in ms'; Source='ArrivalLocationDetail.cs line 28'},
    [pscustomobject]@{Name='gSiteName'; Value='Q72Q74BY8XUT3SKY_TRN_AJP'; Description='Mongoose configuration name'; Source='ArrivalLocationDetail.cs line 29'}
)

$itemProperties = @(
    [pscustomobject]@{JsonName='Item'; PropertyName='Item'; Type='string'; Description='Item code'; Source='ArrivalLocationDetail.cs line 37'},
    [pscustomobject]@{JsonName='Description'; PropertyName='Description'; Type='string'; Description='Item description'; Source='ArrivalLocationDetail.cs line 40'},
    [pscustomobject]@{JsonName='QtyOrderedConv'; PropertyName='QtyOrderedConv'; Type='string'; Description='Ordered quantity'; Source='ArrivalLocationDetail.cs line 43'},
    [pscustomobject]@{JsonName='QtyShipped'; PropertyName='QtyShipped'; Type='string'; Description='Shipped quantity'; Source='ArrivalLocationDetail.cs line 46'},
    [pscustomobject]@{JsonName='CoNum'; PropertyName='CoNum'; Type='string'; Description='Order number'; Source='ArrivalLocationDetail.cs line 49'},
    [pscustomobject]@{JsonName='ue_Uf_update_time_from_mongoose'; PropertyName='ue_Uf_update_time_from_mongoose'; Type='string'; Description='Update time'; Source='ArrivalLocationDetail.cs line 58'},
    [pscustomobject]@{JsonName='ue_Uf_updator_from_mongoose'; PropertyName='ue_Uf_updator_from_mongoose'; Type='string'; Description='Updator'; Source='ArrivalLocationDetail.cs line 61'},
    [pscustomobject]@{JsonName='CoLine'; PropertyName='CoLine'; Type='string'; Description='Order line'; Source='ArrivalLocationDetail.cs line 65'},
    [pscustomobject]@{JsonName='CoRelease'; PropertyName='CoRelease'; Type='string'; Description='Order release'; Source='ArrivalLocationDetail.cs line 68'},
    [pscustomobject]@{JsonName='RecordDate'; PropertyName='RecordDate'; Type='string'; Description='Record timestamp'; Source='ArrivalLocationDetail.cs line 72'},
    [pscustomobject]@{JsonName='RowPointer'; PropertyName='RowPointer'; Type='string'; Description='Row pointer'; Source='ArrivalLocationDetail.cs line 75'},
    [pscustomobject]@{JsonName='_ItemId'; PropertyName='_ItemId'; Type='string'; Description='IDO internal id'; Source='ArrivalLocationDetail.cs line 78'}
)

$callApiProperties = @(
    [pscustomobject]@{Name='Item'; InitialValue='""'; Modified='true'; Comment='Item code'; Source='ArrivalLocationDetail.cs line 95'},
    [pscustomobject]@{Name='Description'; InitialValue='""'; Modified='true'; Comment='Item description'; Source='ArrivalLocationDetail.cs line 96'},
    [pscustomobject]@{Name='QtyOrderedConv'; InitialValue='""'; Modified='true'; Comment='Ordered quantity'; Source='ArrivalLocationDetail.cs line 97'},
    [pscustomobject]@{Name='QtyShipped'; InitialValue='""'; Modified='true'; Comment='Shipped quantity'; Source='ArrivalLocationDetail.cs line 98'},
    [pscustomobject]@{Name='CoNum'; InitialValue='""'; Modified='(mode==1?true:false)'; Comment='Order number'; Source='ArrivalLocationDetail.cs line 99'},
    [pscustomobject]@{Name='ue_Uf_update_time_from_mongoose'; InitialValue='""'; Modified='true'; Comment='Update time'; Source='ArrivalLocationDetail.cs line 100'},
    [pscustomobject]@{Name='ue_Uf_updator_from_mongoose'; InitialValue='""'; Modified='true'; Comment='Updator'; Source='ArrivalLocationDetail.cs line 101'}
)

$callApiExtra = @(
    [pscustomobject]@{Name='RecordDate'; FormVariable='selectRecordDate'; Description='Record timestamp'; Source='ArrivalLocationDetail.cs line 104'},
    [pscustomobject]@{Name='RowPointer'; FormVariable='selectRowPointer'; Description='Row pointer'; Source='ArrivalLocationDetail.cs line 105'},
    [pscustomobject]@{Name='_ItemId'; FormVariable='selectItemId'; Description='IDO internal id'; Source='ArrivalLocationDetail.cs line 106'},
    [pscustomobject]@{Name='CoLine'; FormVariable='selectCoLine'; Description='Order line'; Source='ArrivalLocationDetail.cs line 109'},
    [pscustomobject]@{Name='CoRelease'; FormVariable='selectCoRelease'; Description='Order release'; Source='ArrivalLocationDetail.cs line 110'}
)

$tailLines = Get-Content -Path 'ArrivalLocationDetail.cs' | Select-Object -Skip 177

$generatorPath = Join-Path (Get-Location) 'ArrivalLocationDetailGenerator.xlsx'
if(Test-Path $generatorPath){ Remove-Item $generatorPath -Force }

$excel = $null
$wb = $null
$spec = $null
$tailSheet = $null
$module = $null
$readme = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Add()
    $spec = $wb.Worksheets.Item(1)
    $spec.Name = 'SpecData'

    $spec.Cells.Item(1,1).Value2 = 'OutputPath'
    $spec.Cells.Item(1,2).Value2 = 'ArrivalLocationDetail.cs'
    $spec.Cells.Item(1,3).Value2 = 'Output file (relative to workbook)'
    $spec.Range('B1').Name = 'OutputPath'

    $spec.Cells.Item(2,1).Value2 = 'ReferenceLine'
    $spec.Cells.Item(2,2).Value2 = '//<ref>Newtonsoft.Json.dll</ref>'
    $spec.Cells.Item(2,3).Value2 = 'Reference assembly directive'
    $spec.Range('B2').Name = 'ReferenceLine'

    $spec.Cells.Item(3,1).Value2 = 'NamespaceName'
    $spec.Cells.Item(3,2).Value2 = 'Mongoose.FormScripts'
    $spec.Cells.Item(3,3).Value2 = 'Namespace from current implementation'
    $spec.Range('B3').Name = 'NamespaceName'

    $spec.Cells.Item(4,1).Value2 = 'ClassName'
    $spec.Cells.Item(4,2).Value2 = 'ArrivalLocationDetail'
    $spec.Cells.Item(4,3).Value2 = 'FormScript class name'
    $spec.Range('B4').Name = 'ClassName'

    $spec.Cells.Item(5,1).Value2 = 'ResultVariableName'
    $spec.Cells.Item(5,2).Value2 = 'vJSONResult'
    $spec.Cells.Item(5,3).Value2 = 'WSForm variable for JSON payload'
    $spec.Range('B5').Name = 'ResultVariableName'

    $row = 9
    $spec.Cells.Item($row,1).Value2 = 'Usings'
    $row += 1
    $columns = @(
        [pscustomobject]@{Header='Order'; Key='Order'},
        [pscustomobject]@{Header='Namespace'; Key='Namespace'},
        [pscustomobject]@{Header='Category'; Key='Category'},
        [pscustomobject]@{Header='Source'; Key='Source'}
    )
    $row = Add-ExcelTable -Sheet $spec -StartRow $row -Columns $columns -Data ($usings | Sort-Object Order) -Name 'Usings'

    $spec.Cells.Item($row,1).Value2 = 'Global Configuration'
    $row += 1
    $columns = @(
        [pscustomobject]@{Header='Name'; Key='Name'},
        [pscustomobject]@{Header='Value'; Key='Value'},
        [pscustomobject]@{Header='Description'; Key='Description'},
        [pscustomobject]@{Header='Source'; Key='Source'}
    )
    $row = Add-ExcelTable -Sheet $spec -StartRow $row -Columns $columns -Data $globalConfig -Name 'GlobalConfig'

    $spec.Cells.Item($row,1).Value2 = 'IDO Item Properties'
    $row += 1
    $columns = @(
        [pscustomobject]@{Header='JsonName'; Key='JsonName'},
        [pscustomobject]@{Header='PropertyName'; Key='PropertyName'},
        [pscustomobject]@{Header='Type'; Key='Type'},
        [pscustomobject]@{Header='Description'; Key='Description'},
        [pscustomobject]@{Header='Source'; Key='Source'}
    )
    $row = Add-ExcelTable -Sheet $spec -StartRow $row -Columns $columns -Data $itemProperties -Name 'ItemProperties'

    $spec.Cells.Item($row,1).Value2 = 'callAPI Property List'
    $row += 1
    $columns = @(
        [pscustomobject]@{Header='Name'; Key='Name'},
        [pscustomobject]@{Header='InitialValue'; Key='InitialValue'},
        [pscustomobject]@{Header='Modified'; Key='Modified'},
        [pscustomobject]@{Header='Comment'; Key='Comment'},
        [pscustomobject]@{Header='Source'; Key='Source'}
    )
    $row = Add-ExcelTable -Sheet $spec -StartRow $row -Columns $columns -Data $callApiProperties -Name 'CallAPIProperties'

    $spec.Cells.Item($row,1).Value2 = 'callAPI Extra Properties'
    $row += 1
    $columns = @(
        [pscustomobject]@{Header='Name'; Key='Name'},
        [pscustomobject]@{Header='FormVariable'; Key='FormVariable'},
        [pscustomobject]@{Header='Description'; Key='Description'},
        [pscustomobject]@{Header='Source'; Key='Source'}
    )
    $row = Add-ExcelTable -Sheet $spec -StartRow $row -Columns $columns -Data $callApiExtra -Name 'CallAPIExtra'

    $spec.Columns('A:E').AutoFit()

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
    $tailSheet.Columns('A').AutoFit()

    $readme = $wb.Worksheets.Add($tailSheet)
    $readme.Name = 'Readme'
    $readme.Cells.Item(1,1).Value2 = 'ArrivalLocationDetail.cs Generator'
    $readme.Cells.Item(2,1).Value2 = '1. Review SpecData to adjust business-specific inputs (rows 1-178).'
    $readme.Cells.Item(3,1).Value2 = '2. Run Macro: Developer > Macros > GenerateArrivalLocationDetail.'
    $readme.Cells.Item(4,1).Value2 = '3. The macro writes ArrivalLocationDetail.cs next to this workbook.'
    $readme.Cells.Item(6,1).Value2 = 'Spec references originate from ArrivalLocationDetail.md (section numbers noted per row).'
    $readme.Columns('A').EntireColumn.AutoFit()

    $module = $wb.VBProject.VBComponents.Add(1)
    $module.Name = 'Generator'
    $module.CodeModule.AddFromString($macroCode)

    $wb.SaveAs($generatorPath, 51)
    $wb.Close($true)
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
