Option Explicit

' =========================================================
' 財務データ 横並び変換スクリプト
' 
' 以下の列番号（A列=1, B列=2...）は実際のExcelの形式に合わせて
' 適宜変更してください。
' =========================================================
Const COL_YEAR = 1  ' 年度が入っている列（例: A列ならば 1）
Const COL_TYPE = 2  ' 表の種類(貸借対照表等)が入っている列（例: B列ならば 2）
Const COL_ITEM = 3  ' 勘定科目が入っている列（例: C列ならば 3）
Const COL_VAL  = 4  ' 金額が入っている列（例: D列ならば 4）
Const START_ROW = 2 ' データが開始する行（見出しが1行目にある場合は 2）
' =========================================================

Dim objArgs
Set objArgs = WScript.Arguments
If objArgs.Count = 0 Then
    MsgBox "変換したいExcelファイルをこのスクリプトのアイコンにドラッグ＆ドロップしてください。", vbInformation, "使い方"
    WScript.Quit
End If

Dim filePath
filePath = objArgs(0)

Dim objExcel, objWB_in, objWS_in, objWB_out
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False

On Error Resume Next
Set objWB_in = objExcel.Workbooks.Open(filePath)
If Err.Number <> 0 Then
    MsgBox "ファイルの展開に失敗しました。対象がExcel形式か確認してください。", vbCritical, "エラー"
    objExcel.Quit
    Set objExcel = Nothing
    WScript.Quit
End If
On Error GoTo 0

Set objWS_in = objWB_in.Sheets(1)

Dim dictData, dictYears, dictTypesItems
Set dictData = CreateObject("Scripting.Dictionary")
Set dictYears = CreateObject("Scripting.Dictionary")
Set dictTypesItems = CreateObject("Scripting.Dictionary") 

Dim maxRow, r
' データの最終行を取得 (勘定科目の列を基準とする)
maxRow = objWS_in.Cells(objWS_in.Rows.Count, COL_ITEM).End(-4162).Row ' xlUp = -4162

Dim valYear, valType, valItem, valAmount, keyData
For r = START_ROW To maxRow
    valYear = Trim(CStr(objWS_in.Cells(r, COL_YEAR).Value))
    valType = Trim(CStr(objWS_in.Cells(r, COL_TYPE).Value))
    valItem = Trim(CStr(objWS_in.Cells(r, COL_ITEM).Value))
    valAmount = objWS_in.Cells(r, COL_VAL).Value
    
    If valYear <> "" And valType <> "" And valItem <> "" Then
        ' 年度の登録
        If Not dictYears.Exists(valYear) Then
            dictYears.Add valYear, valYear
        End If
        
        ' 表の種類と勘定科目の登録 (出現順序の保持)
        If Not dictTypesItems.Exists(valType) Then
            Dim tempDict
            Set tempDict = CreateObject("Scripting.Dictionary")
            dictTypesItems.Add valType, tempDict
        End If
        If Not dictTypesItems(valType).Exists(valItem) Then
            dictTypesItems(valType).Add valItem, valItem
        End If
        
        ' 金額の保存
        keyData = valType & "_" & valItem & "_" & valYear
        dictData.Item(keyData) = valAmount
    End If
Next

objWB_in.Close False

' 全年度のリストを取得して文字列順でソート（昇順）
Dim arrYears, i, j, temp
arrYears = dictYears.Keys
For i = 0 To UBound(arrYears) - 1
    For j = i + 1 To UBound(arrYears)
        If arrYears(i) > arrYears(j) Then
            temp = arrYears(i)
            arrYears(i) = arrYears(j)
            arrYears(j) = temp
        End If
    Next
Next

' 新規ブック追加
Set objWB_out = objExcel.Workbooks.Add()
' デフォルトシート数を調整
While objWB_out.Sheets.Count < dictTypesItems.Count
    objWB_out.Sheets.Add
Wend
While objWB_out.Sheets.Count > dictTypesItems.Count
    objWB_out.Sheets(objWB_out.Sheets.Count).Delete
Wend

Dim sheetIdx
sheetIdx = 1
Dim currentType, arrItems, objWS_out

' 表の種類ごとにシートを作成
For Each currentType In dictTypesItems.Keys
    Set objWS_out = objWB_out.Sheets(sheetIdx)
    ' シート名に使えない記号を除外・31文字制限
    objWS_out.Name = Left(Replace(Replace(Replace(Replace(Replace(Replace(Replace(currentType, ":", ""), "", ""), "/", ""), "?", ""), "*", ""), "[", ""), "]", ""), 31)
    
    ' 見出し行の作成
    objWS_out.Cells(1, 1).Value = "勘定科目"
    For i = 0 To UBound(arrYears)
        objWS_out.Cells(1, i + 2).Value = arrYears(i)
    Next
    
    ' 勘定科目とデータの書き込み
    arrItems = dictTypesItems(currentType).Keys
    For i = 0 To UBound(arrItems)
        valItem = arrItems(i)
        objWS_out.Cells(i + 2, 1).Value = valItem
        
        For j = 0 To UBound(arrYears)
            valYear = arrYears(j)
            keyData = currentType & "_" & valItem & "_" & valYear
            If dictData.Exists(keyData) Then
                objWS_out.Cells(i + 2, j + 2).Value = dictData.Item(keyData)
            End If
        Next
    Next
    
    ' 列幅の自動調整
    objWS_out.Columns.AutoFit
    sheetIdx = sheetIdx + 1
Next

Dim newFilePath
If InStrRev(filePath, ".") > 0 Then
    newFilePath = Left(filePath, InStrRev(filePath, ".") - 1) & "_横展開.xlsx"
Else
    newFilePath = filePath & "_横展開.xlsx"
End If

objExcel.DisplayAlerts = False
objWB_out.SaveAs newFilePath
objWB_out.Close False

objExcel.Quit
Set objExcel = Nothing

MsgBox "処理が完了しました。" & vbCrLf & "出力ファイル: " & newFilePath, vbInformation, "完了"
