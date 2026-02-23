import os

vbs_code = """Option Explicit

' =========================================================
' ブロック型CSVデータ 横並び変換スクリプト (VBScript版)
' 
' 動作環境: Windows機 (Excelがインストールされていること)
' 使い方: このスクリプト(.vbs)にCSVファイルをドラッグ＆ドロップしてください。
' =========================================================

Dim objArgs
Set objArgs = WScript.Arguments
If objArgs.Count = 0 Then
    MsgBox "変換したいCSVファイルをこのスクリプトのアイコンにドラッグ＆ドロップしてください。", vbInformation, "使い方"
    WScript.Quit
End If

Dim filePath
filePath = objArgs(0)

' 拡張子チェック
If LCase(Right(filePath, 4)) <> ".csv" Then
    MsgBox "CSVファイル以外のファイルが指定されました。" & vbCrLf & "「.csv」のファイルを対象としてください。", vbExclamation, "エラー"
    WScript.Quit
End If

' FileSystemObjectの準備
Dim objFSO, objFile
Set objFSO = CreateObject("Scripting.FileSystemObject")
On Error Resume Next
' Shift-JISで開く (TristateFalse = 0)
Set objFile = objFSO.OpenTextFile(filePath, 1, False, 0)
If Err.Number <> 0 Then
    MsgBox "ファイルの展開に失敗しました。他のプログラムで開かれているか確認してください。", vbCritical, "エラー"
    Set objFSO = Nothing
    WScript.Quit
End If
On Error GoTo 0

' データ格納用辞書
Dim dictData, dictYears, dictItemsOrder
Set dictData = CreateObject("Scripting.Dictionary")
Set dictYears = CreateObject("Scripting.Dictionary")
Set dictItemsOrder = CreateObject("Scripting.Dictionary")

Dim currentYear, currentType, line, cols
Dim col0, col1, col2, col3, itemName, amount, keyData

currentYear = ""
currentType = ""

Do Until objFile.AtEndOfStream
    line = objFile.ReadLine
    ' 簡易的なCSVパース (ダブルクォーテーションが含まれない前提)
    cols = Split(line, ",")
    
    col0 = "" : col1 = "" : col2 = "" : col3 = ""
    If UBound(cols) >= 0 Then col0 = Trim(cols(0))
    If UBound(cols) >= 1 Then col1 = Trim(cols(1))
    If UBound(cols) >= 2 Then col2 = Trim(cols(2))
    If UBound(cols) >= 3 Then col3 = Trim(cols(3))
    
    ' ダブルクォーテーションの除去
    col0 = Replace(col0, """", "")
    col1 = Replace(col1, """", "")
    col2 = Replace(col2, """", "")
    col3 = Replace(col3, """", "")
    
    ' 1. 「年度」の判定
    If InStr(col0, "現在") > 0 Then
        currentYear = Trim(Replace(col0, "現在", ""))
        If Not dictYears.Exists(currentYear) Then
            dictYears.Add currentYear, currentYear
        End If
    ' 2. 表名称の判定
    ElseIf col0 = "表名称" Then
        currentType = col1
        If Not dictData.Exists(currentType) Then
            Set dictData.Item(currentType) = CreateObject("Scripting.Dictionary")
            Set dictItemsOrder.Item(currentType) = CreateObject("Scripting.Dictionary")
        End If
    ElseIf col0 <> "" And col1 <> "" And col2 = "" And col3 = "" Then
        ' より厳密な表名判定 (A列B列があってCD列が空)
        If InStr(col0, "連結") > 0 Or InStr(col0, "計算書") > 0 Or InStr(col0, "キャッシュ・フロー") > 0 Or currentType = "" Then
            currentType = col0
            If Not dictData.Exists(currentType) Then
                Set dictData.Item(currentType) = CreateObject("Scripting.Dictionary")
                Set dictItemsOrder.Item(currentType) = CreateObject("Scripting.Dictionary")
            End If
        End If
    ' 3. 勘定科目と金額の取得
    ElseIf col0 <> "" And currentType <> "" Then
        itemName = col0
        amount = ""
        
        ' 金額の判定 (C列かD列)
        Dim numStr2, numStr3
        numStr2 = Replace(Replace(col2, "-", ""), ",", "")
        numStr3 = Replace(Replace(col3, "-", ""), ",", "")
        
        If numStr2 <> "" And IsNumeric(numStr2) Or col2 = "-" Then
            amount = col2
        ElseIf numStr3 <> "" And IsNumeric(numStr3) Or col3 = "-" Then
            amount = col3
        End If
        
        ' 出現順序の記録
        If Not dictItemsOrder.Item(currentType).Exists(itemName) Then
            dictItemsOrder.Item(currentType).Add itemName, itemName
        End If
        
        If Not dictData.Item(currentType).Exists(itemName) Then
            Set dictData.Item(currentType).Item(itemName) = CreateObject("Scripting.Dictionary")
        End If
        
        If currentYear <> "" And amount <> "" Then
            dictData.Item(currentType).Item(itemName).Item(currentYear) = amount
        End If
    End If
Loop
objFile.Close

If dictData.Count = 0 Then
    MsgBox "解析できるデータが見つかりませんでした。フォーマットを確認してください。", vbExclamation, "エラー"
    WScript.Quit
End If

' 年度のソート（古い順）
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

' Excelの起動と出力
Dim objExcel, objWB_out, objWS_out
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False

Set objWB_out = objExcel.Workbooks.Add()
' デフォルトシート数を調整
While objWB_out.Sheets.Count < dictData.Count
    objWB_out.Sheets.Add , objWB_out.Sheets(objWB_out.Sheets.Count)
Wend
While objWB_out.Sheets.Count > dictData.Count
    objWB_out.Sheets(objWB_out.Sheets.Count).Delete
Wend

Dim sheetIdx
sheetIdx = 1
Dim typeName, arrItems, valItem, valAmount, numVal

For Each typeName In dictData.Keys
    Set objWS_out = objWB_out.Sheets(sheetIdx)
    objWS_out.Name = Left(Replace(Replace(Replace(Replace(Replace(Replace(Replace(typeName, ":", ""), "\", ""), "/", ""), "?", ""), "*", ""), "[", ""), "]", ""), 31)
    
    ' 見出し行
    objWS_out.Cells(1, 1).Value = "勘定科目"
    For i = 0 To UBound(arrYears)
        objWS_out.Cells(1, i + 2).Value = arrYears(i)
    Next
    
    ' データ書き込み
    arrItems = dictItemsOrder.Item(typeName).Keys
    For i = 0 To UBound(arrItems)
        valItem = arrItems(i)
        objWS_out.Cells(i + 2, 1).Value = valItem
        
        For j = 0 To UBound(arrYears)
            currentYear = arrYears(j)
            If dictData.Item(typeName).Item(valItem).Exists(currentYear) Then
                valAmount = dictData.Item(typeName).Item(valItem).Item(currentYear)
                
                ' 数値変換を試みる
                On Error Resume Next
                valAmount = Replace(valAmount, ",", "")
                If IsNumeric(valAmount) And valAmount <> "-" Then
                    numVal = CDbl(valAmount)
                    objWS_out.Cells(i + 2, j + 2).Value = numVal
                Else
                    objWS_out.Cells(i + 2, j + 2).Value = valAmount
                End If
                On Error GoTo 0
            End If
        Next
    Next
    objWS_out.Columns.AutoFit
    sheetIdx = sheetIdx + 1
Next

Dim newFilePath
If InStrRev(filePath, ".") > 0 Then
    newFilePath = Left(filePath, InStrRev(filePath, ".") - 1) & "_横展開.xlsx"
Else
    newFilePath = filePath & "_横展開.xlsx"
End If

objWB_out.SaveAs newFilePath
objWB_out.Close False
objExcel.Quit
Set objExcel = Nothing

MsgBox "処理が完了しました。" & vbCrLf & "出力ファイル: " & newFilePath, vbInformation, "完了"
"""

vbs_path = "/home/tomo/work_office/財務データ横展開_CSV対応版.vbs"
with open(vbs_path, "w", encoding="cp932", errors="ignore") as f:
    f.write(vbs_code)

print("VBScript (CSV版) created at", vbs_path)
