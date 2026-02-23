Option Explicit

' =========================================================
' ブロック型CSVデータ 横並び変換スクリプト (VBScript版) - 階層・重複対応・不要行・空シート対応版
' 
' 動作環境: Windows機 (Excelがインストールされていること)
' 使い方: このスクリプト(.vbs)にCSVファイルをドラッグ＆ドロップしてください。
' =========================================================

Function ArrayInsertAfter(arr, valToInsert, prevVal)
    Dim newArr(), i, idx, currentBound
    
    currentBound = -1
    On Error Resume Next
    currentBound = UBound(arr)
    On Error GoTo 0
    
    If currentBound < 0 Then
        ReDim newArr(0)
        newArr(0) = valToInsert
        ArrayInsertAfter = newArr
        Exit Function
    End If
    
    idx = -1
    For i = 0 To currentBound
        If arr(i) = prevVal Then
            idx = i
            Exit For
        End If
    Next
    
    ReDim newArr(currentBound + 1)
    
    If idx = -1 Then
        For i = 0 To currentBound
            newArr(i) = arr(i)
        Next
        newArr(currentBound + 1) = valToInsert
    Else
        For i = 0 To idx
            newArr(i) = arr(i)
        Next
        newArr(idx + 1) = valToInsert
        For i = idx + 1 To currentBound
            newArr(i + 1) = arr(i)
        Next
    End If
    
    ArrayInsertAfter = newArr
End Function

Function ArrayAdd(arr, valToInsert)
    Dim newArr(), i, currentBound
    currentBound = -1
    On Error Resume Next
    currentBound = UBound(arr)
    On Error GoTo 0
    
    If currentBound < 0 Then
        ReDim newArr(0)
        newArr(0) = valToInsert
    Else
        ReDim newArr(currentBound + 1)
        For i = 0 To currentBound
            newArr(i) = arr(i)
        Next
        newArr(currentBound + 1) = valToInsert
    End If
    ArrayAdd = newArr
End Function

Function IsTargetSheet(name)
    Dim targetNames, i
    targetNames = Array("連結貸借対照表", "連結損益計算書", "連結包括利益計算書", "連結株主資本等変動計算書", "連結損益（及び包括利益）計算書", "連結キャッシュ・フロー計算書")
    IsTargetSheet = False
    For i = 0 To UBound(targetNames)
        If name = targetNames(i) Then
            IsTargetSheet = True
            Exit For
        End If
    Next
End Function

Function GetIndentLevel(str)
    Dim i, c, level
    level = 0
    For i = 1 To Len(str)
        c = Mid(str, i, 1)
        If c = " " Or c = "　" Or c = vbTab Then
            level = level + 1
        Else
            Exit For
        End If
    Next
    GetIndentLevel = level
End Function


Dim objArgs
Set objArgs = WScript.Arguments
If objArgs.Count = 0 Then
    MsgBox "変換したいCSVファイルをこのスクリプトのアイコンにドラッグ＆ドロップしてください。", vbInformation, "使い方"
    WScript.Quit
End If

Dim filePath
filePath = objArgs(0)

If LCase(Right(filePath, 4)) <> ".csv" Then
    MsgBox "CSVファイル以外のファイルが指定されました。" & vbCrLf & "「.csv」のファイルを対象としてください。", vbExclamation, "エラー"
    WScript.Quit
End If

Dim objFSO, objFile
Set objFSO = CreateObject("Scripting.FileSystemObject")
On Error Resume Next
Set objFile = objFSO.OpenTextFile(filePath, 1, False, 0)
If Err.Number <> 0 Then
    MsgBox "ファイルの展開に失敗しました。他のプログラムで開かれているか確認してください。", vbCritical, "エラー"
    Set objFSO = Nothing
    WScript.Quit
End If
On Error GoTo 0

Dim dictData, dictYears, dictItemsOrder, dictItemNames
Set dictData = CreateObject("Scripting.Dictionary")
Set dictYears = CreateObject("Scripting.Dictionary")
Set dictItemsOrder = CreateObject("Scripting.Dictionary")
Set dictItemNames = CreateObject("Scripting.Dictionary")

Dim currentYear, currentType, prevItem, firstYear
Dim line, cols, col0, col1, col2, col3, itemName, amount
Dim stackLevels(), stackNames(), stackCount

currentYear = ""
currentType = ""
firstYear = ""
prevItem = ""
stackCount = 0

Do Until objFile.AtEndOfStream
    line = objFile.ReadLine
    cols = Split(line, ",")
    
    col0 = "" : col1 = "" : col2 = "" : col3 = ""
    If UBound(cols) >= 0 Then col0 = Trim(cols(0))
    If UBound(cols) >= 1 Then col1 = Trim(cols(1))
    If UBound(cols) >= 2 Then col2 = Trim(cols(2))
    If UBound(cols) >= 3 Then col3 = Trim(cols(3))
    
    ' 元のインデントを維持するため、前後のクォートだけを剥がした rawCol0 を用意
    Dim rawCol0
    rawCol0 = ""
    If UBound(cols) >= 0 Then
        rawCol0 = cols(0)
        Do While Left(rawCol0, 1) = """" And Len(rawCol0) > 0
            rawCol0 = Mid(rawCol0, 2)
        Loop
        Do While Right(rawCol0, 1) = """" And Len(rawCol0) > 0
            rawCol0 = Left(rawCol0, Len(rawCol0) - 1)
        Loop
    End If
    
    col0 = Replace(col0, """", "")
    col1 = Replace(col1, """", "")
    col2 = Replace(col2, """", "")
    col3 = Replace(col3, """", "")
    
    ' 不要なヘッダーデータを強制除外
    If InStr(col0, "現在") > 0 And (InStr(col0, "/") > 0 Or InStr(col0, "年") > 0) Then
        ' 1. 年度の判定
        currentYear = Trim(Replace(col0, "現在", ""))
        If Not dictYears.Exists(currentYear) Then
            dictYears.Add currentYear, currentYear
        End If
        If firstYear = "" Then
            firstYear = currentYear
        End If
        col0 = "" ' 以降の勘定科目取得ロジックには流さない
        
    ElseIf col0 = "表名称" Then
        ' 2. 表名称の判定 (厳密化)
        If IsTargetSheet(col1) Then
            currentType = col1
            prevItem = ""
            stackCount = 0
            If Not dictData.Exists(currentType) Then
                Set dictData.Item(currentType) = CreateObject("Scripting.Dictionary")
                Set dictItemNames.Item(currentType) = CreateObject("Scripting.Dictionary")
                dictItemsOrder.Add currentType, Array()
            End If
        End If
        col0 = "" ' 以降の勘定科目取得ロジックには流さない
        
    ElseIf col0 = "企業名" Or col0 = "証券ｺｰﾄﾞ" Or col0 = "（百万円）" Or (InStr(col0, "/") > 0 And InStr(col0, "-") > 0) Then
        ' 3. 不要なヘッダー行の除外
        col0 = "" ' 無視するフラグの代わり
        
    ElseIf col0 <> "" And col1 <> "" And col2 = "" And col3 = "" Then
        If IsTargetSheet(col0) Then
            currentType = col0
            prevItem = ""
            stackCount = 0
            If Not dictData.Exists(currentType) Then
                Set dictData.Item(currentType) = CreateObject("Scripting.Dictionary")
                Set dictItemNames.Item(currentType) = CreateObject("Scripting.Dictionary")
                dictItemsOrder.Add currentType, Array()
            End If
            col0 = "" ' 表見出し行そのものは勘定科目に登録しない
        End If
    End If
    
    ' 4. 勘定科目と金額の取得
    If col0 <> "" And currentType <> "" Then
        itemName = RTrim(rawCol0)
        amount = ""
        
        Dim numStr2, numStr3
        numStr2 = Replace(Replace(col2, "-", ""), ",", "")
        numStr3 = Replace(Replace(col3, "-", ""), ",", "")
        
        If numStr2 <> "" And IsNumeric(numStr2) Or col2 = "-" Then
            amount = col2
        ElseIf numStr3 <> "" And IsNumeric(numStr3) Or col3 = "-" Then
            amount = col3
        End If
        
        ' --- 階層キーの生成 (同名「その他」等の重複対応) ---
        Dim currentIndent, kk
        currentIndent = GetIndentLevel(rawCol0)
        
        Dim newStackCount
        newStackCount = stackCount
        Do While newStackCount > 0
            If stackLevels(newStackCount - 1) >= currentIndent Then
                newStackCount = newStackCount - 1
            Else
                Exit Do
            End If
        Loop
        stackCount = newStackCount
        
        ReDim Preserve stackLevels(stackCount)
        ReDim Preserve stackNames(stackCount)
        stackLevels(stackCount) = currentIndent
        stackNames(stackCount) = col0
        stackCount = stackCount + 1
        
        Dim uniqueKey
        uniqueKey = ""
        For kk = 0 To stackCount - 1
            If uniqueKey = "" Then
                uniqueKey = stackNames(kk)
            Else
                uniqueKey = uniqueKey & "::" & stackNames(kk)
            End If
        Next
        
        ' 表示名の保存 (VBScriptの辞書内で重複して保存しても名前は変わらないのでOK)
        dictItemNames.Item(currentType).Item(uniqueKey) = itemName
        
        ' 改良アンカー挿入の実装
        Dim arrItems, bExists, k
        arrItems = dictItemsOrder.Item(currentType)
        bExists = False
        
        On Error Resume Next
        For k = 0 To UBound(arrItems)
            If arrItems(k) = uniqueKey Then
                bExists = True
                Exit For
            End If
        Next
        On Error GoTo 0
        
        If Not bExists Then
            ' 1年目は無条件末尾追加 (重複バグ防止)
            If currentYear <> firstYear Then
                arrItems = ArrayInsertAfter(arrItems, uniqueKey, prevItem)
            Else
                arrItems = ArrayAdd(arrItems, uniqueKey)
            End If
            dictItemsOrder.Item(currentType) = arrItems
        End If
        
        prevItem = uniqueKey
        
        If Not dictData.Item(currentType).Exists(uniqueKey) Then
            Set dictData.Item(currentType).Item(uniqueKey) = CreateObject("Scripting.Dictionary")
        End If
        
        If currentYear <> "" And amount <> "" Then
            dictData.Item(currentType).Item(uniqueKey).Item(currentYear) = amount
        End If
    End If
Loop
objFile.Close

If dictData.Count = 0 Then
    MsgBox "解析できるデータが見つかりませんでした。フォーマットを確認してください。", vbExclamation, "エラー"
    WScript.Quit
End If

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

Dim objExcel, objWB_out, objWS_out
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = False
objExcel.DisplayAlerts = False

Set objWB_out = objExcel.Workbooks.Add()

Dim validSheetCount
validSheetCount = 0

Dim typeName, valItem, valAmount, numVal, displayName

For Each typeName In dictData.Keys
    arrItems = dictItemsOrder.Item(typeName)
    
    ' 空シート防止チェック（抽出された勘定科目が1件以上存在する場合のみシートを作成）
    Dim hasItems
    hasItems = False
    On Error Resume Next
    If UBound(arrItems) >= 0 Then hasItems = True
    On Error GoTo 0
    
    If hasItems Then
        validSheetCount = validSheetCount + 1
        If validSheetCount = 1 Then
            Set objWS_out = objWB_out.Sheets(1)
        Else
            Set objWS_out = objWB_out.Sheets.Add(, objWB_out.Sheets(objWB_out.Sheets.Count))
        End If
        
        objWS_out.Name = Left(Replace(Replace(Replace(Replace(Replace(Replace(Replace(typeName, ":", ""), "\", ""), "/", ""), "?", ""), "*", ""), "[", ""), "]", ""), 31)
        
        objWS_out.Cells(1, 1).Value = "勘定科目"
        For i = 0 To UBound(arrYears)
            objWS_out.Cells(1, i + 2).Value = arrYears(i)
        Next
        
        On Error Resume Next
        For i = 0 To UBound(arrItems)
            valItem = arrItems(i)
            
            ' 画面に表示する名前は、重複しないための階層キーではなく元の項目名にする
            displayName = valItem
            If dictItemNames.Item(typeName).Exists(valItem) Then
                displayName = dictItemNames.Item(typeName).Item(valItem)
            End If
            
            objWS_out.Cells(i + 2, 1).Value = displayName
            
            For j = 0 To UBound(arrYears)
                currentYear = arrYears(j)
                If dictData.Item(typeName).Item(valItem).Exists(currentYear) Then
                    valAmount = dictData.Item(typeName).Item(valItem).Item(currentYear)
                    
                    valAmount = Replace(valAmount, ",", "")
                    If IsNumeric(valAmount) And valAmount <> "-" Then
                        numVal = CDbl(valAmount)
                        objWS_out.Cells(i + 2, j + 2).Value = numVal
                    Else
                        objWS_out.Cells(i + 2, j + 2).Value = valAmount
                    End If
                End If
            Next
        Next
        On Error GoTo 0
        objWS_out.Columns.AutoFit
    End If
Next

' デフォルトの余分なSheet2, Sheet3等があれば削除
While objWB_out.Sheets.Count > validSheetCount
    objWB_out.Sheets(objWB_out.Sheets.Count).Delete
Wend

If validSheetCount = 0 Then
    objWB_out.Close False
    objExcel.Quit
    Set objExcel = Nothing
    MsgBox "有効な勘定科目データが1件も見つかりませんでした。", vbExclamation, "エラー"
    WScript.Quit
End If

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

MsgBox "処理が完了しました（階層・重複対応・不要データ除去版）。" & vbCrLf & "出力ファイル: " & newFilePath, vbInformation, "完了"
