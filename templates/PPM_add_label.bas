Attribute VB_Name = "PPM_add_label"
Sub PPMグラフにデータラベル範囲選択し設定()

    Dim labelRange As Range
    Dim cht As Chart

    ' ===== グラフ取得 =====
    On Error Resume Next
    Set cht = ActiveChart
    On Error GoTo 0

    If cht Is Nothing Then
        MsgBox "グラフを選択してください"
        Exit Sub
    End If

    ' ===== 範囲選択ダイアログ =====
    On Error Resume Next
    Set labelRange = Application.InputBox( _
        Prompt:="データラベルの範囲を選択してください", _
        Type:=8)
    On Error GoTo 0

    If labelRange Is Nothing Then Exit Sub

    ' ===== データラベル設定 =====
    With cht.SeriesCollection(1)
        .ApplyDataLabels
        
        .DataLabels.Format.TextFrame2.TextRange. _
            InsertChartField msoChartFieldRange, _
            "=" & labelRange.address(External:=True), 0
        
        .DataLabels.ShowRange = True
        .DataLabels.ShowValue = False
    End With

End Sub

