' 上書きインポート
Sub OverwriteImport()
    ImportData True
End Sub

' 追加インポート
Sub AppendImport()
    ImportData False
End Sub

' 共通のインポート処理
Sub ImportData(isOverwrite As Boolean)
    Dim folderPath As Variant
    Dim fileName As String
    Dim wsMaster As Worksheet
    Dim wsControl As Worksheet
    Dim wbImport As Workbook
    Dim lastRow As Long
    Dim isFirstFile As Boolean
    
    ' 処理実行シートとマスタシートを設定
    Set wsControl = ThisWorkbook.Sheets("処理実行")
    Set wsMaster = ThisWorkbook.Sheets("マスタ")
    
    ' フォルダパスを取得
    folderPath = wsControl.Range("B4").Value
    If IsEmpty(folderPath) Then
        MsgBox "処理実行シートのB4セルが空白です。フォルダパスを入力してください。", vbInformation
        Exit Sub
    Else
        If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    End If
    
    ' マスタシートのクリア
    If isOverwrite Then
        wsMaster.Cells.ClearContents
        isFirstFile = True ' 上書きの場合は最初のファイルを初回として扱う
    Else
        isFirstFile = (wsMaster.Cells(1, 1).Value = "") ' マスタシートが空なら初回と判断
    End If
    
    ' フォルダ内のファイルをループ
    fileName = Dir(folderPath & "*.xls*") ' Excelファイルを対象
    Do While fileName <> ""
        Set wbImport = Workbooks.Open(folderPath & fileName)
        
        ' インポート元シートを設定（最初のシート）
        With wbImport.Sheets(1)
            ' マスタシートの最後の行を取得（空の場合は1行目に設定）
            If wsMaster.Cells(1, 1).Value = "" Then
                lastRow = 1
            Else
                lastRow = wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row + 1
            End If
            
            ' 初回のみヘッダー行を含めてコピー
            If isFirstFile Then
                Set importRange = .UsedRange
                isFirstFile = False ' 次回以降は初回でなくなる
            Else
                ' ヘッダー行を除いた範囲をコピー
                Set importRange = .Range(.Rows(2), .Rows(.UsedRange.Rows.Count))
            End If
            
            ' マスタシートに貼り付け
            wsMaster.Cells(lastRow, 1).Resize(importRange.Rows.Count, importRange.Columns.Count).Value = importRange.Value
        End With
        
        ' 取り込んだファイルを閉じる
        wbImport.Close False
        fileName = Dir ' 次のファイル
    Loop
    
    ' 後処理
    Application.CutCopyMode = False
    MsgBox "データの取り込みが完了しました。", vbInformation
End Sub

' 1件ずつPDFを作成
Sub OnceCreatePDF()
    Main False
End Sub

' 全件PDFを作成
Sub AllCreatePDF()
    Main True
End Sub

' メイン処理
Sub Main(isAllCreate As Boolean)
    Dim wsControl As Worksheet
    Dim wsMaster As Worksheet
    Dim wsAlert As Worksheet
    Dim searchValue As Variant
    Dim matchRow As Long
    Dim targetRange As Range
    Dim col As Long
    Dim avValue As Integer
    Dim awValue As String
    Dim subject As String
    Dim productCode As String
    Dim pdfPath As Variant
    Dim currentDate As String
    Dim masterLastRow As Long

    ' シートの設定
    Set wsControl = ThisWorkbook.Sheets("処理実行")
    Set wsMaster = ThisWorkbook.Sheets("マスタ")
    Set wsAlert = ThisWorkbook.Sheets("アラート")
    
    ' アウトプットパスの記載があるかチェック
    pdfPath = wsControl.Range("B11").Value
    If IsEmpty(pdfPath) Then
        MsgBox "処理実行シートのB11に値がありません。フォルダパスを入力して下さい。", vbExclamation
        Exit Sub
    End If
    
    ' アラートシートの初期化
    wsAlert.Cells.ClearContents
    wsAlert.Range("A1").Value = "商品コード"
    
    If isAllCreate Then
        ' マスタシートの最終行を取得
        masterLastRow = wsMaster.Cells(wsMaster.Rows.Count, "E").End(xlUp).Row
        
        ' マスタシートのE列を上から順に処理
        For i = 2 To masterLastRow
            targetCode = wsMaster.Cells(i, "E").Value
            ' 商品コードが空白の場合はメッセージを表示して終了
            If IsEmpty(targetCode) Then
                MsgBox "マスタシートに商品コードが空の行があります。修正してください。", vbExclamation
                Exit Sub
            Else
                ' 更新対象シート判定処理呼び出し
                ChooseTagetSheet i, targetCode
            End If
        Next i
    Else
        ' 処理実行シートB16の値を取得
        searchValue = wsControl.Range("B16").Value
        If IsEmpty(searchValue) Then
            MsgBox "処理実行シートのB16に値がありません。商品コードを入力して下さい。", vbExclamation
            Exit Sub
        Else
            ' マスタシートのE列で検索
            On Error Resume Next
            matchRow = wsMaster.Columns("E").Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole).Row
            On Error GoTo 0
            If matchRow = 0 Then
                MsgBox "マスタシートに一致する商品コードが見つかりません。処理実行シートのB16の値を修正してください。", vbExclamation
                Exit Sub
            Else
                 ' 更新対象シート判定処理呼び出し
                ChooseTagetSheet matchRow, searchValue
            End If
        End If
    End If
MsgBox "処理終了しました。失敗している商品コードがないかアラートシートを確認してください。", vbExclamation
End Sub
    
' 更新シート判定処理
Sub ChooseTagetSheet(targetRow As Variant, produceCode As Variant)
    Dim wsControl As Worksheet
    Dim wsMaster As Worksheet
    Dim wsAlert As Worksheet
    Dim col As Long

    ' シートの設定
    Set wsControl = ThisWorkbook.Sheets("処理実行")
    Set wsMaster = ThisWorkbook.Sheets("マスタ")
    Set wsAlert = ThisWorkbook.Sheets("アラート")
    

    ' 小学生シート対象かチェック（M列からU列）
        For col = 13 To 21
            If wsMaster.Cells(targetRow, col).Value = 1 Then
                UpdateElementarySheet targetRow, produceCode
                Exit Sub
            End If
        Next col

        ' 中学生シート対象かチェック（V列からAB列）
        For col = 22 To 28
            If wsMaster.Cells(targetRow, col).Value = 1 Then
                UpdateJuniorSheet targetRow, produceCode
                Exit Sub
            End If
        Next col

        ' 高校生シート対象（AC列からAF列）
        For col = 29 To 32
            If wsMaster.Cells(targetRow, col).Value = 1 Then
                UpdateHighSheet targetRow, produceCode
                Exit Sub
            End If
        Next col

End Sub

' 小学生シート更新
Sub UpdateElementarySheet(targetRow As Variant, produceCode As Variant)
    Dim wsControl As Worksheet
    Dim wsMaster As Worksheet
    Dim wsElementary As Worksheet
    Dim wsAlert As Worksheet
    Dim target As String
    Dim targetRange As Range
    Dim col As Long
    Dim avValue As Variant
    Dim awValue As Variant
    Dim axValue As Variant
    Dim subject As String
    Dim pdfPath As String
    Dim currentDate As String

    ' シートの設定
    Set wsControl = ThisWorkbook.Sheets("処理実行")
    Set wsMaster = ThisWorkbook.Sheets("マスタ")
    Set wsElementary = ThisWorkbook.Sheets("小学生")
    Set wsAlert = ThisWorkbook.Sheets("アラート")
        
    ' 小学生シートの対象を背景色白、文字色グレーに
    With wsElementary
        .Range("D10:AM12").Interior.Color = RGB(255, 255, 255)
        .Range("D10:AM12").Font.Color = RGB(192, 192, 192)
    End With
    
    ' マスタシートのM列～U列に1以外の値があったらアラートシートに商品コードを記入して終了
    For col = 13 To 21
        If wsMaster.Cells(targetRow, col).Value <> 1 And wsMaster.Cells(targetRow, col).Value <> "" Then
            wsAlert.Cells(wsAlert.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = produceCode
            Exit Sub
        Else
            If wsMaster.Cells(targetRow, col).Value = 1 Then
                ' 学年に応じて色を変更
                Select Case col
                    Case 13 ' 小1
                        Set targetRange = wsElementary.Range("D10:G12")
                    Case 14 ' 小2
                        Set targetRange = wsElementary.Range("H10:K12")
                    Case 15 ' 小3
                        Set targetRange = wsElementary.Range("L10:O12")
                    Case 16 ' 小4
                        Set targetRange = wsElementary.Range("P10:S12")
                    Case 17 ' 小5
                        Set targetRange = wsElementary.Range("T10:W12")
                    Case 18 ' 小6
                        Set targetRange = wsElementary.Range("X10:AA12")
                    Case 19 ' 中学受験
                        Set targetRange = wsElementary.Range("AB10:AH12")
                    Case 20 ' 公立中高一貫校受験対策
                        Set targetRange = wsElementary.Range("AB10:AH12")
                    Case 21 ' 中学準備講座
                        Set targetRange = wsElementary.Range("AI10:AM12")
                End Select

                targetRange.Interior.Color = RGB(226, 107, 10) '  背景色オレンジ
                targetRange.Font.Color = RGB(255, 255, 255) ' 文字色白
                
            End If
        End If
    Next col

    ' 小学生シートの難易度を背景色白に
    With wsElementary
        .Range("D5:AM8").Interior.Color = RGB(255, 255, 255)
    End With

    ' マスタシートのAV列、AW列の値を取得
    avValue = wsMaster.Cells(targetRow, 48).Value
    If IsEmpty(avValue) Then
        wsAlert.Cells(wsAlert.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = produceCode
        Exit Sub
    Else
        awValue = wsMaster.Cells(targetRow, 49).Value
        If IsEmpty(awValue) Then
            wsAlert.Cells(wsAlert.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = produceCode
            Exit Sub
        End If
    End If

    ' 難易度に応じて背景色黄色を設定
    target = avValue & ":" & awValue
    wsElementary.Range(target).Interior.Color = RGB(255, 255, 0)

    ' マスタシートのAX列の値を取得
    axValue = wsMaster.Cells(targetRow, 50).Value
    If IsEmpty(axValue) Then
        wsAlert.Cells(wsAlert.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = produceCode
        Exit Sub
    Else
     ' 小学生シートの判数へ転記
        With wsElementary
            .Range("A14").Value = axValue
            .Range("A14").Font.Color = RGB(255, 255, 255)
        End With
    End If

    ' 小学生シートの科目の背景色はグレー、文字色は白に設定
    With wsElementary
        .Range("J14:AM16").Interior.Color = RGB(192, 192, 192)
        .Range("J14:AM16").Font.Color = RGB(255, 255, 255)
    End With

    ' マスタシートのAG列の文字列を取得
    subject = wsMaster.Cells(targetRow, 33).Value

    ' 科目に応じて色を変更
    Select Case subject
        Case "算数"
            Set targetRange = wsElementary.Range("J14:O16")
        Case "国語"
            Set targetRange = wsElementary.Range("P14:U16")
        Case "理科"
            Set targetRange = wsElementary.Range("V14:AA16")
        Case "社会"
            Set targetRange = wsElementary.Range("AB14:AG16")
        Case "英語"
            Set targetRange = wsElementary.Range("AH14:AM16")
        Case Else
            wsAlert.Cells(wsAlert.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = searchValue
            Exit Sub
    End Select

    targetRange.Interior.Color = RGB(226, 107, 10) '  背景色オレンジ
    targetRange.Font.Color = RGB(255, 255, 255) ' 文字色白
    
    ' 日付形式のフォーマット
    currentDate = Format(Now, "yyyymmdd_Hhmmss")

    ' 小学生シートをPDFに出力
    pdfPath = wsControl.Range("B11").Value & "\小学生_" & produceCode & "_" & currentDate & ".pdf"
    wsElementary.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath
    
End Sub


' 中学生シート更新
Sub UpdateJuniorSheet(targetRow As Variant, produceCode As Variant)
    Dim wsControl As Worksheet
    Dim wsMaster As Worksheet
    Dim wsElementary As Worksheet
    Dim wsAlert As Worksheet
    Dim target As String
    Dim targetRange As Range
    Dim col As Long
    Dim avValue As Variant
    Dim awValue As Variant
    Dim axValue As Variant
    Dim subject As String
    Dim pdfPath As String
    Dim currentDate As String

    ' シートの設定
    Set wsControl = ThisWorkbook.Sheets("処理実行")
    Set wsMaster = ThisWorkbook.Sheets("マスタ")
    Set wsJunior = ThisWorkbook.Sheets("中学生")
    Set wsAlert = ThisWorkbook.Sheets("アラート")
    
    ' 中学生シートの対象を背景色白、文字色グレーに
    With wsJunior
        .Range("D10:AM12").Interior.Color = RGB(255, 255, 255) '文字色白
        .Range("D10:AM12").Font.Color = RGB(192, 192, 192)
    End With
    
    ' マスタシートのV列～AB列に1以外の値があったらアラートシートに商品コードを記入して終了
    For col = 22 To 28
        If wsMaster.Cells(targetRow, col).Value <> 1 And wsMaster.Cells(targetRow, col).Value <> "" Then
            wsAlert.Cells(wsAlert.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = produceCode
            Exit Sub
        Else
            If wsMaster.Cells(targetRow, col).Value = 1 Then
                ' 学年に応じて色を変更
                Select Case col
                    Case 22 ' 中1
                        Set targetRange = wsJunior.Range("D10:K12")
                    Case 23 ' 中2
                        Set targetRange = wsJunior.Range("L10:S12")
                    Case 24 ' 中3
                        Set targetRange = wsJunior.Range("T10:AA12")
                    Case 25 ' 高校受験
                        Set targetRange = wsJunior.Range("AB10:AH12")
                    Case 26 ' 1・2年の復習教材
                        Set targetRange = wsJunior.Range("AB10:AH12")
                    Case 27 ' まとめ教材
                        Set targetRange = wsJunior.Range("AB10:AH12")
                    Case 28 ' 高校準備講座
                        Set targetRange = wsJunior.Range("AI10:AM12")
                End Select
                
                targetRange.Interior.Color = RGB(83, 141, 213) ' 背景色青
                targetRange.Font.Color = RGB(255, 255, 255) '文字色白
                
            End If
        End If
    Next col

    ' 中学生シートの難易度を背景色白に
    With wsJunior
        .Range("D5:AM8").Interior.Color = RGB(255, 255, 255)
    End With

    ' マスタシートのAV列、AW列の値を取得
    avValue = wsMaster.Cells(targetRow, 48).Value
    If IsEmpty(avValue) Then
        wsAlert.Cells(wsAlert.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = produceCode
        Exit Sub
    Else
        awValue = wsMaster.Cells(targetRow, 49).Value
        If IsEmpty(awValue) Then
            wsAlert.Cells(wsAlert.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = produceCode
            Exit Sub
        End If
    End If

    ' 難易度に応じて背景色黄色を設定
    target = avValue & ":" & awValue
    wsJunior.Range(target).Interior.Color = RGB(255, 255, 0)

    ' マスタシートのAX列の値を取得
    axValue = wsMaster.Cells(targetRow, 50).Value
    If IsEmpty(axValue) Then
        wsAlert.Cells(wsAlert.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = produceCode
        Exit Sub
    Else
     ' 中学生シートの判数へ転記
        With wsJunior
            .Range("A14").Value = axValue
            .Range("A14").Font.Color = RGB(255, 255, 255)
        End With
    End If

    ' 中学生シートの科目の背景色グレー、文字色白に設定
    With wsJunior
        .Range("J14:AM16").Interior.Color = RGB(192, 192, 192)
        .Range("J14:AM16").Font.Color = RGB(255, 255, 255)
    End With

    ' マスタシートのAG列の文字列を取得
    subject = wsMaster.Cells(targetRow, 33).Value

    ' 科目に応じて色を変更
    Select Case subject
        Case "数学"
            Set targetRange = wsJunior.Range("J14:O16")
        Case "英語"
            Set targetRange = wsJunior.Range("P14:U16")
        Case "国語"
            Set targetRange = wsJunior.Range("V14:AA16")
        Case "理科"
            Set targetRange = wsJunior.Range("AB14:AG16")
        Case "社会"
            Set targetRange = wsJunior.Range("AH14:AM16")
        Case Else
            wsAlert.Cells(wsAlert.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = produceCode
            Exit Sub
    End Select

    targetRange.Interior.Color = RGB(83, 141, 213) ' 背景色青
    targetRange.Font.Color = RGB(255, 255, 255) ' 文字色白

    ' 日付形式のフォーマット
    currentDate = Format(Now, "yyyymmdd_Hhmmss")

    ' 中学生シートをPDFに出力
    pdfPath = wsControl.Range("B11").Value & "\中学生_" & produceCode & "_" & currentDate & ".pdf"
    wsJunior.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath
    
End Sub

' 高校生シート更新
Sub UpdateHighSheet(targetRow As Variant, produceCode As Variant)
    Dim wsControl As Worksheet
    Dim wsMaster As Worksheet
    Dim wsHigh As Worksheet
    Dim wsAlert As Worksheet
    Dim target As String
    Dim targetRange As Range
    Dim col As Long
    Dim avValue As Variant
    Dim awValue As Variant
    Dim axValue As Variant
    Dim subject As String
    Dim pdfPath As String
    Dim currentDate As String

    ' シートの設定
    Set wsControl = ThisWorkbook.Sheets("処理実行")
    Set wsMaster = ThisWorkbook.Sheets("マスタ")
    Set wsHigh = ThisWorkbook.Sheets("高校生")
    Set wsAlert = ThisWorkbook.Sheets("アラート")
    
    ' 高校生シートの対象を背景色白、文字色グレーに
    With wsHigh
        .Range("D10:AM12").Interior.Color = RGB(255, 255, 255)
        .Range("D10:AM12").Font.Color = RGB(192, 192, 192)
    End With
    
    ' マスタシートのAC列～AF列に1以外の値があったらアラートシートに商品コードを記入して終了
    For col = 29 To 32
        If wsMaster.Cells(targetRow, col).Value <> 1 And wsMaster.Cells(targetRow, col).Value <> "" Then
            wsAlert.Cells(wsAlert.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = produceCode
            Exit Sub
        Else
            If wsMaster.Cells(targetRow, col).Value = 1 Then
                ' 学年に応じて色を変更
                Select Case col
                    Case 29 ' 高1
                        Set targetRange = wsHigh.Range("D10:L12")
                    Case 30 ' 高2
                        Set targetRange = wsHigh.Range("M10:U12")
                    Case 31 ' 高3
                        Set targetRange = wsHigh.Range("V10:AD12")
                    Case 32 ' 大学受験
                        Set targetRange = wsHigh.Range("AE10:AM12")
                End Select
                
                targetRange.Interior.Color = RGB(146, 208, 80) ' 背景色黄緑
                targetRange.Font.Color = RGB(255, 255, 255) '文字色白
            End If
        End If
    Next col

    ' 高校生シートの難易度を背景色白に
    With wsHigh
        .Range("D5:AM8").Interior.Color = RGB(255, 255, 255)
    End With

    ' マスタシートのAV列、AW列の値を取得
    avValue = wsMaster.Cells(targetRow, 48).Value
    If IsEmpty(avValue) Then
        wsAlert.Cells(wsAlert.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = produceCode
        Exit Sub
    Else
        awValue = wsMaster.Cells(targetRow, 49).Value
        If IsEmpty(awValue) Then
            wsAlert.Cells(wsAlert.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = produceCode
            Exit Sub
        End If
    End If

    ' 難易度に応じて背景色黄色を設定
    target = avValue & ":" & awValue
    wsHigh.Range(target).Interior.Color = RGB(255, 255, 0)

    ' マスタシートのAX列の値を取得
    axValue = wsMaster.Cells(targetRow, 50).Value
    If IsEmpty(axValue) Then
        wsAlert.Cells(wsAlert.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = produceCode
        Exit Sub
    Else
     ' 高校生シートの判数へ転記
        With wsHigh
            .Range("A14").Value = axValue
            .Range("A14").Font.Color = RGB(255, 255, 255)
        End With
    End If

    ' 高校生シートの科目の背景色をグレー、文字色を白に設定
    With wsHigh
        .Range("J14:AM16").Interior.Color = RGB(192, 192, 192)
        .Range("J14:AM16").Font.Color = RGB(255, 255, 255)
    End With

    ' マスタシートのAG列の文字列を取得
    subject = wsMaster.Cells(targetRow, 33).Value

    ' 科目に応じて色を変更
    Select Case subject
        Case "数学"
            Set targetRange = wsHigh.Range("J14:O16")
        Case "英語"
            Set targetRange = wsHigh.Range("P14:U16")
        Case "国語"
            Set targetRange = wsHigh.Range("V14:AA16")
        Case "理科"
            Set targetRange = wsHigh.Range("AB14:AG16")
        Case "社会"
            Set targetRange = wsHigh.Range("AH14:AM16")
        Case Else
            wsAlert.Cells(wsAlert.Rows.Count, 1).End(xlUp).Offset(1, 0).Value = produceCode
            Exit Sub
    End Select

    targetRange.Interior.Color = RGB(146, 208, 80) ' 黄緑
    targetRange.Font.Color = RGB(255, 255, 255) ' 文字色白

    ' 日付形式のフォーマット
    currentDate = Format(Now, "yyyymmdd_Hhmmss")

    ' 高校生シートをPDFに出力
    pdfPath = wsControl.Range("B11").Value & "\高校生_" & searchValue & "_" & currentDate & ".pdf"
    wsHigh.ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdfPath
    
End Sub

