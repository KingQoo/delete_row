Attribute VB_Name = "改ページ削除"
Option Explicit

Sub 改ページ削除()


'*****使えません******


'改ページの行のリストを取得
    Dim pageBreaks As hPageBreaks:
    Set pageBreaks = ActiveSheet.hPageBreaks
    Dim pageList As New Collection
   
'hPageBreaks型をCollection型に変換
    Dim obj As Variant
    For Each obj In ActiveSheet.hPageBreaks
        pageList.Add (obj.Location.Row)
    Next
        
    Dim reList As New Collection
    Set reList = リスト逆順(pageList)
    
    Dim i, dwnLine, upLine As Long
    For i = 1 To reList.count Step 2
         dwnLine = reList(i) - 1
         upLine = reList(i + 1)
         ActiveSheet.Range(Rows(dwnLine), _
                           Rows(upLine)).Delete
    Next i

End Sub

Public Function リスト逆順(ByRef list As Collection) As Collection
    
    Dim i As Integer
    Dim reList As Collection
    Set reList = New Collection
    
    For i = list.count To 1 Step -1
        reList.Add (list(i))
    Next i
    
    Set リスト逆順 = reList

End Function


Public Sub リスト逆順テスト()

    Dim list As New Collection
    list.Add (1)
    list.Add (2)
    list.Add (3)
        
    Dim nothingList As New Collection

    
    Debug.Assert リスト逆順(list)(1) = list(3)
    Debug.Assert リスト逆順(list)(2) = list(2)
    Debug.Assert リスト逆順(list)(3) = list(1)
    
    
    MsgBox "テスト完了"
End Sub

Public Sub 改ページ設定()
    
'１ページを取得
    Dim pageRange As Range
    On Error Resume Next
    Set pageRange = Application.InputBox( _
                        "１ページの範囲を選択してください。" _
                        , "改ページ設定" _
                        , Type:=8)
    If Err.Number <> 0 Then
        MsgBox "キャンセルされました。"
        Exit Sub
    End If
    
'印刷範囲全体を取得
    Dim printRange As Range
    Set printRange = Application.InputBox( _
                        "印刷範囲を選択してください。" _
                        , "印刷範囲の設定" _
                        , Type:=8)
    If Err.Number <> 0 Then
        MsgBox "キャンセルされました。"
        Exit Sub
    End If
    
    
    With ActiveSheet
    
        '印刷範囲初期化、設定
        .PageSetup.PrintArea = False
        .PageSetup.PrintArea = printRange.Address
        
        '改ページの初期化
        .ResetAllPageBreaks
        
        Dim i As Long
        
        '１ページの最終行から、印刷範囲の最終行まで
        '１ページの行数おきに改ページを設定
        For i = pageRange.Rows.count + 1 To printRange.Rows.count Step pageRange.Rows.count
             .Rows(i).PageBreak = xlPageBreakManual
        Next i
    End With
End Sub

Public Sub インプット伝票削除()

'任意の範囲から「インプット」を含むセルを探す

'見つけたセルの上x行、下y行を範囲選択・削除


End Sub

Public Sub 全シート適用()
    Dim Ws As Worksheet
    For Each Ws In Worksheets
        Ws.Activate
        
        If Ws.Name = "" Then
            
        End If
    Next Ws
End Sub
