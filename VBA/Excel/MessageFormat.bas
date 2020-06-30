Option Explicit

'''
' MessageFormat関数
' パラメータ埋め込みメッセージフォーマッティング処理
'
' 処理概要:
'   テンプレート文字列に対して、パラメータを埋め込んだメッセージを返却します。
'     ・テンプレート中のパラメータ埋め込み部分(プレースホルダ)は "{}" で記述します
'     ・プレースホルダは登録順に引数の埋め込みパラメータ分、置換されていきます
'
' 引数:
'   x_templateMessage - テンプレートメッセージ
'   x_params          - 埋め込みパラメータ
'
' 参考:
'   https://npnl.hatenablog.jp/entry/2020/06/02/232221
'
Function MessageFormat(ByVal x_templateMessage As String, ParamArray x_params() As Variant) As String
    Dim p_pos As Long: p_pos = 1
    Dim p_posFound As Long: p_posFound = 0
    Dim p_result As String: p_result = ""
    
    Dim i As Integer
    For i = 0 To UBound(x_params)
        ' プレースホルダ {} を検索
        p_posFound = InStr(p_pos, x_templateMessage, "{}")
        ' プレースホルダ {} が見つからなかった場合
        If p_posFound = 0 Then
            Exit For ' 置換終了
        End If
        
        p_result = p_result & Mid$(x_templateMessage, p_pos, p_posFound - p_pos)
        p_result = p_result & CStr(x_params(i))
        p_pos = p_posFound + 2
    Next i

    ' 残分を追加
    If p_pos < Len(x_templateMessage) Then
        p_result = p_result & Mid$(x_templateMessage, p_pos, Len(x_templateMessage) - p_pos + 1)
    End If
    
    ' 結果返却
    MessageFormat = p_result
End Function
