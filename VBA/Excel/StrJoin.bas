Option Explicit

''
' 文字列の結合関数
'
' 処理概要:
'   文字列の結合を行います。
'   特徴
'     ・値数式で参照可能
'     ・範囲でも指定可能
'     ・直接指定可能
'     ・接続文字、先頭、末尾文字を指定可能
'
' 引数:
'  x_startStr  - 開始文字列
'  x_endStr    - 終了文字列
'  x_joinStr   - 接続文字列
'  x_values    - 文字列や範囲（複数指定可）
'
Public Function StrJoin(x_startStr As String, x_endStr As String, x_joinStr As String, ParamArray x_values() As Variant) As String
    Dim p_result As String: p_result = ""
    Dim p_ratch As Boolean: p_ratch = False
    Dim p_value As Variant
    For Each p_value In x_values
        If TypeName(p_value) = "Range" Then
            Dim p_RangeValue As Range: Set p_RangeValue = p_value
            Dim p_rng As Range
            For Each p_rng In p_RangeValue
                If p_ratch Then p_result = p_result & x_joinStr Else p_ratch = True
                p_result = p_result & p_rng.Value
            Next p_rng
        Else
            If p_ratch Then p_result = p_result & x_joinStr Else p_ratch = True
            p_result = p_result & CStr(p_value)
        End If
    Next p_value
    StrJoin = x_startStr & p_result & x_endStr
End Function
