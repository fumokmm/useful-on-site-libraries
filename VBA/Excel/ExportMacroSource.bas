Option Explicit

''
' プロジェクトの全ソースを一括エクスポートする
' 
' 参考:
'   https://npnl.hatenablog.jp/entry/2020/05/04/185054
'
Sub ExportMacroSource()
    ' [参照設定] Microsoft Visual Basic for Applications Extensibility 必要
    ' [参照設定] Microsoft Scripting Rungime 必要

    Dim p_fso As Scripting.FileSystemObject
    Set p_fso = New Scripting.FileSystemObject
    
    Dim p_macroDir As String
    p_macroDir = p_fso.BuildPath(Application.ActiveWorkbook.Path, "MacroSource")
    If Not p_fso.FolderExists(p_macroDir) Then
        p_fso.CreateFolder p_macroDir
    End If

    Dim p_vbComp As VBIDE.VBComponent
    Dim p_typeLabel As String
    Dim p_extension As String
    Dim p_outputFileName As String
    For Each p_vbComp In Application.VBE.ActiveVBProject.VBComponents
        ' タイプから拡張子を特定
        Select Case p_vbComp.Type
            Case vbext_ct_ActiveXDesigner
                p_typeLabel = "ActiveXDesigner"
                p_extension = "cls"
            
            Case vbext_ct_ClassModule
                p_typeLabel = "ClassModule"
                p_extension = "cls"
            
            Case vbext_ct_Document
                p_typeLabel = "Document"
                p_extension = "cls"
            
            Case vbext_ct_MSForm
                p_typeLabel = "MSForm"
                p_extension = "frm"
            
            Case vbext_ct_StdModule
                p_typeLabel = "StdModule"
                p_extension = "bas"
        End Select
    
        ' エクスポート実施
        p_outputFileName = p_fso.BuildPath(p_macroDir, p_vbComp.Name & "." & p_extension)
        Debug.Print "[export] " & p_outputFileName
        p_vbComp.Export Filename:=p_outputFileName
     
    Next p_vbComp
    
    Debug.Print "Finish export."

End Sub
