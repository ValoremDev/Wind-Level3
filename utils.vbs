Function FileThere(FileName As String) As Boolean
    FileThere = (Dir(FileName) > "")
End Function

Function UseOverwriteFile(FileName As String)
' Use or Overwrite a file in same directory
' :fileName: Complete path to file
   Set fso = CreateObject("Scripting.FileSystemObject")
   
    If FileThere(FileName) Then
        If MsgBox("Le fichier" & FileName & " existe déjà. Le remplacer?", vbYesNoCancel + vbExclamation, "Fichier existant") = vbNo Then
            Set UseOverwriteFile = fso.OpenTextFile(FileName, ForAppending)
        Else
            Set UseOverwriteFile = fso.CreateTextFile(FileName, True)
        End If
    Else
        Set UseOverwriteFile = fso.CreateTextFile(FileName, True)
    End If


End Function