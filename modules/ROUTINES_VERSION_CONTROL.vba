Attribute VB_Name = "ROUTINES_VERSION_CONTROL"
'Option Explicit

Sub SaveCodeModules()

    'This code Exports all VBA modules
    Dim i%, sName$
    
    With ThisWorkbook.VBProject
        For i% = 1 To .VBComponents.Count
            If .VBComponents(i%).CodeModule.CountOfLines > 0 Then
                sName$ = .VBComponents(i%).CodeModule.Name
                .VBComponents(i%).Export ThisWorkbook.Path & "\modules\" & sName$ & ".vba"
            End If
        Next i
    End With

End Sub

Sub ImportCodeModules()

    With ThisWorkbook.VBProject
        For i% = 1 To .VBComponents.Count
    
            ModuleName = .VBComponents(i%).CodeModule.Name
            
            pos1 = InStr(ModuleName, "Sheet")
            pos2 = InStr(ModuleName, "ThisWorkbook")
            
            If ModuleName <> "ROUTINES_VERSION_CONTROL" And pos1 = 0 And pos2 = 0 Then
                .VBComponents.Remove .VBComponents(ModuleName)
                .VBComponents.Import ThisWorkbook.Path & "\modules\" & ModuleName & ".vba"
            End If
        Next i
    End With

End Sub
