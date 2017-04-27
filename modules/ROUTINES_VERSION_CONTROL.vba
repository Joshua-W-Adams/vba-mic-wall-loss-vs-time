Attribute VB_Name = "ROUTINES_VERSION_CONTROL"
'AUTHOR: Joshua William Adams
'REV HISTORY:
'REV: A DESC.: Issued for Review                    DATE: 27/04/2017
'REV: 0 DESC.: Issued for Use                       DATE: 27/04/2017
'DESCRIPTION: Module version control code.

'DESCRIPTION: This code Exports all VBA modules
Sub SaveCodeModules()

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

'DESCRIPTION: This code Imports all VBA modules
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
