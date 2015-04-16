Imports Microsoft.Office.Interop.Excel
Imports ExcelDna.Integration

Public Module CommandExamples
    Dim Application As Application = ExcelDnaUtil.Application

    ' Top run this, press Alt+F8 and type in the macro name
    Sub dnaDumpData()
        Application.Range("A5").Value = "Hello from the CustomRegistration Example"
    End Sub

    ' This uses the ExcelCommand attribute to add a menu easily (under the Add-Ins tab)
    ' and also a ShortCut (Ctrl+Shift+D)
    <ExcelCommand(MenuName:="Custom Registration Example", MenuText := "Dump into A7", ShortCut := "^D")>
    Sub dnaDumpData2()
        Application.Range("A7").Value = "Hello from the other method"
    End Sub


End Module
