Imports Microsoft.Office.Tools.Ribbon
Public Class RibbonEdc
    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Globals.ThisAddIn.ExtractAndMerge()
    End Sub
End Class
