Imports Microsoft.Office.Tools.Ribbon

Public Class RibbonEDC

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        Globals.ThisAddIn.ExtractAndMerge()
    End Sub
End Class
