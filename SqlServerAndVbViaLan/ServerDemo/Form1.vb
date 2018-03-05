Imports SqlServerAndVbViaLan.Core

Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim sqlcore As New SqlServerAndVbViaLanCore
        sqlcore.Connect()

    End Sub
End Class
