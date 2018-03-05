Imports System.Data.SqlClient

Public Class SqlServerAndVbViaLanCore


    ''' <summary>
    ''' The connection that used to connect database
    ''' </summary>
    Public mConnection As SqlConnection

    ''' <summary>
    ''' Sub used to connect to server from host
    ''' </summary>
    Public Sub ConnectToServer()
        Try
            mConnection = New SqlConnection("Data Source = ABUBAKRMAHDI-PC\SQLEXPRESS,1433;Initial Catalog=TestDB;User ID=user1;Password=1234")

            If mConnection.State = ConnectionState.Open Then mConnection.Close()
            mConnection.Open()
            MsgBox(mConnection.State)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Sub used to connect to server directly
    ''' </summary>
    Public Sub Connect()
        Try
            mConnection = New SqlConnection("Data Source=ABUBAKRMAHDI-PC\SQLEXPRESS,1433;Initial Catalog=TestDB;Integrated Security=True")

            If mConnection.State = ConnectionState.Open Then mConnection.Close()
            mConnection.Open()
            MsgBox(mConnection.State)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub



End Class
