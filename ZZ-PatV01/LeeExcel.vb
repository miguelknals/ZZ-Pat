Imports System.Data
Imports System.Data.OleDb
Module LeeExcel

    Dim SerieConexion As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes'"


    Function LeeProyectoMS(ProIBM As String, Idioma As String, archivo As String, SQL As String) As String
        SerieConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='{0}';Extended Properties='Excel 12.0;HDR=YES'"
        Dim debuga As Boolean = True

        Dim SC As String = String.Format(SerieConexion, archivo)
        Dim conExcel As New OleDbConnection(SC)
        Dim cmdExcel As New OleDbCommand
        ' aquí más información de cómo acceder a las tablas, pero ahora mismo paso
        '
        ' http://aspsnippets.com/Articles/Read-and-Import-Excel-Sheet-using-ADO.Net-and-C.aspx
        ' 
        ' voy a leer 

        Dim dsAux As New DataSet
        SC = String.Format(SerieConexion, archivo)
        'SQL = "SELECT 'CodiMSS' From [Hoja1$] where IU='{0}' and  Idioma='{1}'"
        'SQL = "SELECT [ProMss] FROM [Hoja1$] where ProIBM='{0}' and  Idioma='{1}'"
        SQL = String.Format(SQL, ProIBM, Idioma)

        Dim excelConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection(SC)
        Dim resultado As String = ""

        Try

            excelConnection.Open()
            Dim dbCommand As OleDbCommand = New OleDbCommand(SQL, excelConnection)
            'Dim dataAdapter As OleDbDataAdapter = New OleDbDataAdapter(dbCommand)
            'Dim dr1 As OleDbDataReader
            'dr1 = dbCommand.ExecuteReader
            'While dr1.Read
            ' resultado = dr1(0)
            'End While
            resultado = dbCommand.ExecuteScalar

            '            dataAdapter.Fill(dsAux, "dTable")
        Catch ex As OleDbException
            resultado = "ErrBdD"
            Dim KK As String = ex.ToString
            If conExcel.State <> ConnectionState.Closed Then conExcel.Close()
        Catch e As Exception
            MsgBox(e.ToString)

        Finally
            If excelConnection.State <> ConnectionState.Closed Then excelConnection.Close()
        End Try
        Return resultado

    End Function

End Module
