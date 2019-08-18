Imports System.Data
Imports System.Data.OleDb
Imports Npgsql
Imports NpgsqlTypes

Module LeeCodigoProveedor

    Dim SerieConexion As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes'"
    Structure strucCodigoProveedor

        Dim CodProv As String
        Dim TodoOk As Boolean
        Dim info As String
        Function BuscaCodigoProveedorEXCEL(carpeta As String, SQLAlternativo As Boolean)
            Me.CodProv = "N/D" : Me.TodoOk = False : Me.info = ""
            Me.info = "N/D"
            Dim par As structParametros
            par = CargaParametros()
            If par.ParTipoOrigen <> "EXCEL" Then
                Me.CodProv = "ErrInt"
                Me.TodoOk = False
                Me.info = String.Format("Internal Error source type is '{0}', but BuscaCodigoProveedorEXCEL have been called. ", par.ParTipoOrigen)
                Return Nothing
            End If

            Dim auxSQLMascara As String = par.ParSQL
            If SQLAlternativo Then auxSQLMascara = par.ParSQL_Alternativo

            Dim inst As New strucInterpretaSQL(carpeta, auxSQLMascara)
            'inst = InterpretaSQL(inst)
            inst.interpreta() ' 
            ' si me da error, ya ni sigo
            If inst.TodoOK = False Then
                Me.CodProv = "ErrSQL"
                Me.info = String.Format("SQL cannot be resolved for {0} and {1}. Check syntaxis", carpeta, auxSQLMascara)
                Me.TodoOk = False
                Return Nothing
            End If
            Dim sql As String = inst.sql
            Dim info As String = inst.info


            SerieConexion = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='{0}';Extended Properties='Excel 12.0;HDR=YES'"
            Dim debuga As Boolean = True

            Dim SC As String = String.Format(serieconexion, par.ParArchivoExcel)
            Dim conExcel As New OleDbConnection(SC)
            Dim cmdExcel As New OleDbCommand
            ' aquí más información de cómo acceder a las tablas, pero ahora mismo paso
            '
            ' http://aspsnippets.com/Articles/Read-and-Import-Excel-Sheet-using-ADO.Net-and-C.aspx
            ' 
            ' voy a leer 

            Dim dsAux As New DataSet
            SC = String.Format(serieconexion, par.ParArchivoExcel)
            'SQL = "SELECT 'CodiMSS' From [Hoja1$] where IU='{0}' and  Idioma='{1}'"
            'SQL = "SELECT [ProMss] FROM [Hoja1$] where ProIBM='{0}' and  Idioma='{1}'"

            Dim excelConnection As OleDb.OleDbConnection = New OleDb.OleDbConnection(SC)
            Dim resultado As String = ""

            Try

                excelConnection.Open()
                Dim dbCommand As OleDbCommand = New OleDbCommand(sql, excelConnection)
                'Dim dataAdapter As OleDbDataAdapter = New OleDbDataAdapter(dbCommand)
                'Dim dr1 As OleDbDataReader
                'dr1 = dbCommand.ExecuteReader
                'While dr1.Read
                ' resultado = dr1(0)
                'End While
                Me.CodProv = dbCommand.ExecuteScalar

                '            dataAdapter.Fill(dsAux, "dTable")
            Catch ex As OleDbException
                Me.CodProv = "ErrBdD"
                Me.TodoOk = False
                Me.info = String.Format("Command = '{0}' / Error = '{1}'", sql, ex.Message.ToString)
                Return Nothing
            Catch e As Exception
                Me.CodProv = "ErrorInt"
                Me.TodoOk = False
                Me.info = String.Format("Internal error {0}", e.Message.ToString)
                Return Nothing
            Finally
                If excelConnection.State <> ConnectionState.Closed Then excelConnection.Close()
            End Try
            If Me.CodProv = "" Then
                Me.CodProv = "N/D"
                Me.TodoOk = False
                Me.info = String.Format("Ref. not found for sql {0}", sql)
            Else
                Me.TodoOk = True
                Me.info = ""

            End If
            Return Nothing

        End Function


        Function BuscaCodigoProveedorPOSTGRES(carpeta As String, SQLAlternativo As Boolean)
            Me.CodProv = "N/D" : Me.TodoOk = False : Me.info = ""
            Me.info = "N/D"
            Dim par As structParametros
            par = CargaParametros()
            If par.ParTipoOrigen <> "POSTGRES" Then
                Me.CodProv = "ErrInt"
                Me.TodoOk = False
                Me.info = String.Format("Internal Error source type is '{0}', but BuscaCodigoProveedorPOSTGRES have been called. ", par.ParTipoOrigen)
                Return Nothing
            End If

            Dim auxSQLMascara As String = par.ParSQL
            If SQLAlternativo Then auxSQLMascara = par.ParSQL_Alternativo

            Dim inst As New strucInterpretaSQL(carpeta, auxSQLMascara)
            'inst = InterpretaSQL(inst)
            inst.interpreta() ' 
            ' si me da error, ya ni sigo
            If inst.TodoOK = False Then
                Me.CodProv = "ErrSQL"
                Me.info = String.Format("SQL cannot be resolved for {0} and {1}. Check syntaxis", carpeta, auxSQLMascara)
                Me.TodoOk = False
                Return Nothing
            End If
            Dim sql As String = inst.sql
            Dim info As String = inst.info


            ' "Server=192.168.0.48;Port=5432;User Id=mss;Password=mss1928;Database=kmkey_zodb;"
            ' "Server={0};Port={1};User Id={2};Password={3};Database={4};"
            Dim contrasenya As String = Decrypt(par.ParPGContrasenya)
            contrasenya = Replace(contrasenya, par.ParPGUsuario, "", , 1)
            Dim serieconexion As String = "Server={0};Port={1};User Id={2};Password={3};Database={4};"
            serieconexion = String.Format(serieconexion, par.ParPGHost, par.ParPGPuerto, par.ParPGUsuario, contrasenya, par.ParPGBaseDatos)
            Dim con As New NpgsqlConnection(serieconexion)
            Dim cmd As New NpgsqlCommand
            'Dim sql As String = "select reference, title, * from kmkey_project where title like '%AIX%' limit 10"
            Dim dr As NpgsqlDataReader
            Dim encontrado As Boolean = False
            Try
                con.Open()

                cmd.CommandText = sql
                cmd.Connection = con
                dr = cmd.ExecuteReader
                While dr.Read
                    ' Console.WriteLine("{0} {1}", dr("reference"), dr("title")) 
                    Me.CodProv = dr("reference") : encontrado = True : Me.TodoOk = True : Me.info = "" : Exit While
                End While

            Catch ex As Exception
                'Dim auxS As String = ex.Message.ToString
                Me.CodProv = "ErrBD"
                Me.info = String.Format("Database error: {0}", ex.Message.ToString)
                Me.TodoOk = False
                Return Nothing
            Finally
                If con.State <> ConnectionState.Closed Then con.Close()
            End Try
            If encontrado = False Then
                Me.info = String.Format("Ref. not found for sql {0}", sql)
                Me.TodoOk = False
            End If

            Return Nothing

        End Function

    End Structure
    Function LeeCodigoPOSTGRES_OBSOLETA(carpeta As String, SQLAlternativo As Boolean) As String
        ' primero he de obtener el SQL real
        Dim CodProv As String = "N/D"
        Dim par As structParametros
        par = CargaParametros()
        If par.ParTipoOrigen <> "POSTGRES" Then
            CodProv = "ErrInt"
            Return CodProv
        End If

        Dim auxSQLMascara As String = par.ParSQL
        If SQLAlternativo Then auxSQLMascara = par.ParSQL_Alternativo

        Dim inst As New strucInterpretaSQL(carpeta, auxSQLMascara)
        'inst = InterpretaSQL(inst)
        inst.interpreta() ' 
        ' si me da error, ya ni sigo
        If inst.TodoOK = False Then
            CodProv = "ErrSQL"
            Return CodProv
        End If
        Dim sql As String = inst.sql
        Dim info As String = inst.info


        ' "Server=192.168.0.48;Port=5432;User Id=mss;Password=mss1928;Database=kmkey_zodb;"
        ' "Server={0};Port={1};User Id={2};Password={3};Database={4};"
        Dim contrasenya As String = Decrypt(par.ParPGContrasenya)
        contrasenya = Replace(contrasenya, par.ParPGUsuario, "", , 1)
        Dim serieconexion As String = "Server={0};Port={1};User Id={2};Password={3};Database={4};"
        serieconexion = String.Format(serieconexion, par.ParPGHost, par.ParPGPuerto, par.ParPGUsuario, contrasenya, par.ParPGBaseDatos)
        Dim con As New NpgsqlConnection(serieconexion)
        Dim cmd As New NpgsqlCommand
        'Dim sql As String = "select reference, title, * from kmkey_project where title like '%AIX%' limit 10"
        Dim dr As NpgsqlDataReader

        Try
            con.Open()

            cmd.CommandText = sql
            cmd.Connection = con
            dr = cmd.ExecuteReader
            While dr.Read
                ' Console.WriteLine("{0} {1}", dr("reference"), dr("title")) 
                CodProv = dr("reference") : Exit While
            End While

        Catch ex As Exception
            'Dim auxS As String = ex.Message.ToString
            CodProv = "ErrBD"
        Finally
            If con.State <> ConnectionState.Closed Then con.Close()
        End Try

        Return CodProv
        'Console.ReadLine()

    End Function



    Function LeeProyectoMS_OBSOLETA(ProIBM As String, Idioma As String, archivo As String, SQL As String) As String
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
