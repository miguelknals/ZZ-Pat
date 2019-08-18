Imports System.Data
Imports System.Xml
Imports System.Xml.Serialization
Imports System.IO
Public Module ModuloRutPaT 'Public de lo contrario serializer no me va.
    Public Class TestEnCommon
        Sub New(nombre As String)
            Me.nombre = nombre
        End Sub
        Dim nombre As String

    End Class
    Public Function Devuelve3() As Integer
        Return 4
    End Function
    Public Class ClaseListaReferencias
        Public todoOK As Boolean
        Public Referencias As List(Of String)
        Public Sub New()
            todoOK = True
            Referencias = New List(Of String)
        End Sub
        Sub Anyade(s As String)
            Referencias.Add(s)
        End Sub
    End Class
    Public Class ClaseListaReferenciasProyectoBase
        Public todoOK As Boolean
        Public TengoProyectoBase As Boolean
        Public ReferenciasProyectoBase As List(Of String)
        Public ProyectoBase As String
        Public ArchivoConPB As String
        Public Sub New()
            TengoProyectoBase = False ' crea pero nada mas
            todoOK = True
            ReferenciasProyectoBase = New List(Of String)
        End Sub
        Public Sub anyade(s As String) ' añade referencia
            ReferenciasProyectoBase.Add(s)
        End Sub

    End Class
    Public Class ClaseListaCarpetasTraducir
        ' ojo este código está en ZZ-CommonCode/ModuloRutPat y ZZ-Pat/MainWindows
        Public todoOK As Boolean
        Public tCarpetasTraducir As DataTable
        Public nCarpetas As Integer
        Public NombrePaT As String
        Public FechaEntrega As Date
        Public Notas As String
        Public ListaReferencias As ClaseListaReferencias
        Public ListaReferenciasProyectoBase As ClaseListaReferenciasProyectoBase
        Public FechaCreacionPaT As Date
        Public FirmaInstancia As Integer
        Public CorreoGestor As String
        Public DirectorioGestor As String
        Public NombreTraductor As String
        Public FlagDeleteFIfExists As Boolean
        Public FlagAutoInstall As Boolean

        ' ojo este código está en ZZ-CommonCode/ModuloRutPat y ZZ-Pat/MainWindows
        ' resto de informción
        Public Sub New()
            tCarpetasTraducir = New DataTable
            tCarpetasTraducir.TableName = "CarpetasATraducir"
            tCarpetasTraducir.Columns.Add(New DataColumn("ProMSS", Type.GetType("System.String")))
            tCarpetasTraducir.Columns.Add(New DataColumn("ProBase", Type.GetType("System.String")))
            tCarpetasTraducir.Columns.Add(New DataColumn("Idioma", Type.GetType("System.String")))
            tCarpetasTraducir.Columns.Add(New DataColumn("ProIBM", Type.GetType("System.String")))
            tCarpetasTraducir.Columns.Add(New DataColumn("Carpeta", Type.GetType("System.String")))
            tCarpetasTraducir.Columns.Add(New DataColumn("Envio", Type.GetType("System.String")))
            tCarpetasTraducir.Columns.Add(New DataColumn("Perfil", Type.GetType("System.String")))
            tCarpetasTraducir.Columns.Add(New DataColumn("bCNT", Type.GetType("System.Boolean")))
            tCarpetasTraducir.Columns.Add(New DataColumn("CNT", Type.GetType("System.Int32")))
            tCarpetasTraducir.Columns.Add(New DataColumn("bIniCal", Type.GetType("System.Boolean")))
            tCarpetasTraducir.Columns.Add(New DataColumn("IniCal", Type.GetType("System.Int32")))
            tCarpetasTraducir.Columns.Add(New DataColumn("bFinCal", Type.GetType("System.Boolean")))
            tCarpetasTraducir.Columns.Add(New DataColumn("FinCal", Type.GetType("System.Int32")))
            tCarpetasTraducir.Columns.Add(New DataColumn("Contaje", Type.GetType("System.Int32")))
            tCarpetasTraducir.Columns.Add(New DataColumn("Tarifa", Type.GetType("System.Single")))
            tCarpetasTraducir.Columns.Add(New DataColumn("bTotal", Type.GetType("System.Boolean")))
            tCarpetasTraducir.Columns.Add(New DataColumn("Total", Type.GetType("System.Single")))

            nCarpetas = 0
            todoOK = True

        End Sub
        Sub AnyadeCarpeta(carpeta As String, ProBase As String, Perfil As String, _
                          bCNT As Boolean, CNT As Integer, _
                          bIniCal As Boolean, IniCal As Integer, Idioma As String, envio As String, _
                          tarifa As Single)

            Dim fila As DataRow = tCarpetasTraducir.NewRow
            fila("Carpeta") = carpeta
            fila("ProBase") = ProBase
            fila("Perfil") = Perfil
            fila("Idioma") = Idioma
            fila("bCNT") = bCNT
            fila("CNT") = CNT
            fila("bIniCal") = bIniCal
            fila("IniCal") = IniCal
            fila("Envio") = Trim(envio.PadRight(16).Substring(0, 15)) ' como maximo 15
            fila("bFinCal") = False ' inicialmente 0
            fila("FinCal") = 0 ' final 0
            fila("Contaje") = 0 ' inicialmente
            fila("Tarifa") = tarifa
            fila("bTotal") = False
            fila("Total") = 0
            tCarpetasTraducir.Rows.Add(fila)
            ' tengo la carpeta desordenada
            Dim auxTabla As New DataTable
            auxTabla = tCarpetasTraducir.Clone
            Dim filas() As DataRow
            filas = tCarpetasTraducir.Select(Nothing, "Carpeta")
            For Each fila In filas
                auxTabla.ImportRow(fila)
            Next
            auxTabla.AcceptChanges() ' la ordenada
            tCarpetasTraducir = New DataTable
            tCarpetasTraducir = auxTabla.Clone
            ' ahora la recreo
            For Each row In auxTabla.Rows
                tCarpetasTraducir.ImportRow(row)
            Next
            tCarpetasTraducir.AcceptChanges() ' 


            nCarpetas += 1
            todoOK = True
        End Sub
        Sub BorraCarpeta(carpeta As String)
            ' en teoría podría borrar directamente las filas de la tabla
            ' pero al hacerlo, cuando se graba el XML no me conserva bien
            ' las etiquetas del xml generado así que la voy a recrear 
            ' la tabla y de esta forma se restablecen los puntero.
            Dim auxTabla As New DataTable
            auxTabla = tCarpetasTraducir.Clone
            Dim row As DataRow
            For Each row In tCarpetasTraducir.Rows
                If row("Carpeta") <> carpeta Then
                    auxTabla.ImportRow(row)
                End If
            Next
            auxTabla.AcceptChanges()
            tCarpetasTraducir = New DataTable
            tCarpetasTraducir = auxTabla.Clone
            ' ahora la recreo
            For Each row In auxTabla.Rows
                tCarpetasTraducir.ImportRow(row)
            Next
            tCarpetasTraducir.AcceptChanges() ' 
            '
            nCarpetas = tCarpetasTraducir.Rows.Count ' número de carpetas
            todoOK = True

        End Sub
        Function ObtenHash(secretoservidor As String) As Integer
            ' para obtener el hash serializo la clase en un string
            Dim UltimaFirma As Integer = Me.FirmaInstancia
            Me.FirmaInstancia = 0
            Dim x As New XmlSerializer(Me.GetType) ' serializo mi estructura
            Dim sw As New IO.StringWriter()
            x.Serialize(sw, Me)
            Dim auxS As String = sw.ToString
            Dim NuevaFirmaS As String = CType(auxS.GetHashCode, String) & secretoservidor
            ' restauro
            Me.FirmaInstancia = UltimaFirma
            Return NuevaFirmaS.GetHashCode
        End Function
        Function SerializateOLD(archivo As String) As Boolean
            ' el problema está en que newline lo convierte a LF. Cambiar a CR al escribir
            ' tampoco me arregla gran cosa, pq el hash es anterior a la conversión
            '
            'Dim xmlSet As New XmlWriterSettings
            'xmlSet.NewLineHandling = NewLineHandling.Entitize  ' utilizar CR
            'xmlSet.NewLineHandling = NewLineHandling.Replace
            Dim ser As New XmlSerializer(Me.GetType)
            Dim wr As XmlWriter = XmlWriter.Create(archivo)
            'Dim wr As XmlWriter = XmlWriter.Create(archivo,xmlset)
            Dim todoOK As Boolean = True
            Try
                ser.Serialize(wr, Me)

            Catch ex As Exception
                todoOK = False
            Finally
                wr.Close()
                'ex = ex.InnerException ' la normal no dice nada
                ' de http://msdn.microsoft.com/en-us/library/aa302290.aspx
                'txtSalida.AppendText(nl)
                'txtSalida.AppendText(String.Format("Message: {0}" & nl, ex.Message))
                'txtSalida.AppendText(String.Format("Exception Type: {0}" & nl, ex.GetType().FullName))
                'txtSalida.AppendText(String.Format("Source: {0}" & nl, ex.Source))
                'txtSalida.AppendText(String.Format("StrackTrace: {0}" & nl, ex.StackTrace))
                'txtSalida.AppendText(String.Format("TargetSite: {0}" & nl, ex.TargetSite))
                'Exit Function  ' no hago nada mas
            End Try

            Return todoOK
        End Function
        Function Serializate(archivo As String) As Boolean
            Dim todoOK As Boolean = True ' optimismo
            Dim objSW As New StreamWriter(archivo)
            Dim x As New XmlSerializer(Me.GetType) ' serializo mi estructura
            Try
                x.Serialize(objSW, Me) ' guardo el par
            Catch ex As Exception
                'ex = ex.InnerException ' la normal no dice nada
                '' de http://msdn.microsoft.com/en-us/library/aa302290.aspx
                'txtSalida.AppendText(nl)
                'txtSalida.AppendText(String.Format("Message: {0}" & nl, ex.Message))
                ' txtSalida.AppendText(String.Format("Exception Type: {0}" & nl, ex.GetType().FullName))
                ' txtSalida.AppendText(String.Format("Source: {0}" & nl, ex.Source))
                'txtSalida.AppendText(String.Format("StrackTrace: {0}" & nl, ex.StackTrace))
                'txtSalida.AppendText(String.Format("TargetSite: {0}" & nl, ex.TargetSite))
                'Exit Sub  ' no hago nada mas
                todoOK = False
            Finally
                objSW.Close()
            End Try
            Return todoOK
        End Function
    End Class
 



    Sub mandatoZip(mandato As String, ByRef procEC As Integer, Optional ByRef dirTra As String = "")
        procEC = 0
        Dim newProc As New Diagnostics.Process
        ' esto divide el mandato en principio (el exe) y los parámetros

        Dim principio As String = "" : Dim resto As String = "" : Dim primerblanco As Integer = 0
        primerblanco = InStr(mandato, " ")
        principio = mandato.Substring(0, primerblanco - 1)
        resto = mandato.Substring(primerblanco, Len(mandato) - primerblanco)

        'info proceso
        newProc.StartInfo.FileName = principio
        newProc.StartInfo.WindowStyle = ProcessWindowStyle.Minimized
        newProc.StartInfo.Arguments = resto
        newProc.StartInfo.RedirectStandardOutput = True
        newProc.StartInfo.UseShellExecute = False
        newProc.StartInfo.RedirectStandardOutput = True
        If dirTra <> "" Then
            newProc.StartInfo.WorkingDirectory = dirTra
        End If
        newProc.Start()
        ' Do not wait for the child process to exit before
        ' reading to the end of its redirected stream.
        ' p.WaitForExit();
        ' Read the output stream first and then wait.
        ' necesito esperar de lo contrario a veces salía de aquí sin haber acabado
        ' http://msdn.microsoft.com/en-us/library/system.diagnostics.process.standardoutput.aspx
        Dim output As String = newProc.StandardOutput.ReadToEnd()
        newProc.WaitForExit()
        procEC = newProc.ExitCode
    End Sub

    Sub ejecuta(ByVal mandato As String, ByRef procEC As Integer)
        Dim debug = False
        Dim procID As Integer
        Dim newProc As Diagnostics.Process
        If debug Then
            Dim sw As New StreamWriter("C:\tmp.tmp", True)
            sw.WriteLine(mandato)
            sw.Close()
        End If
        Dim startInfo As New ProcessStartInfo("OTMBATCH.EXE")
        startInfo.WindowStyle = ProcessWindowStyle.Minimized
        startInfo.Arguments = mandato
        newProc = Process.Start(startInfo)
        procID = newProc.Id
        'Mouse.OverrideCursor = Cursors.Wait
        newProc.WaitForExit()
        'Mouse.OverrideCursor = Nothing
        procEC = -1
        If newProc.HasExited Then
            procEC = newProc.ExitCode
        End If
        If debug Then
            Dim sw As New StreamWriter("C:\tmp.tmp", True)
            sw.WriteLine("RC=" & procEC.ToString)
            sw.Close()
        End If
    End Sub

End Module
