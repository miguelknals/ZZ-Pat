Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
<Serializable()> Public Structure structParametros
    Dim ParDirTemporal As String
    Dim ParDirSalidaPaT As String
    Dim ParDirTPot As String
    Dim ParArchivoExcel As String
    Dim ParSQL As String
    Dim ParSecretoServidor As String
    Dim ParIVA As Single
    Dim ParIRPF As Single
    Dim ParTarifaPredeterminada As Single
    Dim ParPuertoUDP As Long
    Dim ParHostUDP As String
    Dim ParProveedor As String
    Dim ParRecopilarUso As Boolean
    Dim ParMasDeUnPB As Boolean
End Structure
<Serializable()> Public Structure structReferenciasProyectoBase
    Dim ProyectoBase As String
    Dim Referencias As List(Of String)
End Structure



Module General

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
        Mouse.OverrideCursor = Cursors.Wait
        newProc.WaitForExit()
        Mouse.OverrideCursor = Nothing
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

    End Sub
    Function ObtenArchivoParametros() As String
        Dim auxS As String = ""
        auxS = System.Reflection.Assembly.GetExecutingAssembly().Location


        ' ahora me interesa la parte ejecutable 
        Dim FullPath As String = Path.GetDirectoryName(auxS)


        Dim archivoValores As String = FullPath & "\" & Path.GetFileNameWithoutExtension(auxS) & "_PAR.XML"
        Return archivoValores

    End Function
    Function CargaParametros() As structParametros ' devuelve los parámetors
        ' tengo que encontrar el directorio de mi ejecutable
        'Dim archivoParametros As String = ObtenArchivoParametros()
        ' será ejecutable.XML
        Dim Parametros As structParametros
        ' creo valores predeterminados por si no puedo leer
        Parametros.ParDirTemporal = My.Settings.ParDirTemporal
        Parametros.ParDirSalidaPaT = My.Settings.ParDirSalida
        Parametros.ParDirTPot = My.Settings.ParDirPB
        Parametros.ParArchivoExcel = My.Settings.ParArchivoEXCEL
        Parametros.ParSQL = My.Settings.ParSQL
        Parametros.ParSecretoServidor = My.Settings.ParSecreto
        Parametros.ParIVA = My.Settings.ParIVA
        Parametros.ParIRPF = My.Settings.ParIRPF
        Parametros.ParTarifaPredeterminada = My.Settings.ParTarifaPredeterminada
        Parametros.ParPuertoUDP = My.Settings.ParPuertoUDP
        Parametros.ParHostUDP = My.Settings.ParHostUDP
        Parametros.ParRecopilarUso = My.Settings.ParRecopilarUso
        Parametros.ParProveedor = My.Settings.ParProveedor
        Parametros.ParMasDeUnPB = My.Settings.ParMasDeUnPB

        
        Return Parametros '

    End Function
    Function CargaParametrosOLD() As structParametros ' devuelve los parámetors
        ' tengo que encontrar el directorio de mi ejecutable
        Dim archivoParametros As String = ObtenArchivoParametros()
        ' será ejecutable.XML
        Dim Parametros As structParametros
        ' creo valores predeterminados por si no puedo leer
        Parametros.ParDirTemporal = ""
        Parametros.ParDirSalidaPaT = ""
        Parametros.ParDirTPot = ""
        Parametros.ParArchivoExcel = ""
        Parametros.ParSQL = ""
        Parametros.ParSecretoServidor = ""
        Parametros.ParIVA = 0
        Parametros.ParIRPF = 0

        Dim todoOK As Boolean = False
        If IO.File.Exists(archivoParametros) = True Then ' lo intento leer
            Try
                Dim objStreamReader As New StreamReader(archivoParametros)
                Dim x As New XmlSerializer(Parametros.GetType)
                Parametros = x.Deserialize(objStreamReader)
                objStreamReader.Close()
            Catch ex As Exception
                ' no puedo hacer nada, el archivo está mal
            End Try
            'Deserialize text file to a new object.
        End If

        Return Parametros '

    End Function

    Sub GuardaParametros(par As structParametros)
        With My.Settings
            .ParDirSalida = par.ParDirSalidaPaT
            .ParDirPB = par.ParDirTPot
            .ParDirTemporal = par.ParDirTemporal
            .ParArchivoEXCEL = par.ParArchivoExcel
            .ParSQL = par.ParSQL
            .ParSecreto = par.ParSecretoServidor
            .ParIVA = par.ParIVA
            .ParIRPF = par.ParIRPF
            .ParTarifaPredeterminada = par.ParTarifaPredeterminada
            .ParPuertoUDP = par.ParPuertoUDP
            .ParHostUDP = par.ParHostUDP
            .ParRecopilarUso = par.ParRecopilarUso
            .ParProveedor = par.ParProveedor
            .ParMasDeUnPB = par.ParMasDeUnPB
            .Save()
        End With

    End Sub

    Sub GuardaParametrosOLD(Par As structParametros)

        ' tengo que encontrar el directorio de mi ejecutable
        Dim archivoParametros As String = ObtenArchivoParametros()

        ' será ejecutable.XML
        Dim todoOK As Boolean = False

        If IO.File.Exists(archivoParametros) = True Then ' lo intento leer
            ' lo borro
            IO.File.Delete(archivoParametros)
        End If
        Dim ObjSW As New StreamWriter(archivoParametros) ' lo guardaré en mi ejecutuable
        Dim x As New XmlSerializer(Par.GetType) ' serializo mi estructura
        x.Serialize(ObjSW, Par) ' guardo el par
        ObjSW.Close()

    End Sub


End Module
