Imports System
Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
Imports System.Data
Imports System.Text
' Para el refresh
'
Imports System.Threading
Imports System.Runtime.CompilerServices
Imports System.Windows.Threading
Imports System.Net
Imports System.Net.Mail


Public Class ProcesaManifiesto
    Dim nl As String = Environment.NewLine
    Dim ListaCarpetasTraducir As New ClaseListaCarpetasTraducir
    ' aquí los parametros para toda la apliación
    Dim ParDirTemporal As String = ""
    Dim ParDirSalidaPaT As String = ""
    Dim ParLongInterno As Long = 0
    Dim ParMailSSL As Boolean = True
    Dim ParMailPuerto As Integer = 0
    Dim ParMailContrase As String = ""
    Dim ParMailHost As String = ""
    Dim ParMailUsuario As String = ""
    Dim ViaArchivoMF As String = ""
    Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Dim Par As structParametros
        Par = CargaParametros()
        ParDirTemporal = Environment.ExpandEnvironmentVariables(Par.ParDirTemporal)
        ParDirSalidaPaT = Par.ParDirSalidaPaT
        ParLongInterno = Par.ParLongInterno
        ParMailSSL = Par.ParMailSSL
        ParMailPuerto = Par.ParMailPuerto
        ParMailContrase = Par.ParMailContrase
        ParMailHost = Par.ParMailHost
        ParMailUsuario = Par.ParMailUsuario

        ' necesito los perfiles del opentm, así que copio código
        Dim InfoOTM As New ClassInfoOTM
        If InfoOTM.TodoOK = False Then
            Dim temp As Window = New Window()
            temp.Visibility = Windows.Visibility.Hidden
            temp.Show()
            MessageBox.Show(temp, InfoOTM.Info, "OTM Error", MessageBoxButton.OK, MessageBoxImage.Error)
            Application.Current.Shutdown()
            Exit Sub
        End If

        ' Dim DiscoOTM As String = InfoOTM.DiscoOTM
        ' CargaListaCarpetasDDW(InfoOTM.Carpetas)



    End Sub


    Private Sub txtArchivoMF_PreviewDragEnter(sender As Object, e As DragEventArgs) Handles txtArchivoMF.PreviewDragEnter
        e.Effects = DragDropEffects.All
        e.Handled = True
    End Sub

    Private Sub txtArchivoMF_PreviewDrop(sender As Object, e As DragEventArgs) Handles txtArchivoMF.PreviewDrop

        Dim files As String()
        files = e.Data.GetData(DataFormats.FileDrop, False)
        Dim filename As String

        For Each filename In files
            txtArchivoMF.Text = filename ' solo admito el primero
            Exit For
        Next



    End Sub

    Private Sub txtArchivoMF_PreviewDragOver(sender As Object, e As DragEventArgs) Handles txtArchivoMF.PreviewDragOver
        e.Handled = True
    End Sub


    Private Sub btnCarga_Click(sender As Object, e As RoutedEventArgs) Handles btnCarga.Click
        SubProcesaManifiesto()
    End Sub
    Sub SubProcesaManifiesto()
        Dim auxi As Integer = 0
        ' vamos a jugar
        ' primero voy a leer el manifiesto
        txtSalida.Clear()
        ' voy a ver si existe el archivo
        '
        ' ListaCarpetas a traducir es global
        '
        If Not File.Exists(txtArchivoMF.Text) Then
            txtSalida.AppendText(String.Format("'' does not seem a file or is not found" & nl, txtArchivoMF.Text))
            MsgBox("Pls, drag and drop a PaT manifest file (*_MFT.XML)", MsgBoxStyle.Critical, "No manifiest file")
            Exit Sub
        End If
        ' 
        Dim auxS As String = "" ' variable auxiliar
        Dim ArchivoMF As String = txtArchivoMF.Text
        ViaArchivoMF = Path.GetDirectoryName(ArchivoMF)
        ' miro si acaba en \
        'If ViaArchivoMF.Substring(Len(ViaArchivoMF) - 1, 1) <> "\" Then ViaArchivoMF &= "\"
        txtSalida.AppendText("Loading... " & ArchivoMF & " ....")
        ' lo leeo 
        Try
            Dim objStreamReader As New StreamReader(ArchivoMF)
            Dim x As New XmlSerializer(ListaCarpetasTraducir.GetType)
            ListaCarpetasTraducir = x.Deserialize(objStreamReader)
            objStreamReader.Close()
        Catch ex As Exception
            auxS = "Fatal error: Looks like xml manifest is not valid. "
            txtSalida.AppendText(nl & auxS & nl)
            MsgBox(auxS, MsgBoxStyle.Critical, "MFT file not valid")
            ' no puedo hacer nada, el archivo está mal
            Exit Sub
        End Try
        txtSalida.AppendText("Done. " & nl)
        ' voy a comprobar que no me han manipulado el archivo
        ' calculo la firma
        ' Modalidad automática
        Dim OriI As Integer = 0
        If ListaCarpetasTraducir.FirmaInstancia <> ListaCarpetasTraducir.ObtenHash(ParLongInterno.ToString) Then
            txtSalida.AppendText("Firm error. Looks like the signature has changed. " & nl)
            MsgBox("Look like the signature of the manifest file has changed.", MsgBoxStyle.Critical, "Signature error - MFT has changed")
            'Exit Sub
            'test 
            If False Then ' crea TEST_FMT es de debug
                Dim ArchivoMFXML As String = "C:\u\tra\TEST_FMT.XML"
                Dim ObjSW As New StreamWriter(ArchivoMFXML) ' lo guardaré en mi ejecutuable
                Dim x As New XmlSerializer(ListaCarpetasTraducir.GetType) ' serializo mi estructura
                Try
                    x.Serialize(ObjSW, ListaCarpetasTraducir) ' guardo el par
                Catch ex As Exception
                    ex = ex.InnerException ' la normal no dice nada
                    ' de http://msdn.microsoft.com/en-us/library/aa302290.aspx
                    txtSalida.AppendText(nl)
                    txtSalida.AppendText(String.Format("Message: {0}" & nl, ex.Message))
                    txtSalida.AppendText(String.Format("Exception Type: {0}" & nl, ex.GetType().FullName))
                    txtSalida.AppendText(String.Format("Source: {0}" & nl, ex.Source))
                    txtSalida.AppendText(String.Format("StrackTrace: {0}" & nl, ex.StackTrace))
                    txtSalida.AppendText(String.Format("TargetSite: {0}" & nl, ex.TargetSite))
                    Exit Sub  ' no hago nada mas
                Finally
                    ObjSW.Close()
                End Try

            End If


            Exit Sub


        End If
        txtSalida.AppendText("Manifest firm is valid." & nl)


        ' automatización 
        ' Voy a ver si se ha enviado a alguien
        Dim NombreTraductor As String = ListaCarpetasTraducir.NombreTraductor
        Dim TipoDestino As String = ""
        Dim TargetDirectorio As String = ""
        Dim TargetDirectorioSinUnidad As String = ""
        'Dim TargetMascara As String = ListaCarpetasTraducir.DirectorioGestor & "\{0}.fxz"
        TargetDirectorio = ViaArchivoMF
        TargetDirectorioSinUnidad = _
                    Path.DirectorySeparatorChar & TargetDirectorio.Replace(Path.GetPathRoot(TargetDirectorio), "")


        Dim destino As New structDestino()
        If NombreTraductor <> "NONE" Then ' voy a intentar copiar carpetas o fxz
            txtSalida.AppendText(String.Format("Translator {0} found. Retreiving information...", NombreTraductor) & nl)
            Dim DestinoEncontrado As Boolean = False
            Dim Destinos As structListaDestinos
            Destinos = CargaDestinos()
            For Each destino In Destinos.ListaDestino
                If NombreTraductor = destino.DestNombre Then DestinoEncontrado = True : Exit For ' he encontrado mi destino
            Next
            If Not DestinoEncontrado Then ' no está en la lista de destinos
                auxS = String.Format("Translator '{0}' information NOT found. Looks like translator name has changed/deleted. ", NombreTraductor) & nl & nl
                auxS &= "You can continue, but only makes sense to do so if in the manifest directory are the PaT folders, otherwise no process will be done. " & nl & nl
                auxS &= "Do you want to continue?" & nl
                txtSalida.AppendText(auxS)
                auxi = MsgBox(auxS, MsgBoxStyle.YesNo, "Translator information not found.")
                If auxi = MsgBoxResult.No Then '
                    auxS = "Process ended by user. Abort. " & nl
                    txtSalida.AppendText(auxS)
                    Exit Sub
                End If
                ' continuamos sin traductor
                NombreTraductor = "NONE"
                GoTo continua_sin_traductor
            End If
            ' tenemos destino
            TipoDestino = destino.DestTipoDestino
            Dim sourceMascara As String = ""
            Dim source As String
            ' Discutible el destino puede ser DirectorioGestor o el de ejecución.

            Dim TargetMascara As String = TargetDirectorio & "\{0}.FXZ"
            Dim target As String = ""
            Dim carpeta As String = ""

            If chkTengoTodo.IsChecked = False Then ' si no lo tengo, lo copio desde el repositirio

                ' aqui es el código que copia desde el repositirio al directorio del XML
                Select Case TipoDestino
                    Case "SHARED_DRIVE"
                        sourceMascara = destino.DestSharedDriveNombre & "\{0}.fxz"
                        If destino.DestSharedDriveNombre <> TargetDirectorio Then ' solo copio y borro si origen y dest son diferentes

                            Try
                                ' solo traspaso borro si source <> target

                                For Each fila In ListaCarpetasTraducir.tCarpetasTraducir.Rows
                                    carpeta = fila("carpeta")
                                    target = String.Format(TargetMascara, carpeta)
                                    source = String.Format(sourceMascara, carpeta)
                                    If Not File.Exists(source) Then ' 
                                        auxS = nl & String.Format("File {0} does not exist.", source) & nl
                                        auxS &= "The missing file should have been placed in the common area. But it not." & nl
                                        auxS &= "Probably this manifest already has been procesed. " & nl
                                        auxS &= "Process continues but this folder will be ignored" & nl
                                        AnyadetxtSalida(auxS)
                                        MsgBox(auxS, MsgBoxStyle.Information, "Missing folder in common area from translator to PM")
                                    Else
                                        AnyadetxtSalida(String.Format("Coping from {0} to {1}...", source, target))
                                        ' ahora lo copio
                                        File.Copy(source, target, True)
                                        AnyadetxtSalida("Done!" & nl)
                                    End If
                                Next
                                ' ahora borro 
                                For Each fila In ListaCarpetasTraducir.tCarpetasTraducir.Rows
                                    carpeta = fila("carpeta")
                                    source = String.Format(sourceMascara, carpeta)
                                    AnyadetxtSalida(String.Format("Deleting {0}...", source))
                                    ' ahora lo copio
                                    File.Delete(source)
                                    AnyadetxtSalida("Done!" & nl)
                                Next
                                ' last step is delete the manifest from the EA
                                File.Delete(ArchivoMF)
                            Catch ex As Exception
                                auxS = nl & String.Format("Cannot copy or delete file(s) from {0} to {1}", destino.DestSharedDriveNombre, ListaCarpetasTraducir.DirectorioGestor) & nl
                                auxS &= "If the fxz files are in the manifest directory is safe to continue, otherwise, probably ther are missing fxz files in the transaltor staging area. " & nl & nl
                                auxS &= "Do you want to continue?"
                                If MsgBoxResult.No = MsgBox(auxS, MsgBoxStyle.Critical, "Error moving files from translator to pm") Then
                                    AnyadetxtSalida("Cannot copy and/or delete files. User has aborted." & nl)
                                    Exit Sub
                                End If
                            End Try
                        End If
                        ' ahora lo unzipo
                    Case "PATH"
                        sourceMascara = destino.DestPath & "\{0}.fxz"
                        Try
                            For Each fila In ListaCarpetasTraducir.tCarpetasTraducir.Rows
                                carpeta = fila("carpeta")
                                target = String.Format(TargetMascara, carpeta)
                                source = String.Format(sourceMascara, carpeta)
                                AnyadetxtSalida(String.Format("Coping from {0} to {1}...", source, target))
                                ' ahora lo copio
                                File.Copy(source, target, True)
                                AnyadetxtSalida("Done!" & nl)
                            Next
                            ' ahora borro 
                            For Each fila In ListaCarpetasTraducir.tCarpetasTraducir.Rows
                                carpeta = fila("carpeta")
                                source = String.Format(sourceMascara, carpeta)
                                AnyadetxtSalida(String.Format("Deleting {0}...", source))
                                ' ahora lo copio
                                File.Delete(source)
                                AnyadetxtSalida("Done!" & nl)
                            Next
                        Catch ex As Exception
                            auxS = nl & String.Format("Cannot copy or delete file(s) from {0} to {1}", destino.DestPath, ListaCarpetasTraducir.DirectorioGestor) & nl
                            auxS &= "If the fxz files are in the manifest directory is safe to continue, otherwise, probably ther are missing fxz files in the transaltor staging area. " & nl & nl
                            auxS &= "Do you want to continue?"
                            AnyadetxtSalida(auxS)
                            If MsgBoxResult.No = MsgBox(auxS, MsgBoxStyle.YesNo, "Error moving files from translator to pm") Then
                                AnyadetxtSalida("Cannot copy and/or delete files. User has aborted." & nl)
                                Exit Sub
                            End If
                        End Try
                        ' ahora lo unzipo
                    Case "FTP", "FTPES"
                        Dim TIPO_FTP = TipoDestino
                        ServicePointManager.ServerCertificateValidationCallback =
                        Function(se As Object,
                        cert As System.Security.Cryptography.X509Certificates.X509Certificate,
                        chain As System.Security.Cryptography.X509Certificates.X509Chain,
                        sslerror As System.Net.Security.SslPolicyErrors) True
                        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

                        ' tengo que bajar los fxz del repositorio a 
                        For Each fila In ListaCarpetasTraducir.tCarpetasTraducir.Rows
                            carpeta = fila("carpeta")
                            source = String.Format(sourceMascara, carpeta)
                            If Not (destino.DestFTPNombre.EndsWith("/")) Then destino.DestFTPNombre &= "/"
                            ' Archivo FTP, voy a conectarme al servidor FTP para leer el directorio
                            Dim ftpURI As String
                            ftpURI = "ftp://" & destino.DestFTPHost & "/" & destino.DestFTPNombre ' si no añado / al final me sale el subdir
                            ftpURI &= carpeta & ".FXZ"
                            Dim contrasenya As String = DecryptWithKey(destino.DestFTPContrasenya, ParLongInterno)
                            Dim ftpCrendenciales As NetworkCredential = New System.Net.NetworkCredential(destino.DestFTPUsuario, contrasenya)
                            Dim FTPSolicitud As FtpWebRequest
                            Dim FTPRespuesta As FtpWebResponse
                            Dim StreamFTPRespuesta As Stream
                            Dim ReaderStreamFTPRespuesta As StreamReader
                            Try
                                ' ahora lo bajaré al directorio temporal si tengo uno
                                'ftpURI = "ftp://" & par.DestFTPHost & "/" & par.DestFTPNombre & "/" ' si no añado / al final me sale el subdir
                                Dim archivolocal As String = String.Format(TargetMascara, carpeta)
                                FTPSolicitud = CType(System.Net.FtpWebRequest.Create(ftpURI), System.Net.FtpWebRequest)
                                FTPSolicitud.Credentials = ftpCrendenciales
                                FTPSolicitud.KeepAlive = True
                                FTPSolicitud.UseBinary = True
                                FTPSolicitud.Method = System.Net.WebRequestMethods.Ftp.DownloadFile
                                FTPSolicitud.Proxy = Nothing
                                'secure
                                If TIPO_FTP = "FTPES" Then
                                    FTPSolicitud.EnableSsl = True
                                    FTPSolicitud.KeepAlive = False
                                End If
                                ' end secure
                                AnyadetxtSalida(String.Format("Downloading to -> {0}", archivolocal) & nl)
                                Using response As System.Net.FtpWebResponse = CType(FTPSolicitud.GetResponse, System.Net.FtpWebResponse)
                                    Using responseStream As IO.Stream = response.GetResponseStream
                                        'loop to read & write to file
                                        Using fs As New IO.FileStream(archivolocal, IO.FileMode.Create)
                                            Dim kk As Integer = 50
                                            Dim buffer(2047) As Byte
                                            Dim read As Integer = 0
                                            Do
                                                read = responseStream.Read(buffer, 0, buffer.Length)
                                                fs.Write(buffer, 0, read)
                                                kk -= 1 : If kk = 0 Then kk = 50 : AnyadetxtSalida(".")
                                            Loop Until read = 0 'see Note(1)
                                            responseStream.Close()
                                            fs.Flush()
                                            fs.Close()
                                        End Using
                                        responseStream.Close()
                                    End Using
                                    response.Close()
                                End Using
                                Console.WriteLine(" Done!" & nl)
                                ' antes de hacer el rename intento borrarlo si existe.
                                FTPSolicitud = DirectCast(WebRequest.Create(New Uri(ftpURI)), FtpWebRequest)
                                FTPSolicitud.Credentials = ftpCrendenciales
                                FTPSolicitud.Method = WebRequestMethods.Ftp.DeleteFile
                                'secure
                                If TIPO_FTP = "FTPES" Then
                                    FTPSolicitud.EnableSsl = True
                                    FTPSolicitud.KeepAlive = False
                                End If
                                ' end secure

                                Dim respuesta = DirectCast(FTPSolicitud.GetResponse(), FtpWebResponse)


                            Catch ex As Exception
                                AnyadetxtSalida("FTP ERROR *.PaT.zip " & nl)
                                AnyadetxtSalida(ex.Message)
                                If Not (ex.InnerException Is Nothing) Then
                                    AnyadetxtSalida(ex.InnerException.Message)
                                End If
                                Exit Sub
                            End Try


                            AnyadetxtSalida(String.Format("Deleting {0}...", source))
                            ' ahora lo copio
                            'File.Delete(source)
                            AnyadetxtSalida("Done!" & nl)
                        Next
                    Case "SFTP"
                        Dim TIPO_FTP = TipoDestino
                        ' sftp info
                        Dim Contrasenya As String = DecryptWithKey(destino.DestFTPContrasenya, ParLongInterno)
                        Dim posicolon As Integer = InStr(destino.DestFTPHost, ":")
                        Dim sftpport As Integer = 22
                        Dim sftphost As String = destino.DestFTPHost
                        If posicolon <> 0 Then
                            sftpport = CType(sftphost.Substring(posicolon, Len(sftphost) - posicolon), Integer)
                            sftphost = sftphost.Substring(0, posicolon - 1)
                        End If


                        ' tengo que bajar los fxz del repositorio a 
                        For Each fila In ListaCarpetasTraducir.tCarpetasTraducir.Rows
                            carpeta = fila("carpeta")
                            source = String.Format(sourceMascara, carpeta)
                            Dim ftpfile = destino.DestFTPNombre
                            If Not ftpfile.EndsWith("/") Then ftpfile &= "/"
                            ftpfile &= carpeta & ".FXZ"
                            Dim archivolocal As String = String.Format(TargetMascara, carpeta)
                            Dim SFTPDownloadFile As strucSFTPDownloadFile = New strucSFTPDownloadFile(
                               sftphost, sftpport, destino.DestFTPUsuario, Contrasenya, archivolocal, ftpfile)
                            Try
                                AnyadetxtSalida(String.Format("File {0}", ftpfile) & nl)
                                AnyadetxtSalida(String.Format("Downloading to -> {0}", archivolocal) & nl)
                                Dim tsk As Task(Of strucSFTPDownloadFile) = Task.Run(Function() DownloadMyFile(SFTPDownloadFile))
                                While Not tsk.IsCompleted
                                    Threading.Thread.Sleep(1000) ' cada segundo
                                    AnyadetxtSalida(".")
                                End While
                                Console.WriteLine(nl & " Done!" & nl)
                                If Not SFTPDownloadFile.todoOk Then
                                    auxS = "Fatal error cannot download PAT file (inner loop): " & nl
                                    auxS &= "source -> " & SFTPDownloadFile.ftpfile & nl
                                    auxS &= "target -> " & SFTPDownloadFile.localfile & nl
                                    auxS &= "Info   -> " & SFTPDownloadFile.salida & nl
                                    AnyadetxtSalida(auxS)
                                    auxi = MsgBox(auxS, MsgBoxStyle.Critical, "OMG sftp download PAT error (inner loop)...")
                                    Exit Sub
                                End If
                                AnyadetxtSalida("Done! File download from sftp server. Now we will try to delete it." & nl)
                                '
                                ' now we delete the folder
                                '
                                Dim SFTPDeleteFile As strucSFTPDeleteFile = New strucSFTPDeleteFile(
                                sftphost, sftpport, destino.DestFTPUsuario, Contrasenya, ftpfile)

                                SFTPDeleteFile = DeleteMyFile(SFTPDeleteFile)

                                If Not SFTPDeleteFile.todoOk Then
                                    auxS = "Fatal error cannot delete ftpdir: " & nl
                                    auxS &= ftpfile & nl
                                    auxS &= "info -> " & SFTPDeleteFile.salida & nl
                                    AnyadetxtSalida(auxS)
                                    auxi = MsgBox(auxS, MsgBoxStyle.Critical, "OMG sftp delete PAT error ...")
                                    Exit Sub
                                End If

                                auxS = "File deleted: " & ftpfile & nl
                                AnyadetxtSalida(auxS)



                            Catch ex As Exception
                                auxS = "Fatal error cannot download or detelete PAT file (outer loop): " & nl
                                auxS &= "local file -> " & SFTPDownloadFile.ftpfile & nl
                                auxS &= "sftp file -> " & SFTPDownloadFile.localfile & nl
                                AnyadetxtSalida(auxS)
                                auxi = MsgBox(auxS, MsgBoxStyle.Critical, "OMG sftp download PAT/delete error (outer loop)...")
                                Exit Sub
                            End Try

                        Next

                    Case Else
                        auxS = "Fatal error internal derror " & nl
                        auxS = String.Format("Target type {0} not supported.", TipoDestino)
                        AnyadetxtSalida(auxS)
                        auxi = MsgBox(auxS, MsgBoxStyle.Critical, "OMG Z^Z-Pat internal error ...")
                        Exit Sub

                End Select
                ' ahora unzipo los fzx
                '
                Dim DirEjeucion As String = ""
                DirEjeucion = System.Reflection.Assembly.GetExecutingAssembly().Location ' ubicación ejecutable
                DirEjeucion = Path.GetDirectoryName(DirEjeucion)
                ' desempaqueto
                Dim PreZip = "unzip.exe" ' recordar el directorio de ejecución es al llamar al mandato zip
                Dim mdto As String = ""
                Dim rc As Integer = 0
                For Each fila In ListaCarpetasTraducir.tCarpetasTraducir.Rows
                    carpeta = fila("carpeta")
                    target = String.Format(TargetMascara, carpeta)
                    mdto = PreZip & String.Format(" -o ""{0}"" -d ""{1}"" ", target, TargetDirectorio)
                    AnyadetxtSalida(String.Format("Unziping {0} in {1}...", target, ListaCarpetasTraducir.DirectorioGestor))
                    mandatoZip(mdto, rc, , False)
                    If rc <> 0 Then
                        auxS = String.Format("Unzip command did not work. ('{0}' exited with rc {1}). ", mdto, rc) & nl & nl
                        auxS &= "Check for return codes in http://www.info-zip.org/mans/unzip.html#DIAGNOSTICS . Also check dir paths. " & nl
                        AnyadetxtSalida(auxS)
                        MsgBox(auxS, MsgBoxStyle.Critical, "OMG unzip error!")
                        AnyadetxtSalida("Fatal error!")
                        Exit Sub
                    End If
                    AnyadetxtSalida("Done!" & nl)
                Next
            End If


continua_sin_traductor:
        End If

        Dim reemplazarcarpetas As Boolean = chkImportar.IsChecked ' Import folders in current OpenTM2 environment
        ' voy a ver que tengo al menos las carpetas
        Dim row As DataRow : Dim auxB As Boolean = True ' en principio suponemos que todo ok.
        Dim unidad As String = Path.GetPathRoot(ParDirTemporal).Substring(0, 1) ' la letra
        auxS = Path.GetPathRoot(ParDirTemporal) ' C:\
        auxS = auxS.Replace("\", "") ' C: que es lo que debo quitar
        Dim DirTemporalSinUnidad As String = ParDirTemporal.Replace(auxS, "")
        Dim mandato As String = "" : Dim opcion As String = ""
        For Each row In ListaCarpetasTraducir.tCarpetasTraducir.Rows
            Dim carpeta As String = row("Carpeta")
            Dim carpetaConPath = TargetDirectorio & String.Format("\{0}.FXP", carpeta)
            If File.Exists(carpetaConPath) = False Then
                txtSalida.AppendText("Folder not found -> " & carpetaConPath & nl)
            Else
                'txtSalida.AppendText("Folder -> " & carpetaConPath & nl)
                AnyadetxtSalida("Folder -> " & carpetaConPath & nl)
                Dim bCNT As Boolean = row("bCNT")
                Dim bIniCal As Boolean = row("bIniCal")
                Dim bFinCal As Boolean = row("bFinCal")
                If bFinCal = False Then ' si es true no hago nada, si es false, tenog que calcularlo todo
                    Dim rc As Integer
                    ' bueno, necesitaré el FinCal
                    ' importaré la carpeta como 
                    Dim carpetaWCT As String = ""
                    Dim ArchivoFinCal As String = ""
                    Dim unidadtarget = Path.GetPathRoot(TargetDirectorio).Substring(0, 1)
                    If reemplazarcarpetas Then ' importar en el entorno de opentm2 (lo habitual)
                        Dim DirectorioCopiaSeguridad As String = TargetDirectorio & "\01_OTMBackup"
                        Dim DirectorioCopiaSeguridadSinUnidad As String =
                            Path.DirectorySeparatorChar & DirectorioCopiaSeguridad.Replace(Path.GetPathRoot(DirectorioCopiaSeguridad), "")
                        ' lo primero que voy a hacer es una copia de seguridad.
                        Try
                            ' Voy a crear los directorios temporales 
                            If (Not System.IO.Directory.Exists(DirectorioCopiaSeguridad)) Then
                                System.IO.Directory.CreateDirectory(DirectorioCopiaSeguridad)
                                txtSalida.AppendText(String.Format("Created temporary dir {0} for before importing folders.", DirectorioCopiaSeguridad) & nl)
                            End If
                            ' ahora exporta la carpeta
                            mandato = " /TAsk=FLDEXP /FLD={0} /TOdrive={1} /ToPath={2} /OPtions=(MEM,ROMEM,DOCMEM) /OVerwrite=NO /QUIET=NOMSG  "
                            mandato = String.Format(mandato, carpeta, unidadtarget, DirectorioCopiaSeguridadSinUnidad)
                            'AnyadeSalida(mdto)
                            ejecuta(mandato, rc)
                            If rc = 0 Then ' mandato ok, voy a comprimir la carpeta
                                AnyadetxtSalida("Folder " & carpeta & " exported in " & DirectorioCopiaSeguridad)
                                ' ahora el zip
                                Dim Prezip = "zip.exe"
                                auxS = Prezip & " -j " & String.Format("""{0}\{1}"" ""{2}\{3}"" ",
                                                     DirectorioCopiaSeguridad,
                                                     carpeta & ".fxz",
                                                     DirectorioCopiaSeguridad,
                                                     carpeta & ".FXP")

                                mandatoZip(auxS, rc, , False) ' DirEjecución y debuga son opcionales
                                ' lo borro
                                File.Delete(DirectorioCopiaSeguridad & "\" & carpeta & ".FXP")

                                '

                            ElseIf rc = 151 Then ' 
                                AnyadetxtSalida(" -> Folder does not exist in current OpenTM2 environment. No backup.")
                            ElseIf rc <> 0 Then ' cualquier otro error
                                MsgBox("Folder to be updtated in OTM2 cannot be exported as backup in the backup directory " _
                                       & DirectorioCopiaSeguridad & nl &
                                     " Maybe there the folder already has been exported in the backup directory. " &
                                     " RC=999 Indicates a problem in the shell environment.  " &
                                     "RC=" & rc.ToString, MsgBoxStyle.Critical,
                                     "Cannot export the folder to be updated backup")
                                Exit Sub
                            End If
                            AnyadetxtSalida(" -> Importing updating folder")
                            mandato = " /TAsk=FLDIMP  /FLD={0} /FromPath={1} /FROMDRIVE={2} /OPtions=MEM /QUIET=NOMSG  "
                            mandato = String.Format(mandato, carpeta, TargetDirectorioSinUnidad, unidadtarget)
                            ' AnyadeSalida(mdto)
                            ejecuta(mandato, rc)
                            If rc <> 0 Then
                                MsgBox(String.Format("Cannot import  updating folder {0} {1} . RC={2}",
                                                     TargetDirectorioSinUnidad, carpeta, rc.ToString), MsgBoxStyle.Critical, "Updating folder cannot be imported")
                                Exit Sub
                            End If
                            ' analizo
                            AnyadetxtSalida(" -> Folder to be udpated has been merged with the updating folder) ")

                            ' ahora el calculating final
                            opcion = "/PROFILE=" & row("Perfil")
                            ArchivoFinCal = TargetDirectorio & String.Format("\{0}_{1}_fin_cal.rpt", carpeta, row("Perfil"))
                            mandato = "/TAsk=CNTRPT /FLD={0} /OUT={1} /RE=CALCULATING /TYPE=BASE_SUMMARY_FACT {2} /OV=YES /QUIET=NOMSG"
                            mandato = String.Format(mandato, carpeta, ArchivoFinCal, opcion)
                            ejecuta(mandato, rc) ' calcula calculating final
                            ' Shell(mandato, AppWinStyle.MinimizedFocus, True, -1)
                            If rc <> 0 Then
                                'Dim debuginfo As String = "" :If Debug Then debuginfo = vbCrLf & mandatoT
                                auxS = "Fatal error: Cannot run the final calculating. RC=" & rc.ToString
                                txtSalida.AppendText(auxS & nl)
                                MsgBox(auxS, MsgBoxStyle.Critical, "Cannot run final calculating report")
                                Exit Sub
                            End If
                            AnyadetxtSalida(String.Format(" -> Folder calculating done ({0}) " & nl, row("Perfil")))


                            If chkContajeAdicional.IsChecked Then ' tengo que cerrar la carpeta con contaje de IBM
                                ' esta parte es desde luego hardcode
                                Dim curretPrf As String = row("Perfil")
                                Select Case curretPrf
                                    Case "PII20184"
                                    Case "PUB20184"
                                    Case Else
                                        If Len(curretPrf) >= 3 Then
                                            Dim nperfil As String = ""
                                            Dim TresLetasPerfil As String = curretPrf.Substring(0, 3)
                                            If TresLetasPerfil = "PII" Then nperfil = "PII20184"
                                            If TresLetasPerfil = "PUB" Then nperfil = "PUB20184"
                                            If nperfil <> "" Then ' tengo que calcularo para el perfil de IBM
                                                ' ahora el calculating final
                                                opcion = "/PROFILE=" & nperfil
                                                Dim ArchivoFinCalIBM As String = "" ' originalmente ArchivoFinCal pero es que se usa para la PO
                                                ArchivoFinCalIBM = TargetDirectorio & String.Format("\{0}_{1}_fin_cal.rpt", carpeta, nperfil)
                                                mandato = "/TAsk=CNTRPT /FLD={0} /OUT={1} /RE=CALCULATING /TYPE=BASE_SUMMARY_FACT {2} /OV=YES /QUIET=NOMSG"
                                                mandato = String.Format(mandato, carpeta, ArchivoFinCalIBM, opcion)
                                                ejecuta(mandato, rc) ' calcula calculating final
                                                ' Shell(mandato, AppWinStyle.MinimizedFocus, True, -1)
                                                If rc <> 0 Then
                                                    'Dim debuginfo As String = "" :If Debug Then debuginfo = vbCrLf & mandatoT
                                                    auxS = "Error: Cannot run the final calculating for IBM profile. RC=" & rc.ToString
                                                    txtSalida.AppendText(auxS & nl)
                                                    MsgBox(auxS, MsgBoxStyle.Information, "Cannot run final IBM calculating report. Process will continue.")
                                                End If
                                                AnyadetxtSalida(String.Format(" -> Folder calculating done ({0}) " & nl, nperfil))
                                            End If
                                        End If
                                End Select

                            End If



                        Catch ex As Exception
                            auxS = "Fatal error: " & ex.Message.ToString
                            txtSalida.AppendText(auxS & nl)
                            Exit Sub

                        End Try

                        ' tengo que importar las carpetas 

                    Else ' obtener contaje SIN sustituir carpetas
                        ' primero suprmiré la carpeta por si acaso
                        carpetaWCT = carpeta & "_WCT" '
                        mandato = "/TAsk=FLDDEL /FLD={0} /QUIET=NOMSG"
                        mandato = String.Format(mandato, carpetaWCT)
                        ejecuta(mandato, rc)
                        If rc <> 0 Then
                            'Dim debuginfo As String = "" : If Debug Then debuginfo = vbCrLf & mandato
                            'MsgBox("No se puede suprimir carpeta temporal. RC=" & rc.ToString & debuginfo)
                        End If
                        Dim archivoCarpetaWCT As String = ParDirTemporal & carpetaWCT & ".FXP"
                        File.Copy(carpetaConPath, archivoCarpetaWCT, True)  ' copio xxx_WCT a dir temporal
                        opcion = "/OPTIONS=()" ' importo sin memorias
                        mandato = "/TAsk=FLDIMP /FLD={0} /FROMDRIVE={1} /FROMPATH={2} /TODRIVE=C {3}  /QUIET=NOMSG"
                        mandato = String.Format(mandato, carpetaWCT, unidad, DirTemporalSinUnidad, opcion)
                        ejecuta(mandato, rc)
                        If rc <> 0 Then
                            'Dim debuginfo As String = "" :If Debug Then debuginfo = vbCrLf & mandatoT
                            auxS = "Fatal error:Cannot delete temporary folder. RC=" & rc.ToString
                            txtSalida.AppendText(auxS & rc.ToString & nl)
                            Exit Sub
                        End If
                        ' ahora suprimo la carpeta temporal
                        File.Delete(archivoCarpetaWCT)
                        ' ahora el calculating final
                        opcion = "/PROFILE=" & row("Perfil")
                        ArchivoFinCal = TargetDirectorio & String.Format("\{0}_{1}_fin_cal.rpt", carpeta, row("Perfil"))
                        mandato = "/TAsk=CNTRPT /FLD={0} /OUT={1} /RE=CALCULATING /TYPE=BASE_SUMMARY_FACT {2} /OV=YES /QUIET=NOMSG"
                        mandato = String.Format(mandato, carpetaWCT, ArchivoFinCal, opcion)
                        ejecuta(mandato, rc) ' calcula calculating final
                        ' Shell(mandato, AppWinStyle.MinimizedFocus, True, -1)
                        If rc <> 0 Then
                            'Dim debuginfo As String = "" :If Debug Then debuginfo = vbCrLf & mandatoT
                            auxS = "Fatal error: Cannot run the final calculating. RC=" & rc.ToString
                            txtSalida.AppendText(auxS & nl)
                            MsgBox(auxS, MsgBoxStyle.Critical, "Cannot run final calculating report")
                            Exit Sub
                        End If
                        ' borro la carpeta temporal
                        mandato = "/TAsk=FLDDEL /FLD={0} /QUIET=NOMSG"
                        mandato = String.Format(mandato, carpetaWCT)
                        ejecuta(mandato, rc)
                        If rc <> 0 Then
                            'Dim debuginfo As String = "" : If Debug Then debuginfo = vbCrLf & mandato
                            'MsgBox("No se puede suprimir carpeta temporal. RC=" & rc.ToString & debuginfo)
                        End If

                    End If

                    ' al llegar aquí ArchivoFinCal tiene el calculating final
                    ' voy a saco

                    Dim lines As String() = IO.File.ReadAllLines(ArchivoFinCal)
                    txtSalida.AppendText("ArchivoFinCal " & ArchivoFinCal & nl)
                    Dim ultimalinea = lines(lines.Length - 3) ' Payable words  : 235.78
                    lines = Nothing
                    Dim ValColSal() As String = Split(ultimalinea)
                    ' debo eliminar los blancos
                    Dim LastNonEmpty As Integer = -1
                    For i As Integer = 0 To ValColSal.Length - 1
                        If ValColSal(i) <> "" Then
                            LastNonEmpty += 1
                            ValColSal(LastNonEmpty) = ValColSal(i)
                        End If
                    Next
                    ReDim Preserve ValColSal(LastNonEmpty)
                    ' ValColSal(2) será 1234.24 debo pasarlo a entero

                    ' Payable 0 Words 1 : 2 mn 3
                    row("FinCal") = Int(Val(ValColSal(3)))
                    row("bFinCal") = True
                    'txtSalida.AppendText("Contaje FIN-> " & row("FinCal") & nl)
                    txtSalida.AppendText("FINal calculating -> " & row("FinCal") & nl)
                    Dim Ci As Integer = 0 ' en principio el CI es cero
                    If row("bIniCal") Then
                        Ci = row("IniCal")
                        txtSalida.AppendText("INItial calculating -> " & Ci & nl)
                    End If
                    row("Contaje") = row("FinCal") - Ci
                    txtSalida.AppendText("Folder wordcount-> " & row("Contaje") & nl)
                    ' añado el total
                    row("bTotal") = True ' OK
                    Dim suma As Single
                    suma = row("Contaje") * row("Tarifa")
                    row("Total") = Math.Round(suma, 2, MidpointRounding.ToEven)
                End If


            End If    ' 
        Next

        ' 
        ' Visualiza DG 
        '
        dgCarpetas.ItemsSource = ListaCarpetasTraducir.tCarpetasTraducir.DefaultView
        dgCarpetas.IsReadOnly = True
        '
        '
        ' tengo que gernar la orden de compra

        Dim archivoOC As String = TargetDirectorio & "\" & ListaCarpetasTraducir.NombrePaT.Replace(".PaT", "") & "_PO.HTM"
        ' La Orden de compra se genera con 2 llamadas
        If ListaCarpetasTraducir.GeneraOC_HTML(archivoOC, "") = False Then
            txtSalida.AppendText("Fatal error. We have not been able to generate the purchase order (1st Part)" & nl)
            Exit Sub
        Else
            txtSalida.AppendText("Generated 1st part purchase order!" & nl)

        End If
        If ListaCarpetasTraducir.GeneraOC_HTML(archivoOC, "FAC") = False Then
            txtSalida.AppendText("Fatal error. We have not been able to generate the purchase order (2nd Part)" & nl)
        Else
            txtSalida.AppendText("Generated 2nd part purchase order!" & nl)
        End If
        ' ahora la última parte, el correo al traductor para que pueda facturar.
        If chkCorreo.IsChecked And NombreTraductor <> "NONE" Then ' " Then
            Dim destinatario As String = destino.DestCorreo
            AnyadetxtSalida(String.Format("Sending note to {0}... ", destino.DestCorreo & nl))
            Dim mailssl As Boolean = ParMailSSL ' True
            Dim mailpuerto As Integer = ParMailPuerto '
            Dim mailcontrase As String = DecryptWithKey(ParMailContrase, ParLongInterno) ' "por ejemplo manolo"
            Dim mailhost As String = ParMailHost ' "smtp.xxxx.com"
            Dim mailusuario As String = ParMailUsuario ' "user@gmail.com"
            Dim client As New SmtpClient
            client.DeliveryMethod = SmtpDeliveryMethod.Network
            client.EnableSsl = mailssl
            client.Host = mailhost
            client.Port = mailpuerto
            ' autenticaciÃ³n
            Dim credentials As New System.Net.NetworkCredential(mailusuario, mailcontrase)

            client.UseDefaultCredentials = False
            client.Credentials = credentials

            Dim msg As New MailMessage
            msg.From = New MailAddress(mailusuario)
            msg.To.Add(New MailAddress(destinatario))
            Dim RtL As New List(Of MailAddress)
            'msg.ReplyTo = New MailAddress("miguelcanals@hotmail.com")
            msg.Subject = String.Format("PO -> PaT name: {0} / Recipient: {1}", Path.GetFileName(archivoOC), destino.DestNombre)
            msg.IsBodyHtml = "true"
            '
            '
            msg.Body = ""
            Dim lin2() As String = {
                    "<HTML><BODY>",
                    String.Format("PaT Purchase Order. PLEASE SEND YOUR TO {0} INVOICE WITH THIS NOTE.", ListaCarpetasTraducir.CorreoGestor),
                    "Pat recipient: " & destino.DestNombre}
            lin2 = {"<HTML><BODY>",
                    String.Format("PaT Purchase Order. THIS IS FOR YOUR INFORMATION ONLY. This is an optional email you can receive in order to:<br>"),
                    String.Format("1) Receive counting information. Ask you provider about the ACTUAL invoicing process. "),
                    String.Format("2) The PM wants to inform you he/she has download and process the PAT."),
                    String.Format("PM: {0} ", ListaCarpetasTraducir.CorreoGestor),
                    "Pat recipient: " & destino.DestNombre}
            For s = 0 To lin2.Count - 1
                msg.Body &= lin2(s) & "<br>"
            Next
            msg.Body &= "</BODY></HTML>"

            msg.Body &= File.ReadAllText(archivoOC)
            ' log estÃ¡ en 
            'Dim ArchivoLog As String = CargaCte("mantenimiento", "log")
            Dim datos As New Attachment(archivoOC)
            msg.Attachments.Add(datos)

            Try
                client.Send(msg)
                AnyadetxtSalida("Done!" & nl)
            Catch ex As Exception
                AnyadetxtSalida(String.Format("Error sending note: {1}", ex.Message) & nl & "Proces continues..." & NombreTraductor & nl)
            End Try
            msg.Dispose()

        End If

        ' lo último que hago es limpieza
        AnyadetxtSalida("Clean up and renaming folders (only if possible) " & nl)
        Dim traductor As String = ListaCarpetasTraducir.NombreTraductor
        Try
            For Each row In ListaCarpetasTraducir.tCarpetasTraducir.Rows

                Dim fold As String = TargetDirectorio & "\" + row("carpeta") & ".fxp" ' folder
                Dim orig As String = TargetDirectorio & "\" + row("carpeta") & ".fxz" ' ziped folder
                Dim dest As String = row("carpeta") & "_" & NombreTraductor + ".fxz" ' renamed ziped folder
                My.Computer.FileSystem.RenameFile(orig, dest) ' renameing
                File.Delete(fold) ' deleting the folder
            Next
            ''AnyadetxtSalida("Clean up and renaming folders -> Done! " & nl)
        Catch ex As Exception
            AnyadetxtSalida("Clean up and renaming folders could not be complete " & nl)
        End Try

        AnyadetxtSalida("************************  END  ************************ " & nl)
    End Sub




    Function AnyadetxtSalida(s As String)
        txtSalida.AppendText(s)
        txtSalida.SelectionStart = txtSalida.Text.Length
        txtSalida.ScrollToEnd()
        txtSalida.Refresh()
        Return 0
    End Function
End Class
