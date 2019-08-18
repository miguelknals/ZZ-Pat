' como mínimo refresh necesita esto
Imports System.Threading
Imports System.Runtime.CompilerServices
Imports System.Windows.Threading
' archivos
Imports System.IO
' 
Imports System.Xml
Imports System.Xml.Serialization
' 
Imports System.Data
'
Imports System.Net
Imports System.Net.Mail


Class MainWindow
    Dim nl As String = Environment.NewLine
    Dim auxS As String = ""
    Dim dt As DispatcherTimer = New DispatcherTimer()

    Dim debuga As Boolean = True
    Dim TiempoBucle As Integer = 300
    Dim NoenviarCorreo = False ' en producción debería ser false

    ' bucle principal

    Sub New()
        '

        ' This call is required by the designer.
        InitializeComponent()


        ' Add any initialization after the InitializeComponent() call.
        AnyadetxtSalida("Welcome to ZZ-InstallPatWPF... " & nl)
        Dim ViaEjecutable As String = System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase
        ViaEjecutable = ViaEjecutable.Replace("file:///", "")
        'Dim fechaEjecutablee As String = File.GetCreationTime(System.Reflection.Assembly.GetExecutingAssembly().Location)
        VentanaPrincipal.Title = "ZZ-InstallPatWPF... " & File.GetLastWriteTime(ViaEjecutable).ToString & "  (" & ViaEjecutable & ")"
        ' lo anterior cambia al cambiar el archivo a otra máquina
        ' de http://stackoverflow.com/questions/804192/how-can-i-get-the-assembly-last-modified-date
        ' había alguna solución más radical
        Dim lastMod2 As String = File.GetLastWriteTime(System.Reflection.Assembly.GetExecutingAssembly().Location).ToString
        Dim arch As FileInfo = New System.IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location)
        Dim lastMod As DateTime = arch.LastWriteTime
        VentanaPrincipal.Title = "ZZ-InstallPatWPF... " & lastMod & "/" & lastMod2 & "  (" & ViaEjecutable & ")"
        ' voy a buscar el path de OTM
        Dim ViaSistema = Environment.GetEnvironmentVariable("PATH")
        'voy a buscar C:\OTM\WIN
        Dim auxi As Integer = InStr(ViaSistema, ":\OTM\WIN")
        ' 

        ' Cargo los parámetros
        Dim Par As structParametros
        Par = CargaParametros()
        ' 
        cargalbxWIP(Par.DestLocalPatDir)


        cargalbxDONE(Par.DestLocalPatDir)

        ' creamos el evento y el manejo del error
        ' https://msdn.microsoft.com/en-us/library/system.windows.threading.dispatchertimer%28v=vs.100%29.aspx?cs-save-lang=1&cs-lang=vb#code-snippet-2
        ' Dim dt As DispatcherTimer = New DispatcherTimer() creado localmente
        AddHandler dt.Tick, AddressOf dispatcherTimer_Tick
        'dt.Interval = New TimeSpan(0, 0, TiempoBucle) ' tiempo bucle variable global 
        dt.Interval = New TimeSpan(0, 0, 1) ' tiempo bucle variable global 
        'dt.Stop()
        'IntentaProcesar() 'Intant procesar activa el timer después de la primera activación
        AnyadetxtSalida("***************************" & nl)
        ' AnyadetxtSalida(String.Format("Entering in the {0} seconds main loop...", TiempoBucle) & nl)
        If NoenviarCorreo Then AnyadetxtSalida("WARNING: Debug setting Do Not Send emails = true..." & nl)
        ' ahora el timer
        'IntentaProcesar()
        ' dt.Start() Esto inicia el temporizador



    End Sub
    Public Sub dispatcherTimer_Tick(ByVal sender As Object, ByVal e As EventArgs)
        CommandManager.InvalidateRequerySuggested()
        dt.Stop() : AnyadetxtSalida("Stoping timer." & nl)
        IntentaProcesar()

        ' dt.Start() : AnyadetxtSalida("Restarting timer." & nl)

        ' si hay que parar el temporizador
        ' http://stackoverflow.com/questions/4063433/dispatchertimer-tick-once
        Dim auxS As String = "Inspection has finished... "
        auxS &= DateTime.Now.Hour.ToString() & ":" & DateTime.Now.Minute.ToString() & ":" & DateTime.Now.Second.ToString()
        auxS &= String.Format(" ... Waiting... {0} seconds more... ", TiempoBucle) & nl
        AnyadetxtSalida(auxS)



    End Sub
    Sub IntentaProcesar()
        ' estas variables quizás podrían ser globales en cuyo caso habría que reiniciarlaslas ahora
        dt.Interval = New TimeSpan(0, 0, TiempoBucle) ' tiempo bucle variable global 
        AnyadetxtSalida("Entering ItentaProcesar" & nl)
        Dim NombrePaTconPath As String
        Dim NombrePatsinExtension As String
        Dim NombrePat As String
        Dim DirOutPat As String
        Dim DirOutPatconPat As String ' 
        Dim DirTMP As String
        Dim DirEjeucion As String
        ' 
        NombrePaTconPath = ""
        NombrePatsinExtension = ""
        NombrePat = ""
        DirOutPat = ""
        DirOutPatconPat = ""
        DirTMP = ""
        DirEjeucion = ""
        '
        ' directorio temporal
        DirTMP = Path.GetTempPath() & "ZZ\"
        AnyadetxtSalida(String.Format("Temporary dir {0}", DirTMP))
        ' voy a crearlo si no existe
        If Not Directory.Exists(DirTMP) Then ' no existe lo voy a intentar crear
            Try
                AnyadetxtSalida(String.Format("Dir {0} does not exist. We will try to create it... ", DirTMP))
                Directory.CreateDirectory(DirTMP)
                AnyadetxtSalida("Done!" & nl)
            Catch ex As Exception
                auxS = nl & "Ops! Temporary PaT Dir {0} cannot be created. You need to fix it in order to process this PaT" & nl
                AnyadetxtSalida(auxS)
                MsgBox(auxS, MsgBoxStyle.Exclamation, "Cannot create temporary Pat directory")
                Exit Sub
            End Try
        Else ' existe
            ' voy a borrar todo lo que haya
            For Each archivo In My.Computer.FileSystem.GetFiles(DirTMP, FileIO.SearchOption.SearchTopLevelOnly, "*.*")
                My.Computer.FileSystem.DeleteFile(archivo)
            Next
        End If


        ' auxiliar
        Dim auxI As Integer = 0
        ' aqui es donde realmente realizo el proceso
        AnyadetxtSalida("Let's inspect the PaT directory..." & nl)
        AnyadetxtSalida("Loading parameters..." & nl)
        Dim par As structParametros
        par = CargaParametros()
        ' tengo que actuar según el tipo de destino
        Dim TipoDeposito As String = par.DestTipoDestino
        ' tengo que obtner la lista
        Dim DirTarget As String = ""
        Select Case TipoDeposito
            Case "PATH", "SHARED_DRIVE" ' creo que es lo mismo
                If TipoDeposito = "SHARED_DRIVE" Then
                    DirTarget = par.DestSharedDriveNombre
                    AnyadetxtSalida("Share drive directory " & DirTarget & nl)
                ElseIf TipoDeposito = "PATH" Then
                    DirTarget = par.DestPath ' path
                    AnyadetxtSalida("Standard directory " & DirTarget & nl)
                End If
                Dim archivos() As String
                Try
                    archivos = Directory.GetFiles(DirTarget, "*.PaT.zip")
                Catch ex As Exception
                    auxS = String.Format("Looks like target file '{0}' for zip pat files is not valid.", DirTarget) & nl
                    auxS &= "Verify you parameters. "
                    MsgBox(auxS, MsgBoxStyle.Critical, "Target directory for zip pat files is not correct")
                    Exit Sub
                End Try

                For Each s In archivos
                    NombrePaTconPath = Trim(s) : Exit For ' a veces sin trim añade espacio en blanco
                Next
                If NombrePaTconPath = "" Then Exit Select ' no hay sago
                ' hay un archivo, pero debo ver si está listo para cocinarlo
                ' tengo que asegurarme que se ha acabado de copiar
                Dim ArchivoListo As Boolean = False
                While Not ArchivoListo

                    Try
                        Dim fs As FileStream
                        fs = File.Open(NombrePaTconPath, FileMode.Open, FileAccess.Read, FileShare.None)
                        fs.Close()
                        ArchivoListo = True ' el archivo está disponible 
                    Catch ex As IOException '  UnauthorizedAccessException no la atraba.
                        '
                        '' quizás están copiando
                        ' laika
                        System.Threading.Thread.Sleep(1000) ' espero un segundo
                        auxS = "Looks like the file is still being copied. " & nl
                        auxS &= "Do you want to retry? If you say 'No' you will return to the main loop." & nl
                        AnyadetxtSalida(auxS)
                        auxI = MsgBox(auxS, MsgBoxStyle.YesNo, "OMG file in use...")
                        If auxI = MsgBoxResult.No Then
                            AnyadetxtSalida("File in use. User has selected return to main loop. " & nl)

                            Exit Sub
                        End If
                    Finally

                    End Try
                End While
            Case "FTP"
                ' Archivo FTP, voy a conectarme al servidor FTP para leer el directorio
                NombrePaTconPath = "" ' supongamos que no hay
                Dim ftpURI As String
                ftpURI = "ftp://" & par.DestFTPHost & "/" & par.DestFTPNombre & "/" ' si no añado / al final me sale el subdir
                Dim contrasenya As String = DecryptWithKey(par.DestFTPContrasenya, par.DestLongInterno)
                Dim ftpCrendenciales As NetworkCredential = New System.Net.NetworkCredential(par.DestFTPUsuario, contrasenya)
                Dim FTPSolicitud As FtpWebRequest
                Dim FTPRespuesta As FtpWebResponse
                Dim StreamFTPRespuesta As Stream
                Dim ReaderStreamFTPRespuesta As StreamReader
                'Dim lista As New List(Of String) ' aqui la lista de ".XLSX"
                Dim auxS As String = ""

                Try
                    FTPSolicitud = CType(System.Net.FtpWebRequest.Create(ftpURI), System.Net.FtpWebRequest)
                    FTPSolicitud.Credentials = ftpCrendenciales
                    'FTPSolicitud.Method = WebRequestMethods.Ftp.ListDirectoryDetails
                    FTPSolicitud.Method = WebRequestMethods.Ftp.ListDirectory
                    FTPRespuesta = CType(FTPSolicitud.GetResponse, FtpWebResponse)
                    StreamFTPRespuesta = FTPRespuesta.GetResponseStream
                    AnyadetxtSalida("We will read ftp directory..." & nl)
                    ReaderStreamFTPRespuesta = New StreamReader(StreamFTPRespuesta)
                    While ReaderStreamFTPRespuesta.Peek >= 0
                        auxS = ReaderStreamFTPRespuesta.ReadLine
                        If UCase(auxS).EndsWith(".PAT.ZIP") Then
                            ' lista.Add(auxS) : Exit While
                            auxS = Path.GetFileName(auxS) ' por si sale con el path
                            NombrePaTconPath = auxS : Exit While
                        End If
                    End While
                    'Console.WriteLine("Directory List Complete, status {0}", FTPRespuesta.StatusDescription)
                    AnyadetxtSalida(String.Format("Done (status {0}", FTPRespuesta.StatusDescription) & nl)
                    ReaderStreamFTPRespuesta.Close()
                    FTPRespuesta.Close()
                    ' ahora lo bajaré al directorio temporal si tengo uno
                    If NombrePaTconPath <> "" Then
                        'ftpURI = "ftp://" & par.DestFTPHost & "/" & par.DestFTPNombre & "/" ' si no añado / al final me sale el subdir
                        ftpURI &= auxS ' ftpuri tiene ahora el nombre del archivo a bajar
                        Dim archivolocal As String = DirTMP & "\" & auxS
                        FTPSolicitud = CType(System.Net.FtpWebRequest.Create(ftpURI), System.Net.FtpWebRequest)
                        FTPSolicitud.Credentials = ftpCrendenciales
                        FTPSolicitud.KeepAlive = True
                        FTPSolicitud.UseBinary = True
                        FTPSolicitud.Method = System.Net.WebRequestMethods.Ftp.DownloadFile
                        FTPSolicitud.Proxy = Nothing
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
                        NombrePaTconPath = archivolocal
                        ' antes de hacer el rename intento borrarlo si existe.
                        'FTPSolicitud = DirectCast(WebRequest.Create(New Uri(ftpURI)), FtpWebRequest)
                        'FTPSolicitud.Credentials = ftpCrendenciales
                        'FTPSolicitud.Method = WebRequestMethods.Ftp.DeleteFile
                        Dim respuesta = DirectCast(FTPSolicitud.GetResponse(), FtpWebResponse)
                    End If


                Catch ex As System.Net.WebException
                    AnyadetxtSalida("FTP ERROR *.PaT.zip " & nl)
                    AnyadetxtSalida(ex.Message)
                    If Not (ex.InnerException Is Nothing) Then
                        AnyadetxtSalida(ex.InnerException.Message)
                    End If
                    Exit Sub
                End Try

            Case Else
                auxS = String.Format("Target type is '{0}'. It should be PATH, SHARED_DRIVE or FTP.", TipoDeposito)
                AnyadetxtSalida(auxS & nl)
                auxS &= nl & "Do you want to exit the program?"

                auxI = MsgBox(auxS, MsgBoxStyle.YesNo, "Do you want to exit the program?")
                AnyadetxtSalida(auxI & nl)
                If auxI = MsgBoxResult.Yes Then
                    Application.Current.Shutdown()
                End If
        End Select
        If NombrePaTconPath = "" Then
            AnyadetxtSalida("There are not PaT file to process")
            Exit Sub
        End If



        AnyadetxtSalida(nl & nl & "**********************************************************" & nl)

        AnyadetxtSalida("  PAT FOUND " & nl)

        AnyadetxtSalida("PaT found: " & NombrePaTconPath & nl)
        ' voy a ver si tengo directorio local (DirOutPat es global)
        DirOutPat = Trim(par.DestLocalPatDir)
        If DirOutPat = "" Then ' 
            auxS = "Local PaT dir is blank. You should specify one in the parameters. " _
                & "You need to fix it in order to process this PaT" & nl & "Returning to the main loop." & nl
            AnyadetxtSalida(auxS)
            MsgBox(auxS, MsgBoxStyle.Exclamation, "Local PaT directory is blank")

            Exit Sub
        End If

        If Not Directory.Exists(DirOutPat) Then ' no existe lo voy a intentar crear
            Try
                AnyadetxtSalida(String.Format("Dir {0} does not exist. We will try to create it... ", DirOutPat))
                Directory.CreateDirectory(DirOutPat)
                AnyadetxtSalida("Done!" & nl)
            Catch ex As Exception
                auxS = nl & "Ops! Local PaT Dir {0} cannot be created. You should specify a correct one in the parameters. You need to fix it in order to process this PaT" _
                     & nl & "Returning to the main loop." & nl
                AnyadetxtSalida(auxS)
                MsgBox(auxS, MsgBoxStyle.Exclamation, "Cannot create Local Pat directory")

                Exit Sub
            End Try
        End If
        '
        ' bien parece que llegó por fin la hora de procesar tengo todo
        ' debería comprobar si OTM está instalado
        Dim InfoOTM As New ClassInfoOTM ' obtiene info de OTM
        If InfoOTM.TodoOK = False Then ' OTM parece no instalado 
            ' para que el splashscreen no se me coma el diálogo
            auxS = String.Format("Error in OTM environment {0}", InfoOTM.Info) & nl
            AnyadetxtSalida(auxS)
            auxS &= "You should exit from the program untill this problem is solved. Otherwise the loop will find de PaT file again." & nl & nl
            auxS &= "Returning to the main loop." & nl
            AnyadetxtSalida(auxS)
            MsgBox(auxS, MsgBoxStyle.Critical, "Error in OTM!")
            Exit Sub
        End If
        Dim DiscoOTM As String = InfoOTM.DiscoOTM
        ' lo primero que haré será ver si la carpetas ya está instalado o no


        ' dejo de verificar si está abierto o no, pq miraré si las carpetas están
        'While OTMEnEjecucion()  ' error fatal no haré nada si está en marcha
        '    auxS = "OTM2 is running. You need to close it. Is it closed?" & nl
        '    auxS &= "Pls, close OTM2 and once is closed answer Retry. If you answer Cancel the search loop will start again. "
        '    AnyadetxtSalida(auxS)
        '    auxI = MsgBox(auxS, MsgBoxStyle.RetryCancel, "Opss OTM2 is running, close it!")
        '    If auxI = MsgBoxResult.Cancel Then Exit Sub ' salgo
        'End While
        ' al ataque por fin
        ' directorio ejecución
        DirEjeucion = System.Reflection.Assembly.GetExecutingAssembly().Location ' ubicación ejecutable
        DirEjeucion = Path.GetDirectoryName(DirEjeucion)

        If InStr(DirEjeucion, " ") Then ' ahora mismo limitación
            auxS = "Currently the install dir for ZZ-InstallPaTWPF CANNOT content any blank space. "
            auxS &= "The current install dir has blanks: " & nl & nl
            auxS &= DirEjeucion & nl & nl
            auxS &= "Exit from the program an copy all the program files in a directory without blanks."
            AnyadetxtSalida(auxS)
            auxI = MsgBox(auxS, MsgBoxStyle.Critical, "Install directory has blanks")
            Exit Sub ' salgo
        End If

        ' desempaqueto

        ' Necesito NombrePatsinExtension
        NombrePat = Path.GetFileName(NombrePaTconPath)
        ' quitarle la extensión es .PaT.zip (case insesitive)
        NombrePatsinExtension = NombrePat
        NombrePatsinExtension = Replace(NombrePatsinExtension, ".zip", "", 1, , Microsoft.VisualBasic.CompareMethod.Text)
        NombrePatsinExtension = Replace(NombrePatsinExtension, ".pat", "", 1, , Microsoft.VisualBasic.CompareMethod.Text)
        ' Donde voy a extraer es DirOutPat con nombre pat
        DirOutPatconPat = DirOutPat & "\" & NombrePatsinExtension
        ' Necesito NombrePatsinExtension
        NombrePat = Path.GetFileName(NombrePaTconPath)



        Dim PreZip = String.Format("""{0}\unzip.exe""", DirEjeucion)
        Dim mdto As String = PreZip & String.Format(" -o ""{0}"" -d ""{1}"" ", NombrePaTconPath, DirOutPatconPat)
        Dim rc As Integer = 0
        AnyadetxtSalida(String.Format("Unziping {0} in {1}...", NombrePaTconPath, DirOutPatconPat))
        'MsgBox(mdto, MsgBoxStyle.Critical, "Debug")


        mandatoZip(mdto, rc)
        If rc <> 0 Then
            auxS = String.Format("Unzip command did not work. ('{0}' exited with rc {1}). ", mdto, rc) & nl & nl
            auxS &= "Check for return codes in http://www.info-zip.org/mans/unzip.html#DIAGNOSTICS . Also check dir paths. " & nl
            auxS &= "If you are unable to solve the problem, process manually this PaT (that is, move it from the PaT local directory to another location and unzip in a convenient place. Then import folders and use reference material if exists.)" & nl & nl
            auxS &= "Returning to the main loop." & nl
            AnyadetxtSalida(auxS)
            MsgBox(auxS, MsgBoxStyle.Critical, "OMG unzip error!")

            Exit Sub
        End If
        AnyadetxtSalida("Done" & nl)
        ' ahora tengo que leer el manifiesto
        ' voy a leer los DirOutPat _MFT.XML"
        Dim archivoMFT As String = ""
        Dim archivoMFTconPath As String = ""
        Dim fileEntries As String() = Directory.GetFiles(DirOutPatconPat)
        For Each archivo In fileEntries
            If archivo.EndsWith("_MFT.XML") Then
                archivoMFT = Path.GetFileName(archivo)
                archivoMFTconPath = archivo
            End If
        Next
        ' si no lo tengo error fatal
        If archivoMFT = "" Then
            auxS = "Cannot find the manifest file  *_MTF.XML. Looks the PaT.zip does not include it, but it should. " & nl & nl
            auxS &= "If you are unable to solve the problem, process manually this PaT (that is, move it from the PaT local directory to another location and unzip in a convenient place. Then import folders and use reference material if exists.)" & nl & nl
            auxS &= "Returning to the main loop." & nl
            AnyadetxtSalida(auxS)
            MsgBox(auxS, MsgBoxStyle.Critical, "OMG Cannot find manifest (*_MFT.XML) in PaT file!")

            Exit Sub
        End If
        AnyadetxtSalida(String.Format("MFT file: {0}", archivoMFT) & nl)
        ' ahora el html que visualizaré
        Dim archivoMFTHTMLconPath As String = ""
        For Each archivo In fileEntries
            If archivo.EndsWith("_MFT.HTM") Then
                archivoMFTHTMLconPath = archivo
            End If
        Next
        ' ahora vamos a leer el archivo MFT
        Dim ListaCarpetasTraducir As New ClaseListaCarpetasTraducir

        Try
            Dim objStreamReader As New StreamReader(archivoMFTconPath)
            Dim x As New XmlSerializer(ListaCarpetasTraducir.GetType)
            ListaCarpetasTraducir = x.Deserialize(objStreamReader)
            objStreamReader.Close()
        Catch ex As Exception
            auxS = "Fatal error: Looks like xml manifest is not valid. " & nl
            auxS &= String.Format("excp: {0}", ex.Message) & nl & nl
            auxS &= "If you are unable to solve the problem, process manually this PaT (that is, move it from the PaT local directory to another location and unzip in a convenient place. Then import folders and use reference material if exists.)" & nl & nl
            auxS &= "Returning to the main loop." & nl
            AnyadetxtSalida(auxS)
            MsgBox(auxS, MsgBoxStyle.Critical, "OMG Manifest file (*_MFT.XML) not valid!")

            Exit Sub
        End Try
        ' 

        ' Voy recuperar las carpetas instaladas. 
        If ListaCarpetasTraducir.FlagAutoInstall = False Then ' pregunto
            ' antes de ponerme a trabajar voy a avisar
            auxS = "New PaT File!" & nl & nl
            If par.DestTipoDestino = "FTP" Then
                auxS &= String.Format("We have find in the ftp server {0} ", par.DestFTPHost) & nl
                auxS &= String.Format("The file has downloaded as {0}", NombrePaTconPath) & nl & nl
            Else
                auxS &= String.Format("There is a new PaT file {0}", NombrePaTconPath) & nl & nl
            End If
            auxS &= String.Format("After running succesfully this program all material will be found in {0}", DirOutPatconPat) & nl & nl
            auxS &= "If you wanto to automatically process this PaT file:" & nl
            auxS &= "1) Say 'Yes' in this dialog." & nl
            auxS &= "- Folders will be imported in OTM2." & nl
            auxS &= String.Format("- All material will be also in the {0} directory ", DirOutPatconPat) & nl & nl
            Select Case par.DestTipoDestino
                Case "FTP"
                    auxS &= "If you wanto to manually process this PaT file:" & nl
                    auxS &= String.Format("1) Go to the ftp host and directory: {0}{0}{1}/{2}{0}{0} ", nl, par.DestFTPHost, par.DestFTPNombre)
                    auxS &= String.Format("2) DOWNLOAD and REMOVE the incoming PaT file: {0}{0}{1}{0}{0} and place it in another directory.{0}", nl, NombrePat)
                    auxS &= "3) Unzip all files manually at your convenience." & nl
                    auxS &= "4) Say 'No' and dismiss this dialog." & nl & nl
                    auxS &= "Note: If you NOT remove the PaT file from the ftp directoy ZZ-Pat will try to download and process it again." & nl & nl

                Case Else
                    auxS &= "If you wanto to manually process this PaT file:" & nl
                    auxS &= String.Format("1) Remove the incoming PaT file:{0}{1}{0} and place it in another directory.", NombrePaTconPath, nl)
                    auxS &= "2) Unzip all files manually at your convenience." & nl
                    auxS &= "3) Say 'No' and dismiss this dialog." & nl
                    auxS &= "Note: If you NOT remove the PaT file from incoming directoy ZZ-Pat will try to process it again." & nl & nl
            End Select

            auxS &= "Do you want the PaT to be automatically unpackaged in the local otput dir and the folders to be imported in OTM2?" & nl
            AnyadetxtSalida(auxS)
            'auxI = MsgBox(auxS, MsgBoxStyle.YesNo, "Process PaT package?")
            Application.Current.MainWindow.WindowState = Windows.WindowState.Normal
            auxI = MessageBox.Show(auxS, "Do we unpackage the PaT file? ", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.Yes)
            If auxI = MsgBoxResult.No Then
                Select Case par.DestTipoDestino
                    Case "FTP"
                        auxS = "Again, new remainder, if you have not, pls remove the PaT file: {0}{0}{1}{0}{0} from {0}{0}{2}/{3}{0}{0}  otherwise you continually receive this message.{0} "
                        auxS = String.Format(auxS, nl, NombrePat, par.DestFTPHost, par.DestFTPNombre)
                    Case Else
                        auxS = "Again, new remainder, if you have not, pls remove the PaT file {0} from {1}, otherwise you continually receive this message. " & nl
                        auxS = String.Format(auxS, NombrePat, DirOutPat)
                End Select
                AnyadetxtSalida(auxS)
                MsgBox(auxS, MsgBoxStyle.Information, "Please, remove PaT file from local PaT directory")

                Exit Sub
            End If




        End If


        ' he leido el archivo pat
        Dim row As DataRow
        AnyadetxtSalida("Folders in PaT" & nl)
        For Each row In ListaCarpetasTraducir.tCarpetasTraducir.Rows
            AnyadetxtSalida(row("carpeta") & nl)
        Next
        ' primero voy a ver que no estén en OTM2

        Dim mandato As String = ""
        Dim carpeta As String = ""
        Dim HayProblema As Boolean = True
        Dim listaCarpetasABorrar As New List(Of String)
        Do
            InfoOTM = New ClassInfoOTM ' vuelvo a recuperar la liwsta de carpetas por si han borrado
            If InfoOTM.TodoOK = False Then
                auxS = String.Format("Error in OTM environment {0}", InfoOTM.Info) & nl
                AnyadetxtSalida(auxS)
                auxS &= "You should exit from the program untill this problem is solved. Otherwise the loop will find de PaT file again." & nl & nl
                auxS &= "Returning to the main loop." & nl
                AnyadetxtSalida(auxS)
                MsgBox(auxS, MsgBoxStyle.Critical, "Error in OTM!")
                Exit Sub
            End If
            auxS = ""
            HayProblema = False  ' optimismo
            For Each row In ListaCarpetasTraducir.tCarpetasTraducir.Rows
                carpeta = row("carpeta")
                For Each s In InfoOTM.Carpetas
                    If s = carpeta Then
                        HayProblema = True
                        auxS &= carpeta & nl
                        listaCarpetasABorrar.Add(carpeta)
                        Exit For
                    End If
                Next
            Next
            If HayProblema Then '
                'podemos borrar las carpetas o bloquear la importación
                'para que decida el usuario segun FlagDeleteFIfExists
                If ListaCarpetasTraducir.FlagDeleteFIfExists Then
                    AnyadetxtSalida("We have to delete the following folders in the target system" & nl)
                    AnyadetxtSalida("because they are in the PaT in order to translate them:" & nl)
                    For Each carpetaABorrar As String In listaCarpetasABorrar
                        AnyadetxtSalida(carpetaABorrar & nl)
                    Next

                    For Each carpetaABorrar As String In listaCarpetasABorrar
                        AnyadetxtSalida(String.Format("Deleting {0}...", carpetaABorrar))

                        mandato = " /TAsk=FLDDEL /FLD={0} /QUIET=NOMSG  "
                        mandato = String.Format(mandato, carpetaABorrar)
                        ejecuta(mandato, rc)
                        If rc <> 0 Then
                            auxS = nl & String.Format("OPS {0} cannot be deleted. RC {1} ", carpeta, rc) & nl & nl
                            auxS &= "Process manually this PaT (that is, move it from the PaT local directory to another location and unzip in a convenient place. Then import folders and use reference material if exists.)" & nl & nl
                            auxS &= "Returning to the main loop." & nl
                            AnyadetxtSalida(auxS)
                            MsgBox(auxS, MsgBoxStyle.Critical, "Error deleting FXP folder!")
                            Exit Sub
                        Else
                            AnyadetxtSalida("done!" & nl) ' empieza con "Importing xxxx...
                        End If

                    Next
                Else ' manual
                    auxS = "The following folders in the PaT file ALREADY exists in" &
                       "your OpenTM environment: " & nl & nl & auxS & nl & nl &
                       "You should:" & nl &
                       "1) Leave this dialog open. " & nl &
                       "2) Open OpenTM and delete the folders  " & nl &
                       "3) Select 'Retry' " & nl & nl &
                       "If you select 'Cancel' you will return to the main loop. In this " & nl &
                       "case you should process the PaT file manually, othewise the " & nl &
                       "this verification will be run again."
                    AnyadetxtSalida(auxS)
                    auxI = MsgBox(auxS, MsgBoxStyle.RetryCancel, "Folder in PaT already in OpenTM2")
                    If auxI = MsgBoxResult.Cancel Then
                        Exit Sub
                    End If
                End If
            End If
        Loop While HayProblema


        AnyadetxtSalida(nl)
        For Each row In ListaCarpetasTraducir.tCarpetasTraducir.Rows
            AnyadetxtSalida(String.Format("Importing {0}...", row("carpeta")))
            carpeta = row("carpeta") ' p.e c:\kkk\kkk.fxp
            Dim fromDrive As String = Path.GetPathRoot(DirOutPatconPat)
            Dim fromPath As String = DirOutPatconPat : fromPath = Replace(fromPath, fromDrive, "")
            If fromPath.Substring(0, 1) <> "\" Then fromPath = "\" & fromPath
            fromDrive = fromDrive.Substring(0, 1) ' la letra de c:\
            Dim FLD As String = carpeta

            mandato = " /TAsk=FLDIMP /FLD={0} /FROMdrive={1} /FromPath={2} /OPtions=(MEM,DICT)  /QUIET=NOMSG  "
            mandato = String.Format(mandato, FLD, fromDrive, fromPath, DiscoOTM, DirOutPatconPat) ' 
            ejecuta(mandato, rc)
            If rc <> 0 Then
                auxS = nl & String.Format("OPS {0} cannot be imported. RC {1} ", carpeta, rc) & nl & nl
                auxS &= "Process manually this PaT (that is, move it from the PaT local directory to another location and unzip in a convenient place. Then import folders and use reference material if exists.)" & nl & nl
                auxS &= "Returning to the main loop." & nl
                AnyadetxtSalida(auxS)
                MsgBox(auxS, MsgBoxStyle.Critical, "Error importing FXP folder!")
            Else
                AnyadetxtSalida("done!" & nl) ' empieza con "Importing xxxx...
            End If
        Next
        ' SACO LA INFO
        Process.Start(archivoMFTHTMLconPath)

        ' tengo que sar el PaT al _WIP
        ' NombrePaTconPath es el source
        Dim DirWIP As String = DirOutPat & "\" & "_WIP"
        If Directory.Exists(DirWIP) = False Then
            AnyadetxtSalida("Working In Progress (_WIP) directoy does not exist. Creating " & DirWIP & "...")
            Try
                Directory.CreateDirectory(DirWIP)
                AnyadetxtSalida("Done!" & nl)
            Catch ex As Exception
                auxS = String.Format("Cannot create WIP dir ({0}). ", DirWIP) & nl & nl
                auxS &= String.Format("Excep. {0}", ex.Message.ToString) & nl
                auxS &= "1) PaT is CORRECTLY instaled." & nl
                auxS &= "2) Create the WIP dir manually " & nl
                auxS &= String.Format("3) Move manually PaT from {0} to {1}", NombrePaTconPath, DirWIP) & nl
                auxS &= "3) Dismiss this dialog. If you miss 2) main loop will find the PaT again, and will reinstalled (another valid option) " & nl & nl
                auxS &= "4) After dismissing the dialog, you will return to the main loop."
                AnyadetxtSalida(auxS)
                MsgBox(auxS, MsgBoxStyle.Critical, "Cannot move PaT to _WIP directory!" & nl)

                Exit Sub
            End Try
        End If
        Dim NombrePaTenWIP As String = DirWIP & "\" & NombrePat
        If TipoDeposito = "PATH" Or TipoDeposito = "SHARED_DRIVE" Or TipoDeposito = "FTP" Then
            Try
                If File.Exists(NombrePaTenWIP) Then
                    File.Delete(NombrePaTenWIP) ' lo borro si existe el move, no tiene overwrite
                End If
                File.Move(NombrePaTconPath, NombrePaTenWIP)
            Catch ex As Exception
                auxS = String.Format("Cannot move PaT from {0} to {1}. ", NombrePaTconPath, DirWIP) & nl & nl
                auxS &= String.Format("Excep. {0}", ex.Message.ToString) & nl
                auxS &= "1) PaT is CORRECTLY instaled." & nl
                auxS &= String.Format("2) Move manually PaT from {0} to {1}", NombrePaTconPath, DirWIP) & nl
                auxS &= "3) Dismiss this dialog. If you miss 2) main loop will find the PaT again, and will try it again. " & nl
                auxS &= "4) After dismissing the dialog, you will return to the main loop."
                AnyadetxtSalida(auxS)
                MsgBox(auxS, MsgBoxStyle.Critical, "Cannot move the PaT file to  the _WIP directory!" & nl)

                Exit Sub
            End Try
        End If
        cargalbxWIP(par.DestLocalPatDir)
        ' limpieza en el caso de FTP
        Select Case TipoDeposito
            Case "FTP"
                ' Archivo FTP, voy a conectarme al servidor FTP para leer el directorio
                NombrePaTconPath = "" ' supongamos que no hay
                Dim ftpURI As String
                ftpURI = "ftp://" & par.DestFTPHost & "/" & par.DestFTPNombre & "/" ' si no añado / al final me sale el subdir
                Dim contrasenya As String = DecryptWithKey(par.DestFTPContrasenya, par.DestLongInterno)
                Dim ftpCrendenciales As NetworkCredential = New System.Net.NetworkCredential(par.DestFTPUsuario, contrasenya)
                Dim FTPSolicitud As FtpWebRequest
                Dim FTPRespuesta As FtpWebResponse
                Dim auxS As String = ""

                Try
                    ftpURI &= NombrePat
                    FTPSolicitud = DirectCast(WebRequest.Create(New Uri(ftpURI)), FtpWebRequest)
                    FTPSolicitud.Credentials = ftpCrendenciales
                    FTPSolicitud.Method = WebRequestMethods.Ftp.DeleteFile
                    Dim respuesta = DirectCast(FTPSolicitud.GetResponse(), FtpWebResponse)


                Catch ex As System.Net.WebException
                    AnyadetxtSalida("FTP ERROR *.PaT.zip " & nl)
                    AnyadetxtSalida(ex.Message)
                    If Not (ex.InnerException Is Nothing) Then
                        AnyadetxtSalida(ex.InnerException.Message)
                    End If
                    Exit Sub
                End Try
            Case "PATH" ' no hago nada
            Case "SHARED_DRIVE" ' no hago nada
            Case Else
                auxS = String.Format("Target type is '{0}'. It should be PATH, SHARED_DRIVE or FTP.", TipoDeposito)
                AnyadetxtSalida(auxS & nl)
                auxS &= nl & "Do you want to exit the program?"

                auxI = MsgBox(auxS, MsgBoxStyle.YesNo, "Do you want to exit the program?")
                AnyadetxtSalida(auxI & nl)
                If auxI = MsgBoxResult.Yes Then
                    Application.Current.Shutdown()
                End If
        End Select
        auxS = "PaT file succesfully instaled!"
        AnyadetxtSalida(auxS & nl)
        If ListaCarpetasTraducir.FlagAutoInstall = False Then
            MsgBox(auxS, MsgBoxStyle.Information, "PaT succesfully instaled!")
        End If

        AnyadetxtSalida("Process finished successfully. Returning to the main loop " & nl)
        AnyadetxtSalida("**********************************************************" & nl & nl)

    End Sub

    Sub cargalbxWIP(LocalPatDir As String)
        ' carga las carpetas en WIP:
        ' tengo que ver si existe _WIP
        Dim DirWIP As String = LocalPatDir & "\_WIP"
        lbxWIP.Items.Clear()
        Try
            If Not Directory.Exists(DirWIP) Then Exit Sub ' no existe salgo
            '        miro los archivos Pat
            Dim archivos() As String = Directory.GetFiles(DirWIP, "*.PaT.zip")
            For Each s In archivos
                s = Path.GetFileName(s)
                ' quitar .pat. y .zip case insensitive
                s = Replace(s, ".zip", "", 1, , Microsoft.VisualBasic.CompareMethod.Text)
                s = Replace(s, ".pat", "", 1, , Microsoft.VisualBasic.CompareMethod.Text)
                lbxWIP.Items.Add(s)
            Next

        Catch ex As Exception
            ' simplemente no hago nada

        End Try

        '
    End Sub

    Sub cargalbxDONE(LocalPatDir As String)
        ' carga las carpetas en DONE:
        ' tengo que ver si existe _DONE
        Dim DirDone As String = LocalPatDir & "\_DONE"
        lbxDone.Items.Clear()
        Try
            If Not Directory.Exists(DirDone) Then Exit Sub ' no existe salgo
            '        miro los archivos Pat
            Dim archivos() As String = Directory.GetFiles(DirDone, "*.PaT.zip")
            For Each s In archivos
                s = Path.GetFileName(s)
                ' quitar .pat. y .zip case insensitive
                s = Replace(s, ".zip", "", 1, , Microsoft.VisualBasic.CompareMethod.Text)
                s = Replace(s, ".pat", "", 1, , Microsoft.VisualBasic.CompareMethod.Text)
                lbxDone.Items.Add(s)
            Next

        Catch ex As Exception
            ' simplemente no hago nada

        End Try

    End Sub
    Function AnyadetxtSalida(s As String)
        txtSalida.AppendText(s)
        txtSalida.SelectionStart = txtSalida.Text.Length
        txtSalida.ScrollToEnd()
        txtSalida.Refresh()
        Return 0
    End Function

    Private Sub btnSalida_Click(sender As Object, e As RoutedEventArgs) Handles btnSalida.Click
        Application.Current.Shutdown()
    End Sub

    Private Sub btnParametros_Click(sender As Object, e As RoutedEventArgs) Handles btnParametros.Click
        Dim vp As New VentanaParametros
        dt.Stop() : AnyadetxtSalida("Timer stoped.")
        vp.ShowDialog()
        ' debo cargar los parámetros
        Dim Par As structParametros
        Par = cargaParametros()
        ' dt.Start() : AnyadetxtSalida("Timer restarted.")

    End Sub

    Private Sub btnInfoWip_Click(sender As Object, e As RoutedEventArgs) Handles btnInfoWip.Click
        If lbxWIP.SelectedIndex >= 0 Then ' algo hay seleccionado
            Dim NombrePat As String = lbxWIP.SelectedItem.ToString
            ' necesito los parámetros
            Dim par As structParametros
            par = CargaParametros()
            Dim LocalPatDir As String = par.DestLocalPatDir
            Dim InfoPaTHTML As String = LocalPatDir & "\" & NombrePat & "\" & NombrePat & "_MFT.HTM"
            If File.Exists(InfoPaTHTML) Then
                Process.Start(InfoPaTHTML)
            Else
                auxS = String.Format("Cannot find {0}. The file or the unpackage PaT has been deleted. If dont need this package anymore, delete it. ", InfoPaTHTML)
                MsgBox(auxS, MsgBoxStyle.Critical, "Cannot find the info html")
            End If
            ' SACO LA INFO
            Dim pro As New Process()

            pro.StartInfo.UseShellExecute = True
            '' String.Format("@""{0}""", par.DestLocalPatDir)
            pro.StartInfo.FileName = LocalPatDir & "\" & NombrePat & "\"
            'Dim auxS As String = 
            pro.Start()
        Else
            MsgBox("Select a PaT name", MsgBoxStyle.Information, "No PaT selected")

        End If

    End Sub

    Private Sub btnDevolverPaT_Click(sender As Object, e As RoutedEventArgs) Handles btnDevolverPaT.Click
        AnyadetxtSalida("************************************************************" & nl)
        AnyadetxtSalida("Starting return proces..." & nl)
        dt.Stop() : AnyadetxtSalida("Timer stoped" & nl)
        Dim NombrePaTconPath As String
        Dim NombrePatsinExtension As String
        'Dim NombrePat As String
        Dim DirOutPat As String
        Dim DirOutPatconPat As String ' 
        Dim DirTMP As String
        Dim DirEjeucion As String
        ' 
        NombrePaTconPath = ""
        NombrePatsinExtension = ""
        'NombrePat = ""
        DirOutPat = ""
        DirOutPatconPat = ""
        DirTMP = ""
        DirEjeucion = ""
        '
        Dim auxi As Integer = 0 ' auxiliar

        ' hay que devolver el pat
        ' primero ver si me han seleccionado algo
        If lbxWIP.SelectedIndex < 0 Then
            MsgBox("In order to return a PaT, first you must select one in the listbox. Select one and try again.", MsgBoxStyle.Exclamation, "No PaT to return")
            ' dt.Start() : AnyadetxtSalida("No PaT selected. Abort. Timer restarted" & nl)
            Exit Sub
        Else ' tenemos uno
            NombrePatsinExtension = lbxWIP.SelectedItem.ToString
        End If
        '

        ' ver si encuentro el manifiesto necesito saber el LocalPatDir
        Dim par As structParametros
        par = CargaParametros()
        Dim LocalPatDir As String = par.DestLocalPatDir

        ' ahora tengo que leer el manifiesto
        ' voy a leer los DirOutPat _MFT.XML"
        Dim archivoMFT As String = NombrePatsinExtension & "_MFT.XML"
        Dim archivoMFTconPath As String = LocalPatDir & "\" & NombrePatsinExtension & "\" & archivoMFT
        ' si no lo tengo error fatal
        If Not File.Exists(archivoMFTconPath) Then
            auxS = String.Format("Cannot find {0}. Unpackaged PaT is not found in ({1}) or the *_MFT.XML has been deleted. ", archivoMFTconPath, LocalPatDir & "\" & NombrePatsinExtension) & nl
            auxS &= String.Format("Restore {0} in the {1} directory or return the folders manually", NombrePatsinExtension & ".PaT.zip", LocalPatDir & "\" & NombrePatsinExtension)
            MsgBox(auxS, MsgBoxStyle.Critical, "Cannot find manifest file")
            ' dt.Start() : AnyadetxtSalida("Cannot find MFT.XML. Abort. Timer restarted" & nl)
            Exit Sub
        End If
        AnyadetxtSalida(String.Format("MFT file found: {0}", archivoMFT) & nl)
        ' ahora debo leerlo 
        ' ahora vamos a leer el archivo MFT
        Dim ListaCarpetasTraducir As New ClaseListaCarpetasTraducir
        Try
            Dim objStreamReader As New StreamReader(archivoMFTconPath)
            Dim x As New XmlSerializer(ListaCarpetasTraducir.GetType)
            ListaCarpetasTraducir = x.Deserialize(objStreamReader)
            objStreamReader.Close()
        Catch ex As Exception
            auxS = "Fatal error: Looks like xml manifest is not valid. " & nl
            auxS &= String.Format("excp: {0}", ex.Message) & nl
            auxS &= "Return the folders manually." & nl & nl

            AnyadetxtSalida(auxS)
            MsgBox(auxS, MsgBoxStyle.Critical, "OMG Manifest file (*_MFT.XML) not valid!")
            ' dt.Start() : AnyadetxtSalida("Manifest file (MFT.XML) not valid. Abort. Timer restarted" & nl)
            Exit Sub
        End Try
        ' he leido el archivo pat
        Dim row As DataRow
        AnyadetxtSalida("Folders in PaT" & nl)
        auxS = "" ' para la lista de carpetas
        For Each row In ListaCarpetasTraducir.tCarpetasTraducir.Rows
            AnyadetxtSalida(row("carpeta") & nl)
            auxS &= row("carpeta") & nl
        Next
        ' Necesito tambien el correo del gestor
        Dim CorreoGestor As String = ListaCarpetasTraducir.CorreoGestor
        '
        auxS = "You are about to: " & nl & nl & "1) Export the following folders:" & nl & nl & auxS & nl
        auxS &= String.Format("2) Zip de folders and send them to :" & nl & nl)
        Dim DirOutFXZ As String = ""
        Dim DirOutFXZTipoDestino = par.DestTipoDestino
        ' por si hay FTP
        Dim Contrasenya As String = DecryptWithKey(par.DestFTPContrasenya, par.DestLongInterno)
        Dim hostFTP As String = "ftp://" & par.DestFTPHost
        'If Not hostFTP.EndsWith("/") Then hostFTP &= "/"
        Dim PrefijoTarget As String = hostFTP & "/" & par.DestFTPNombre
        ' If Not target.EndsWith("/") Then target &= "/"

        Select Case DirOutFXZTipoDestino
            'Case "PATH"
            '    DirOutFXZ = par.DestPath
            '    auxS &= String.Format("'{0}' ({1})", par.DestPath, par.DestTipoDestino) & nl & nl
            Case "SHARED_DRIVE"
                DirOutFXZ = par.DestSharedDriveNombre
                auxS &= String.Format("'{0}' ({1})", par.DestSharedDriveNombre, par.DestTipoDestino) & nl & nl
            Case "FTP"
                auxS &= String.Format("'{0}' ({1})", par.DestFTPHost, par.DestFTPNombre) & nl & nl
        End Select
        auxS &= "3) The PaT file will be moved from the working in progress directory to the done directory" & nl & nl
        auxS &= "Be sure all folders in the list are fully translated and ready tu return!" & nl & nl
        auxS &= "Doy you want to continue?"
        auxi = MsgBox(auxS, MsgBoxStyle.YesNo, "Folders ready to be sent")
        If auxi = MsgBoxResult.No Then
            ' dt.Start() : AnyadetxtSalida("User has aborted sending. Abort. Returning to main loop. Restaring timer." & nl)
            Exit Sub
        End If
        ' cierrame OTM
        While OTMEnEjecucion()  ' error fatal no haré nada si está en marcha
            auxS = "OTM2 is running. You need to close it. Is it closed?" & nl
            auxS &= "Pls, close OTM2 and once is closed answer Retry. If you answer Cancel the search loop will start again. "
            AnyadetxtSalida(auxS)
            auxi = MsgBox(auxS, MsgBoxStyle.RetryCancel, "Opss OTM2 is running, close it!")
            If auxi = MsgBoxResult.Cancel Then
                ' dt.Start() : AnyadetxtSalida("User has aborted sending. Abort. Returning to main loop. Restaring timer." & nl)
                Exit Sub
            End If
        End While

        ' directorio temporal
        DirTMP = Path.GetTempPath() & "ZZ"
        AnyadetxtSalida(String.Format("Temporary dir {0}", DirTMP) & nl)
        Try
            If Not Directory.Exists(DirTMP) Then Directory.CreateDirectory(DirTMP)
            For Each s In Directory.GetFiles(DirTMP)
                File.Delete(s)
            Next
        Catch ex As Exception
        End Try
        ' 
        Dim mandato As String = ""
        ' directorio ejecución
        DirEjeucion = System.Reflection.Assembly.GetExecutingAssembly().Location ' ubicación ejecutable
        DirEjeucion = Path.GetDirectoryName(DirEjeucion)
        ' Primero exporto a al dir temporal

        auxS = Path.GetPathRoot(DirTMP) ' C:\
        auxS = auxS.Replace("\", "") ' C: que es lo que debo quitar
        Dim DirTemporalSinUnidad As String = DirTMP.Replace(auxS, "")

        Dim unidad As String = Path.GetPathRoot(DirTMP).Substring(0, 1) ' la letra
        '
        Dim carpeta As String = ""
        Dim rc As Integer = 0
        For Each row In ListaCarpetasTraducir.tCarpetasTraducir.Rows
            AnyadetxtSalida(String.Format("Exporting folder {0} in {1}", row("carpeta"), DirTMP) & nl)
            carpeta = row("carpeta")
            mandato = " /TAsk=FLDEXP /FLD={0} /TOdrive={1} /ToPath={2} /OPtions=(MEM,ROMEM,DOCMEM) /OVerwrite=YES /QUIET=NOMSG  "
            mandato = String.Format(mandato, carpeta, unidad, DirTemporalSinUnidad) ' exportaré al directorio temporal
            ejecuta(mandato, auxi)
            If rc <> 0 Then
                auxS = "Folder " & carpeta & " cannot be exported in directory " & DirTMP & ". " &
                       "RC=" & rc.ToString
                auxS &= "Try to fix the problem and retry or send the folders manually" & nl
                MsgBox(auxS, MsgBoxStyle.Exclamation, "Folder cannot be exported" & nl)
                txtSalida.AppendText(auxS & nl)
                ' dt.Start() : AnyadetxtSalida("Cannot export folder. Abort. Returning to main loop. Restaring timer." & nl)
                Exit Sub
            End If
            AnyadetxtSalida("Done!" & nl)

        Next


        'txtSalida.AppendText("Carpeta " & carpeta & " exporada en " & Par.ParDirTemporal & nl)
        ' ahora el zip
        Dim PreZip As String = System.Reflection.Assembly.GetExecutingAssembly().Location
        PreZip = Path.GetDirectoryName(PreZip) & "\zip.exe "
        ' zipearemos en el directorio temporal
        For Each row In ListaCarpetasTraducir.tCarpetasTraducir.Rows
            carpeta = row("carpeta")
            auxS = String.Format("Zipping the folder {0} in {1}... ", carpeta, DirTMP) & nl
            AnyadetxtSalida(auxS)
            'mandato = PreZip & " -j " & DirOutFXZ & "\" & carpeta & ".FXZ " & DirTMP & carpeta & ".FXP"
            mandato = String.Format("{0} -j ""{1}\{2}.FXZ""  ""{3}\{4}.FXP"" ", PreZip, DirTMP, carpeta, DirTMP, carpeta)
            mandatoZip(mandato, rc)
            If rc <> 0 Then
                auxS = String.Format("Zip command did not work. ('{0}' exited with rc {1}). ", mandato, rc) & nl & nl
                auxS &= "Check for return codes in http://www.info-zip.org/mans/unzip.html#DIAGNOSTICS . " & nl
                auxS &= "If you are unable to solve the problem, send de folders manually." & nl & nl
                auxS &= "Returning to the main loop." & nl
                AnyadetxtSalida(auxS)
                MsgBox(auxS, MsgBoxStyle.Critical, "OMG zip error!")
                ' dt.Start() : AnyadetxtSalida("Cannot zip folder. Abort. Returning to main loop. Restaring timer." & nl)
                Exit Sub
            End If
            ' lo borro
            AnyadetxtSalida("Done!" & nl)
            File.Delete(String.Format("{0}\{1}.FXP", DirTMP, carpeta))
        Next
        ' ya lo tengo en el directorio temporal
        Dim auxFileSource As String = "" : Dim auxFileTarget As String = ""
        For Each row In ListaCarpetasTraducir.tCarpetasTraducir.Rows
            carpeta = row("carpeta")
            auxFileSource = String.Format(String.Format("{0}\{1}.fxz", DirTMP, carpeta))
            auxFileTarget = String.Format(String.Format("{0}\{1}.fxz", DirOutFXZ, carpeta))
            If par.DestTipoDestino = "PATH" Or par.DestTipoDestino = "SHARED_DRIVE" Then
                ' lo coloco directamente en el destino
                ' source es DirTMP, target es DirOutFXZ
                auxS = String.Format("Copying folder {0}.fxz in {1}... ", carpeta, DirOutFXZ) & nl
                AnyadetxtSalida(auxS)
                'mandato = PreZip & " -j " & DirOutFXZ & "\" & carpeta & ".FXZ " & DirTMP & carpeta & ".FXP"
                ' primero voy a borrar si puedo el target
                Try
                    File.Delete(auxFileTarget)
                Catch ex As Exception

                End Try
                ' ahora lo copio
                File.Copy(auxFileSource, auxFileTarget)
                AnyadetxtSalida("Done!" & nl)

            ElseIf par.DestTipoDestino = "FTP" Then '  debo conectarme y copiar desde el temporal al destino
                Dim source As String = auxFileSource
                ' hostFTP y target calculados al principio
                Dim target As String = PrefijoTarget & "/" & carpeta & ".FXZ" ' destino FTP
                ' https://social.msdn.microsoft.com/Forums/en-US/246ffc07-1cab-44b5-b529-f1135866ebca/exception-the-underlying-connection-was-closed-the-connection-was-closed-unexpectedly?forum=netfxnetcom
                System.Net.ServicePointManager.Expect100Continue = False

                Dim ftprequest As FtpWebRequest = DirectCast(System.Net.WebRequest.Create(target), FtpWebRequest)
                ftprequest.Credentials = New System.Net.NetworkCredential(par.DestFTPUsuario, Contrasenya)
                ftprequest.Method = WebRequestMethods.Ftp.UploadFile
                ftprequest.EnableSsl = False
                ftprequest.UsePassive = True
                ftprequest.UseBinary = True
                ftprequest.KeepAlive = False
                ftprequest.ReadWriteTimeout = 3600000
                ftprequest.Timeout = 3600000
                ftprequest.ContentLength = source.Length
                Dim file As System.IO.FileInfo = New System.IO.FileInfo(source)
                Dim estaLongiArchivo As Long = file.Length
                Dim esta5Porciento As Long = estaLongiArchivo * 0.05
                Dim estaSalidaInfo As Long = esta5Porciento
                Dim estaProgreso As Long = 0
                Dim estaN5Porciento As Long = 0
                Dim estaTamBuffer As Integer = 8192 ' en ejemmplos normalemnte ponen menos, en teoria MTU 8K
                Dim buffer As Integer = estaTamBuffer
                Dim content(buffer - 1) As Byte, dataread As Integer
                Try
                    Using ftpstream As System.IO.FileStream = file.OpenRead()
                        Using request As System.IO.Stream = ftprequest.GetRequestStream
                            Do
                                dataread = ftpstream.Read(content, 0, buffer)
                                request.Write(content, 0, dataread)
                                estaProgreso += estaTamBuffer ' lo transferido
                                If estaProgreso > estaSalidaInfo Then
                                    While estaProgreso > estaSalidaInfo
                                        estaN5Porciento += 1
                                        estaSalidaInfo += esta5Porciento
                                    End While
                                    AnyadetxtSalida(String.Format("Transfered... {0} % ({1}) bytes ", estaN5Porciento * 5, estaProgreso) & nl)
                                End If
                            Loop Until dataread < buffer

                            request.Dispose()
                        End Using
                        ftpstream.Dispose()
                    End Using
                    Dim response As FtpWebResponse = CType(ftprequest.GetResponse(), FtpWebResponse)
                    AnyadetxtSalida(String.Format("Upload File Complete, status {0}", response.StatusDescription))


                Catch ex As System.Net.WebException
                    MsgBox(auxS, MsgBoxStyle.Critical, "OMG FTP error!")
                    ' dt.Start() : AnyadetxtSalida("Looks like ftp server has some problemes. Try later" & nl)
                    Exit Sub
                End Try

            Else
                auxS = "Fatal internal error. Cannot copy from TEMP dir to repository space " & nl
                auxS &= "because par.DestTipoDestino is not PATH, SHARED_DRIVE nor FTP. " & nl
                AnyadetxtSalida(auxS)
                MsgBox(auxS, MsgBoxStyle.Critical, "OMG internal error!")
                ' dt.Start() : AnyadetxtSalida("Internal error. par.DestTipoDestino not FTP, SHARED_DRIVE nor PATH." & nl)
                Exit Sub
            End If
        Next



        ' ahora copiaré el manifesto 
        ' archivoMFTconPath es el origen 
        auxFileSource = archivoMFTconPath
        auxFileTarget = ""
        auxFileTarget = String.Format(String.Format("{0}\{1}", DirOutFXZ, archivoMFT))
        If par.DestTipoDestino = "PATH" Or par.DestTipoDestino = "SHARED_DRIVE" Then
            ' lo coloco directamente en el destino
            ' source es DirTMP, target es DirOutFXZ
            auxS = String.Format("Copying folder {0} in {1}... ", auxFileSource, auxFileTarget) & nl
            AnyadetxtSalida(auxS)
            ' ahora lo copio con overwrite
            File.Copy(auxFileSource, auxFileTarget, True)
            AnyadetxtSalida("Done!" & nl)

        ElseIf par.DestTipoDestino = "FTP" Then '  debo conectarme y copiar desde el temporal al destino
            Dim source As String = auxFileSource
            ' hostFTP y target calculados al principio
            Dim target As String = PrefijoTarget & "/" & archivoMFT ' destino FTP
            ' https://social.msdn.microsoft.com/Forums/en-US/246ffc07-1cab-44b5-b529-f1135866ebca/exception-the-underlying-connection-was-closed-the-connection-was-closed-unexpectedly?forum=netfxnetcom
            System.Net.ServicePointManager.Expect100Continue = False

            Dim ftprequest As FtpWebRequest = DirectCast(System.Net.WebRequest.Create(target), FtpWebRequest)
            ftprequest.Credentials = New System.Net.NetworkCredential(par.DestFTPUsuario, Contrasenya)
            ftprequest.Method = WebRequestMethods.Ftp.UploadFile
            ftprequest.EnableSsl = False
            ftprequest.UsePassive = True
            ftprequest.UseBinary = True
            ftprequest.KeepAlive = False
            ftprequest.ReadWriteTimeout = 3600000
            ftprequest.Timeout = 3600000
            ftprequest.ContentLength = source.Length
            Dim file As System.IO.FileInfo = New System.IO.FileInfo(source)
            Dim estaLongiArchivo As Long = file.Length
            Dim esta5Porciento As Long = estaLongiArchivo * 0.05
            Dim estaSalidaInfo As Long = esta5Porciento
            Dim estaProgreso As Long = 0
            Dim estaN5Porciento As Long = 0
            Dim estaTamBuffer As Integer = 8192 ' en ejemmplos normalemnte ponen menos, en teoria MTU 8K
            Dim buffer As Integer = estaTamBuffer
            Dim content(buffer - 1) As Byte, dataread As Integer
            Try
                Using ftpstream As System.IO.FileStream = file.OpenRead()
                    Using request As System.IO.Stream = ftprequest.GetRequestStream
                        Do
                            dataread = ftpstream.Read(content, 0, buffer)
                            request.Write(content, 0, dataread)
                            estaProgreso += estaTamBuffer ' lo transferido
                            If estaProgreso > estaSalidaInfo Then
                                While estaProgreso > estaSalidaInfo
                                    estaN5Porciento += 1
                                    estaSalidaInfo += esta5Porciento
                                End While
                                AnyadetxtSalida(String.Format("Transfered... {0} % ({1}) bytes ", estaN5Porciento * 5, estaProgreso) & nl)
                            End If
                        Loop Until dataread < buffer

                        request.Dispose()
                    End Using
                    ftpstream.Dispose()
                End Using
                Dim response As FtpWebResponse = CType(ftprequest.GetResponse(), FtpWebResponse)
                AnyadetxtSalida(String.Format("Upload File Complete, status {0}", response.StatusDescription))


            Catch ex As System.Net.WebException
                MsgBox(auxS, MsgBoxStyle.Critical, "OMG FTP error!")
                ' dt.Start() : AnyadetxtSalida("Looks like ftp server has some problemes. Try later" & nl)
                Exit Sub
            End Try

        Else
            auxS = "Fatal internal error. Cannot copy from Manifiest dir to repository space " & nl
            auxS &= "because par.DestTipoDestino is not PATH, SHARED_DRIVE nor FTP. " & nl
            AnyadetxtSalida(auxS)
            MsgBox(auxS, MsgBoxStyle.Critical, "OMG internal error!")
            ' dt.Start() : AnyadetxtSalida("Internal error. par.DestTipoDestino not FTP, SHARED_DRIVE nor PATH." & nl)
            Exit Sub
        End If




        ' ahora borro los archivos temporales
        Try
            For Each row In ListaCarpetasTraducir.tCarpetasTraducir.Rows
                carpeta = row("carpeta")
                File.Delete(String.Format("{0}/{1}.FXZ", DirTMP, carpeta))
            Next
            '
        Catch ex As Exception

        End Try
        ' ahora muevo el pat de _WIP a _DONE
        Dim archivoPaTenWIPconPath As String = LocalPatDir & "\_WIP\" & NombrePatsinExtension & ".PaT.zip"
        Dim archivoPaTenDONEconPath As String = LocalPatDir & "\_DONE\" & NombrePatsinExtension & ".PaT.zip"
        Dim DirDONE As String = LocalPatDir & "\_DONE"
        AnyadetxtSalida("Moving PaT form WIP to Done directory..." & nl)
        Try
            If Directory.Exists(DirDONE) = False Then
                AnyadetxtSalida("Working In Progress (_DONE) directoy does not exist. Creating " & DirDONE & "...")

                auxS = String.Format("Cannot create DONE dir ({0}). ", DirDONE) & nl & nl ' por si hay error
                Directory.CreateDirectory(DirDONE)
                AnyadetxtSalida("Done!")
            End If
            auxS = String.Format("Cannot delete file {0}", archivoPaTenDONEconPath) & nl & nl ' por si hay error
            If File.Exists(archivoPaTenDONEconPath) Then File.Delete(archivoPaTenDONEconPath)
            auxS = String.Format("Cannot move from {0} to {1}.", archivoPaTenWIPconPath, DirDONE) & nl & nl ' por si hay un error
            File.Move(archivoPaTenWIPconPath, archivoPaTenDONEconPath)
        Catch ex As Exception
            auxS &= String.Format("Excep. {0}", ex.Message.ToString) & nl
            auxS &= "1) Folders have been correctly sent. You do not need to send them again." & nl
            auxS &= String.Format("2) Move manually PaT from {0} to {1} or delete PatFile and extracted content.", archivoPaTenWIPconPath, DirDONE) & nl
            auxS &= "3) Process will continue. " & nl & nl
            AnyadetxtSalida(auxS)
            MsgBox(auxS, MsgBoxStyle.Information, "Cannot move PaT to _DONE directory!")
            AnyadetxtSalida("Process continues..." & nl)
        End Try
        AnyadetxtSalida("Done!" & nl)
        ' se ha movido, actulizao los ddw
        cargalbxWIP(LocalPatDir)
        cargalbxDONE(LocalPatDir)

        AnyadetxtSalida("Sending email..." & nl)

        ' correo
        ' necesito los parámetros
        auxS = ListaCarpetasTraducir.InfoEnvioCorreo
        auxS = DecryptWithKey(auxS, par.DestLongInterno)
        Dim IEC As New classInfoEnvioCorreo
        IEC = IEC.DeSerialize(auxS)
        If IEC.TodoOk = False Then
            auxS = "Cannot fetch mail info in the manifest file. Process continues normally but email will not be sent."
            MsgBox(auxS, MsgBoxStyle.Information, "Email will not be sent")
        Else ' enviamos el correo

            Dim destinatario As String = CorreoGestor
            AnyadetxtSalida(String.Format("Sending note to {0}... ", destinatario & nl))
            Dim mailssl As Boolean = IEC.MailSSL
            Dim mailpuerto As Integer = IEC.MailPuerto
            Dim mailcontrase As String = DecryptWithKey(IEC.MailContrase, par.DestLongInterno)
            Dim mailhost As String = IEC.MailHost
            Dim mailusuario As String = IEC.MailUsuario
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
            msg.ReplyTo = New MailAddress(par.DestCorreo)
            msg.IsBodyHtml = "true"
            '
            ' creo que solo es válido con "PATH" y "SHARED_DRIVE"
            Dim lin2() As String
            Select Case par.DestTipoDestino
                Case "FTP"
                    msg.Subject = String.Format("PaT name: {0} / From: {1}({2}) / In: {3}", NombrePatsinExtension, par.DestNombre, par.DestCorreo, par.DestFTPHost & "\" & par.DestFTPNombre)
                    lin2 = {
               "<HTML><BODY>",
               "Submitter " & par.DestNombre,
               "Target type:" & DirOutFXZTipoDestino,
               "Pat finished. You will find the translated folders in:",
               "ftp: " & par.DestFTPNombre & "\" & par.DestFTPNombre,
                "Target dir in PM host: " & ListaCarpetasTraducir.DirectorioGestor,
                "Go to the ftp directory -- OR -- load manifest in ZZ-Pat"
                }

                Case Else
                    msg.Subject = String.Format("PaT name: {0} / From: {1}({2}) / In: {3}", NombrePatsinExtension, par.DestNombre, par.DestCorreo, DirOutFXZ)
                    lin2 = {
            "<HTML><BODY>",
            "Submitter " & par.DestNombre,
            "Target type:" & DirOutFXZTipoDestino,
            "Pat finished. You will find the translated folders in:",
            "Exported folders in: " & DirOutFXZ,
            "Target dir in PM host: " & ListaCarpetasTraducir.DirectorioGestor,
            "Go to the exported folder directory -- OR -- load manifest in ZZ-Pat"
            }
            End Select

            For s = 0 To lin2.Count - 1
                msg.Body &= lin2(s) & "<br>"
            Next
            msg.Body &= "</BODY></HTML>"

            ' Voy a intentar añadir el html
            Dim arhcivoMFT_XML As String = String.Format("{0}\{1}\{1}_MFT.XML", LocalPatDir, NombrePatsinExtension)
            Dim arhcivoMFT_HTM As String = String.Format("{0}\{1}\{1}_MFT.HTM", LocalPatDir, NombrePatsinExtension)
            If File.Exists(arhcivoMFT_XML) Then
                Dim datos As New Attachment(arhcivoMFT_XML)
                msg.Attachments.Add(datos)

            End If
            If File.Exists(arhcivoMFT_HTM) Then
                msg.Body &= File.ReadAllText(arhcivoMFT_HTM)
                Dim datos As New Attachment(arhcivoMFT_HTM)
                msg.Attachments.Add(datos)
            End If



            Try
                If NoenviarCorreo Then
                    AnyadetxtSalida("Email not sent. Debug flag.")
                Else
                    client.Send(msg)
                    AnyadetxtSalida("Done!" & nl)
                End If
            Catch ex As Exception
                AnyadetxtSalida(String.Format("Error sending note: {1}", ex.Message) & nl & "Proces continues..." & nl)
            End Try
            msg.Dispose()

        End If
        ' ahora envío el correo si no hay error




        auxS = "The folders:" & nl & nl
        For Each row In ListaCarpetasTraducir.tCarpetasTraducir.Rows
            auxS &= row("carpeta") & nl
        Next
        auxS &= nl & "have been sent sucessfully. " & nl & nl
        If IEC.TodoOk Then
            auxS &= String.Format("Also the PaT owner has been warned with an email to {0} ", CorreoGestor) & nl & nl
            auxS &= "Package owner shall send contact you soon in order to solve any problem or in order to provide you the invoice information." & nl
        Else
            auxS &= String.Format("We have not been able to sent an email to {0} ", CorreoGestor) & nl & nl
            auxS &= "Plse inform him/her." & nl
        End If
        MsgBox(auxS, MsgBoxStyle.Information, "Folders successfully sent without any errors...")

        AnyadetxtSalida("Done!" & nl)
        AnyadetxtSalida("Folders successfully sent without any errors..." & nl)
        AnyadetxtSalida("*******************************************************" & nl)
        ' If True Then dt.Start() : AnyadetxtSalida("Timer started..." & nl) : Exit Sub



    End Sub

    Private Sub btnMoverAWIP_Click(sender As Object, e As RoutedEventArgs) Handles btnMoverAWIP.Click
        ' primero ver si me han seleccionado algo
        Dim NombrePat As String = ""
        If lbxDone.SelectedIndex < 0 Then
            MsgBox("In order to move a PaT from WIP to Done, first you must select one in the listbox. Select one and try again.", MsgBoxStyle.Exclamation, "No PaT to return")
            AnyadetxtSalida("No PaT selected. Abort. Timer restarted" & nl)
            Exit Sub
        Else ' tenemos uno
            NombrePat = lbxDone.SelectedItem.ToString
        End If
        '
        Dim par As structParametros
        par = CargaParametros()
        Dim localpatdir As String = par.DestLocalPatDir
        ' ahora muevo el pat de _WIP a _DONE
        Dim archivoPaTenWIPconPath As String = localpatdir & "\_WIP\" & NombrePat & ".PaT.zip"
        Dim archivoPaTenDONEconPath As String = localpatdir & "\_DONE\" & NombrePat & ".PaT.zip"
        Dim DirWIP As String = localpatdir & "\_WIP"

        AnyadetxtSalida("Working In Progress (_WIP) directoy does not exist. Creating " & DirWIP & "...")
        Try
            If Directory.Exists(DirWIP) = False Then
                auxS = String.Format("Cannot create WIP dir ({0}). ", DirWIP) & nl & nl ' por si hay error
                Directory.CreateDirectory(DirWIP)
                AnyadetxtSalida("Done!")
            End If
            auxS = String.Format("Cannot delete file {0}", archivoPaTenWIPconPath) & nl & nl ' por si hay error
            If File.Exists(archivoPaTenWIPconPath) Then File.Delete(archivoPaTenWIPconPath)
            auxS = String.Format("Cannot move from {0} to {1}.", archivoPaTenDONEconPath, DirWIP) & nl & nl ' por si hay un error
            File.Move(archivoPaTenDONEconPath, archivoPaTenWIPconPath)
        Catch ex As Exception
            auxS &= String.Format("Excep. {0}", ex.Message.ToString) & nl
            auxS &= "1) Folders have been correctly sent. You do not need to send them again." & nl
            auxS &= String.Format("2) Move manually PaT from {0} to {1} or delete PatFile and extracted content.", archivoPaTenDONEconPath, DirWIP) & nl
            auxS &= "3) Process will continue. " & nl & nl
            AnyadetxtSalida(auxS)
            MsgBox(auxS, MsgBoxStyle.Information, "Cannot move PaT to _WIP directory!")
            AnyadetxtSalida("Process continues..." & nl)
        End Try
        ' se ha movido, actulizao los ddw
        cargalbxWIP(localpatdir)
        cargalbxDONE(localpatdir)
    End Sub

    Private Sub btnMoverADone_Click(sender As Object, e As RoutedEventArgs) Handles btnMoverADone.Click
        Dim NombrePat As String = ""
        If lbxWIP.SelectedIndex < 0 Then
            MsgBox("In order to move a PaT from Done to WIP, first you must select one in the listbox. Select one and try again.", MsgBoxStyle.Exclamation, "No PaT to return")
            AnyadetxtSalida("No PaT selected. Abort. Timer restarted" & nl)
            Exit Sub
        Else ' tenemos uno
            NombrePat = lbxWIP.SelectedItem.ToString
        End If
        '
        Dim par As structParametros
        par = CargaParametros()
        Dim localpatdir As String = par.DestLocalPatDir
        ' ahora muevo el pat de _WIP a _DONE
        Dim archivoPaTenWIPconPath As String = localpatdir & "\_WIP\" & NombrePat & ".PaT.zip"
        Dim archivoPaTenDONEconPath As String = localpatdir & "\_DONE\" & NombrePat & ".PaT.zip"
        Dim DirDONE As String = localpatdir & "\_DONE"

        AnyadetxtSalida("Done (_DONE) directoy does not exist. Creating " & DirDONE & "...")
        Try
            If Directory.Exists(DirDONE) = False Then
                auxS = String.Format("Cannot create DONE dir ({0}). ", DirDONE) & nl & nl ' por si hay error
                Directory.CreateDirectory(DirDONE)
                AnyadetxtSalida("Done!")
            End If
            auxS = String.Format("Cannot delete file {0}", archivoPaTenDONEconPath) & nl & nl ' por si hay error
            If File.Exists(archivoPaTenDONEconPath) Then File.Delete(archivoPaTenDONEconPath)
            auxS = String.Format("Cannot move from {0} to {1}.", archivoPaTenWIPconPath, DirDONE) & nl & nl ' por si hay un error
            File.Move(archivoPaTenWIPconPath, archivoPaTenDONEconPath)
        Catch ex As Exception
            auxS &= String.Format("Excep. {0}", ex.Message.ToString) & nl
            auxS &= "1) Folders have been correctly sent. You do not need to send them again." & nl
            auxS &= String.Format("2) Move manually PaT from {0} to {1} or delete PatFile and extracted content.", archivoPaTenDONEconPath, DirDONE) & nl
            auxS &= "3) Process will continue. " & nl & nl
            AnyadetxtSalida(auxS)
            MsgBox(auxS, MsgBoxStyle.Information, "Cannot move PaT to _WIP directory!")
            AnyadetxtSalida("Process continues..." & nl)
        End Try
        ' se ha movido, actulizao los ddw
        cargalbxWIP(localpatdir)
        cargalbxDONE(localpatdir)
    End Sub

    Private Sub btnMoverAWIP_Copy_Click(sender As Object, e As RoutedEventArgs) Handles btnMoverAWIP_Copy.Click
        ' primero ver si me han seleccionado algo
        Dim NombrePat As String = ""
        If lbxWIP.SelectedIndex < 0 Then
            MsgBox("In order to move a PaT from Done to WIP, first you must select one in the listbox. Select one and try again.", MsgBoxStyle.Exclamation, "No PaT to return")
            AnyadetxtSalida("No PaT selected. Abort. Timer restarted" & nl)
            Exit Sub
        Else ' tenemos uno
            NombrePat = lbxWIP.SelectedItem.ToString
        End If
        '
        Dim par As structParametros
        par = CargaParametros()
        Dim localpatdir As String = par.DestLocalPatDir
        ' ahora muevo el pat de _WIP a _DONE
        Dim archivoPaTenWIPconPath As String = localpatdir & "\_WIP\" & NombrePat & ".PaT.zip"
        Dim archivoPaTenDONEconPath As String = localpatdir & "\_DONE\" & NombrePat & ".PaT.zip"
        Dim DirDONE As String = localpatdir & "\_DONE"

        AnyadetxtSalida("Done (_DONE) directoy does not exist. Creating " & DirDONE & "...")
        Try
            If Directory.Exists(DirDONE) = False Then
                auxS = String.Format("Cannot create DONE dir ({0}). ", DirDONE) & nl & nl ' por si hay error
                Directory.CreateDirectory(DirDONE)
                AnyadetxtSalida("Done!")
            End If
            auxS = String.Format("Cannot delete file {0}", archivoPaTenDONEconPath) & nl & nl ' por si hay error
            If File.Exists(archivoPaTenDONEconPath) Then File.Delete(archivoPaTenDONEconPath)
            auxS = String.Format("Cannot move from {0} to {1}.", archivoPaTenWIPconPath, DirDONE) & nl & nl ' por si hay un error
            File.Move(archivoPaTenWIPconPath, archivoPaTenDONEconPath)
        Catch ex As Exception
            auxS &= String.Format("Excep. {0}", ex.Message.ToString) & nl
            auxS &= "1) Folders have been correctly sent. You do not need to send them again." & nl
            auxS &= String.Format("2) Move manually PaT from {0} to {1} or delete PatFile and extracted content.", archivoPaTenDONEconPath, DirDONE) & nl
            auxS &= "3) Process will continue. " & nl & nl
            AnyadetxtSalida(auxS)
            MsgBox(auxS, MsgBoxStyle.Information, "Cannot move PaT to _WIP directory!")
            AnyadetxtSalida("Process continues..." & nl)
        End Try
        ' se ha movido, actulizao los ddw
        cargalbxWIP(localpatdir)
        cargalbxDONE(localpatdir)
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As RoutedEventArgs) Handles btnDelete.Click


        ' primero ver si me han seleccionado algo
        AnyadetxtSalida("***************************************" & nl)
        AnyadetxtSalida("Starting delete process..." & nl)
        dt.Stop() : AnyadetxtSalida("Timer stoped." & nl)
        Dim auxi As Integer = 0
        Dim NombrePat As String = ""
        If lbxDone.SelectedIndex < 0 Then
            MsgBox("In order to delete a PaT file and the unpackaged files, first you must select one in the listbox. Select one and try again.", MsgBoxStyle.Exclamation, "No PaT to return")
            AnyadetxtSalida("No PaT selected. Abort. " & nl)
            ' dt.Start() : AnyadetxtSalida("Timer restarted." & nl)
            Exit Sub
        Else ' tenemos uno
            NombrePat = lbxDone.SelectedItem.ToString
        End If
        ' primero el directory
        Dim par As structParametros
        par = CargaParametros()
        Dim localpatdir As String = par.DestLocalPatDir
        Dim DirABorrar As String = localpatdir & "\" & NombrePat
        Dim ArchivoABorrar As String = localpatdir & "\_DONE\" & NombrePat & ".PaT.zip"
        auxS = String.Format("You are about to delete the PaT file {0} and the directory {1}.", ArchivoABorrar, DirABorrar) & nl & nl
        auxS &= "Are you sure?"
        auxi = MsgBox(auxS, MsgBoxStyle.YesNo, "Delete PaT and extracted PaT files")
        If auxi = MsgBoxResult.No Then
            ' dt.Start() : AnyadetxtSalida("Usuer abort delete. Timer restarted." & nl)
            Exit Sub
        End If

        Try
            If Directory.Exists(DirABorrar) Then
                BorraDirectorios(DirABorrar)
            End If
            AnyadetxtSalida(String.Format("Dir {0} deleted." & nl, DirABorrar))
        Catch ex As Exception
            AnyadetxtSalida(String.Format("Error trying to delete dir {0} (Excp. {1})" & nl, DirABorrar, ex.Message.ToString) & nl)
            AnyadetxtSalida("Process continues..." & nl)
        End Try
        Try
            File.Delete(ArchivoABorrar)
            AnyadetxtSalida(String.Format("File {0} deleted." & nl, ArchivoABorrar))
        Catch ex As Exception
            AnyadetxtSalida(String.Format("Error trying to delete file {0} (Excp. {1})" & nl, ArchivoABorrar, ex.Message.ToString) & nl)
            AnyadetxtSalida("Process continues..." & nl)
        End Try
        cargalbxDONE(localpatdir)
        cargalbxDONE(localpatdir)
        ' dt.Start() : AnyadetxtSalida("Timer restarted." & nl)
        AnyadetxtSalida("Pat delete process finished succesfully!" & nl)
    End Sub
    ' subs para borrar directorios 
    Sub BorraDirectorios(via As String)
        For Each archivo As String In Directory.GetFiles(via)
            AnyadetxtSalida(archivo & "...")
            File.Delete(archivo)
            AnyadetxtSalida(" deleted!" & nl)
        Next
        For Each subcarpeta As String In Directory.GetDirectories(via)
            BorraDirectorios(subcarpeta) ' recursivo
        Next
        AnyadetxtSalida(String.Format("Directory {0}...", via))
        Directory.Delete(via)
        AnyadetxtSalida(" deleted!" & nl)
    End Sub
    Private Sub btnAhora_Click(sender As Object, e As RoutedEventArgs) Handles btnAhora.Click
        ' dt.Interval = New TimeSpan(0, 0, 1) ' tiempo bucle variable global 
        IntentaProcesar()
    End Sub
End Class


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

Public Class ClaseListaCarpetasTraducir
    ' ojo este código está en ZZ-CommonCode/ModuloRutPat y ZZ-Pat/MainWindows
    Public todoOK As Boolean
    Public tCarpetasTraducir As DataTable
    Public nCarpetas As Integer
    Public NombrePaT As String
    Public FechaEntrega As Date
    Public Notas As String
    Public ListaReferencias As ClaseListaReferencias
    Public FechaCreacionPaT As Date
    Public FirmaInstancia As Integer
    Public CorreoGestor As String
    Public DirectorioGestor As String
    Public NombreTraductor As String
    Public FlagDeleteFIfExists As Boolean
    Public FlagAutoInstall As Boolean
    Public InfoEnvioCorreo As String
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
        tCarpetasTraducir.Columns.Add(New DataColumn("bPreAna", Type.GetType("System.Boolean")))
        tCarpetasTraducir.Columns.Add(New DataColumn("PreAna", Type.GetType("System.Int32")))
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

    Sub AnyadeCarpeta(carpeta As String, ProBase As String, Perfil As String,
                          bCNT As Boolean, CNT As Integer,
                          bIniCal As Boolean, IniCal As Integer,
                          bPreAna As Boolean, PreAna As Integer,
                          Idioma As String, envio As String,
                          tarifa As Single, proMSS As String, ProIBM As String)

        Dim fila As DataRow = tCarpetasTraducir.NewRow
        fila("Carpeta") = carpeta
        fila("ProBase") = ProBase
        fila("ProMSS") = proMSS
        fila("ProIBM") = ProIBM
        fila("Perfil") = Perfil
        fila("Idioma") = Idioma
        fila("bCNT") = bCNT
        fila("CNT") = CNT
        fila("bIniCal") = bIniCal
        fila("IniCal") = IniCal
        fila("bPreAna") = bPreAna
        fila("PreAna") = PreAna
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



' esta extensión para el refresh

Module ExtensionMethods
    ' de http://geekswithblogs.net/NewThingsILearned/archive/2008/08/25/refresh--update-wpf-controls.aspx
    '
    ' No anonymous delegate in VB.NET, so have to empty method
    Private Sub EmptyMethod()

    End Sub

    <Extension()>
    Public Sub Refresh(ByVal uiElement As UIElement)
        uiElement.Dispatcher.Invoke(DispatcherPriority.Render, New Action(AddressOf EmptyMethod))
    End Sub



End Module


