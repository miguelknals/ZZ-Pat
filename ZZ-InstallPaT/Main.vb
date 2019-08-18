Imports System.Xml
Imports System.Xml.Serialization
Imports System.IO
Imports System.Text.RegularExpressions
Module Main
    Dim debuga As Boolean = True
    Dim DirTMP As String
    Dim NombrePaTconPath As String
    Dim NombrePatsinExtension As String
    Dim NombrePat As String
    Dim DirOutPat As String
    Dim DirEjeucion As String
    Sub Main()

        Console.WriteLine("ZZ-Instala PaT (c) 2016 miguel canals")
        Dim auxs As String = "" ' auxs
        ' leer la consola 
        Dim TengoNombrePaT As Boolean = False
        Dim TengoDirOutPat As Boolean = False

        Dim result As Integer = MsgBox("Test", MsgBoxStyle.Information)
        Dim args As String() = Environment.GetCommandLineArgs()
        For i = 1 To args.GetUpperBound(0) - 1
            Try
                Dim CurArg() As Char = args(i).ToCharArray(0, args(i).Length)
                If (CurArg(0) = "-") Or (CurArg(0) = "/") Then
                    Select Case Char.ToLower(CurArg(1), System.Globalization.CultureInfo.CurrentCulture)
                        Case "p"        ' archivo PaT
                            i = i + 1
                            NombrePaTconPath = args(i)
                            TengoNombrePaT = True
                        Case "e"        ' archivo PaT
                            i = i + 1
                            DirOutPat = args(i)
                            TengoDirOutPat = True
                        Case Else
                            AyudaZZInstallPaT()
                            If debuga Then GoTo fin
                            Exit Sub
                    End Select
                End If
            Catch e As Exception
                AyudaZZInstallPaT()
                Exit Sub
            End Try
        Next
        ' me tienen que haber especificado el nombre del pat
        Dim TengoError As Boolean = False
        If TengoNombrePaT = False Then
            Console.WriteLine("You must to specify a PaT file, i.e. '-p c:\tmp\MyPat.PaT'")
            TengoError = True
        Else
            Console.WriteLine("Pat File: {0}", NombrePaTconPath)
        End If
        ' Necesito NombrePatsinExtension
        NombrePat = Path.GetFileName(NombrePaTconPath)
        ' quitarle la extensión es .PaT.zip (case insesitive)
        NombrePatsinExtension = NombrePat
        NombrePatsinExtension = Replace(NombrePatsinExtension, ".zip", "", 1, , Microsoft.VisualBasic.CompareMethod.Text)
        NombrePatsinExtension = Replace(NombrePatsinExtension, ".pat", "", 1, , Microsoft.VisualBasic.CompareMethod.Text)



        If TengoDirOutPat = False Then
            Console.WriteLine("You must to specify an extract output PaT directory, i.e. '-e  C:\tmp\outPat'")
            TengoError = True
        Else
            DirOutPat &= "\" & NombrePatsinExtension
            Console.WriteLine("Pat File extract directory: {0}", DirOutPat)
        End If



        If TengoError Then
            AyudaZZInstallPaT()
            If debuga Then GoTo fin
            Exit Sub
        End If
        ' antes de hacer nada vamos a ver que existe el archivo
        ' directorio temporal
        DirTMP = Path.GetTempPath() & "ZZ\"
        Console.WriteLine("Temporary dir {0}", DirTMP)
        ' directorio ejecución
        DirEjeucion = System.Reflection.Assembly.GetExecutingAssembly().Location ' ubicación ejecutable
        DirEjeucion = Path.GetDirectoryName(DirEjeucion)
        ' desempaqueto
        Dim PreZip = DirEjeucion & "\unzip.exe"
        Dim mdto As String = PreZip & String.Format(" -o {0} -d {1} ", NombrePaTconPath, DirOutPat)
        Dim rc As Integer = 0
        Console.WriteLine("Unziping {0} in {1}...", NombrePaTconPath, DirOutPat)
        mandatoZip(mdto, rc)
        If rc <> 0 Then
            Console.WriteLine("Unzip command ('{0}' exited with rc {1}", mdto, rc)
            If debuga Then GoTo fin
            Exit Sub
        End If
        Console.WriteLine("Done")
        ' ahora tengo que leer el manifiesto
        ' voy a leer los DirOutPat _MFT.XML"
        Dim archivoMFT As String = ""
        Dim archivoMFTconPath As String = ""
        Dim fileEntries As String() = Directory.GetFiles(DirOutPat)
        For Each archivo In fileEntries
            If archivo.EndsWith("_MFT.XML") Then
                archivoMFT = Path.GetFileName(archivo)
                archivoMFTconPath = archivo
            End If
        Next
        ' si no lo tengo error fatal
        If archivoMFT = "" Then
            Console.WriteLine("Cannot find the manifest file  *_MTF.XML")
            If debuga Then GoTo fin
            Exit Sub
        End If
        Console.WriteLine("MFT file: {0}", archivoMFT)
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

            Console.WriteLine("Fatal error: Looks like xml manifest is not valid. ")
            Console.WriteLine("excp: {0}", ex.Message)
            If debuga Then GoTo fin
            Exit Sub
        End Try
        ' he leido el archivo pat
        Dim row As DataRow
        Console.WriteLine("Folders in PaT")
        For Each row In ListaCarpetasTraducir.tCarpetasTraducir.Rows
            Console.WriteLine(row("carpeta"))
        Next
        ' debería comprobar si OTM está instalado
        Dim ViaSistema = Environment.GetEnvironmentVariable("PATH")
        'voy a buscar C:\OTM\WIN
        Dim auxi As Integer = InStr(ViaSistema, ":\OTM\WIN")

        If auxi = 0 Then ' OTM parece no instalado 
            ' para que el splashscreen no se me coma el diálogo
            Console.WriteLine("Cannot find :\OTM\WIN in PATH variable. Looks like OTM is no installed.")
            If debuga Then GoTo fin
            Exit Sub
        End If
        Dim DiscoOTM As String = ViaSistema.Substring(auxi - 2, 1)
        Dim mandato As String = ""

        Dim carpeta As String = ""
        For Each row In ListaCarpetasTraducir.tCarpetasTraducir.Rows
            Console.WriteLine(row("carpeta"))
            carpeta = row("carpeta") ' p.e c:\kkk\kkk.fxp

            Dim fromDrive As String = Path.GetPathRoot(DirOutPat)
            Dim fromPath As String = DirOutPat : fromPath = Replace(fromPath, fromDrive, "")
            If fromPath.Substring(0, 1) <> "\" Then fromPath = "\" & fromPath
            fromDrive = fromDrive.Substring(0, 1) ' la letra de c:\
            Dim FLD As String = carpeta

            mandato = " /TAsk=FLDIMP /FLD={0} /FROMdrive={1} /FromPath={2} /OPtions=(MEM,DICT) " ' /QUIET=NOMSG  "
            mandato = String.Format(mandato, FLD, fromDrive, fromPath, DiscoOTM, DirOutPat) ' 
            ejecuta(mandato, rc)
            If rc <> 0 Then
                Console.WriteLine("OPS {0} cannot be imported. RC {1} ", carpeta, rc)


            End If


        Next



        'Dim unidad As String = Path.GetPathRoot(Par.ParDirTemporal).Substring(0, 1) ' la letra
        'auxs = Path.GetPathRoot(Par.ParDirTemporal) ' C:\
        'auxs = auxs.Replace("\", "") ' C: que es lo que debo quitar
        'Dim DirTemporalSinUnidad As String = Par.ParDirTemporal.Replace(auxs, "")
        'mandato = " /TAsk=FLDEXP /FLD={0} /TOdrive={1} /ToPath={2} /OPtions=(MEM,ROMEM,DOCMEM) /OVerwrite=YES /QUIET=NOMSG  "
        'mandato = String.Format(mandato, carpeta, unidad, DirTemporalSinUnidad) ' exportaré al directorio temporal
        'ejecuta(mandato, rc)


        ' vamos a ver si otm está en ejecución

        Process.Start(archivoMFTHTMLconPath)


        If OTMEnEjecucion() Then ' error fatal no haré nada si está en marcha
            Console.WriteLine("OTM is running. You need to close it.")
            If debuga Then GoTo fin
            Exit Sub
        End If


fin:
        If debuga Then Console.WriteLine("Press Intro to continue") : Console.ReadLine()



    End Sub

End Module
