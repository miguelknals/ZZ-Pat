Imports System
Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
Imports System.Data
Imports System.Text
'
Imports System.Threading
Imports System.Runtime.CompilerServices
Imports System.Windows.Threading
' 
Imports System.Security.Principal
Imports System.Net
'
Imports System.Net.Mail
Imports System.Collections.Generic
Imports System.Configuration ' to read app.settings
' 


Class MainWindow
    Dim debug As Boolean = False
    Dim nl As String = Environment.NewLine
    Dim dq As String = Chr(34) ' Double quote
    Dim ListaCarpetasTraducir As New ClaseListaCarpetasTraducir
    Dim ListaReferencias As New ClaseListaReferencias


    Dim ValorP As String
    '  Propiedades

    ' esta funcion no se usa es de pruebas
    Function GeneratesSS() As Long
        Dim n As Long
        n = 1234343444443434
        Dim value As Double = Math.Sin(Math.Log(n))
        Dim expo As Double = Math.Floor(Math.Log10(Math.Abs(value)))
        Dim mant As Double = Math.Abs(value / Math.Pow(10, expo))
        Dim longmant As Double = Convert.ToInt64(Math.Abs(mant * 10 ^ 14))
        Dim onethirdbytes As Byte() = BitConverter.GetBytes(longmant) ' 8 byte lets make it longer 
        Dim ClearBytes(onethirdbytes.Length * 3 - 1) As Byte
        For i = 0 To 2
            onethirdbytes.CopyTo(ClearBytes, onethirdbytes.Length * i)
        Next
        Dim clearBytes2 As Byte() = Encoding.Unicode.GetBytes("xxxxxxxxx")
        Dim kk As String = "22"
    End Function

    Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        ' nueva clase de lista de carpetas

        ' No he podido utilizar el mecanismo de MVS de elegir BuildActionSplsh 
        ' el problema es que si se utiliza, el splashscreen cierra la ventana
        ' aquí lo explican
        ' http://stackoverflow.com/questions/576503/how-to-set-wpf-messagebox-owner-to-desktop-window-because-splashscreen-closes-me
        ' 
        ' más de lo mismo en las soluciones alternativas en 
        ' https://connect.microsoft.com/VisualStudio/feedback/details/600197/wpf-splash-screen-dismisses-dialog-box
        '
        ' de momento he utilizado la de crear una ventana temporal
        ' la más limpia es quitar lo de BuildAction en el splashscreeny hacerlo manualmento con
        '
        '
        'Dim splash As New System.Windows.SplashScreen("mknals.jpg")
        'splash.Show(False)
        'System.Threading.Thread.Sleep(500)
        'splash.Close(New TimeSpan(0, 0, 0))
        '
        ' configuración

        ' Encripto si no lo está AppConfig
        ' al compilar, copia appconfig y la encripta la primera vez que se ejecuta
        ' El mecansimo es en App.Config colocamos la contraseña, la primera vez que se ejecuta, se copia
        ' al directorio de ejecución y EncriptAppConfig encripta la sección. Por lo tanto en el fuente NO
        ' debería 

        ' si hay un parámetro es el ususario. este código se utiliza tb en
        ' en la obtención del nombre de parámetro 
        Dim VersionGitHub = "20190818"
        Dim args() As String = Environment.GetCommandLineArgs()
        Dim usuario As String = ""
        If args.Count > 1 Then
            usuario = args(1)
        Else
            Dim temp As Window = New Window()
            temp.Visibility = Windows.Visibility.Hidden
            temp.Show()
            Dim Info As String = "You need to specify a user. Call as ZZ-PatV02.exe username, a ZZ-PatV02_PAR_username.XML will be created if not exist."
            MessageBox.Show(temp, Info, "Invoking this application, you must user a username", MessageBoxButton.OK, MessageBoxImage.Error)
            Application.Current.Shutdown()
            Exit Sub

        End If


        Dim ViaEjecutable As String = System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase
        ViaEjecutable = ViaEjecutable.Replace("file:///", "")
        'Dim fechaEjecutablee As String = File.GetCreationTime(System.Reflection.Assembly.GetExecutingAssembly().Location)
        VentanaPrincipal.Title = "ZZ-Pat " & File.GetLastWriteTime(ViaEjecutable).ToString & "  (" & ViaEjecutable & ")"
        ' lo anterior cambia al cambiar el archivo a otra máquina
        ' de http://stackoverflow.com/questions/804192/how-can-i-get-the-assembly-last-modified-date
        ' había alguna solución más radical
        Dim lastMod2 As String = File.GetLastWriteTime(System.Reflection.Assembly.GetExecutingAssembly().Location).ToString
        Dim arch As FileInfo = New System.IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location)
        Dim lastMod As DateTime = arch.LastWriteTime
        VentanaPrincipal.Title = "User: " & usuario & " - " & "ZZ-Pat " & lastMod & "  (" & ViaEjecutable & ")"

        btnTest.Visibility = Windows.Visibility.Hidden

        Dim InfoOTM As New ClassInfoOTM
        If InfoOTM.TodoOK = False Then
            Dim temp As Window = New Window()
            temp.Visibility = Windows.Visibility.Hidden
            temp.Show()
            MessageBox.Show(temp, InfoOTM.Info, "OTM Error", MessageBoxButton.OK, MessageBoxImage.Error)
            Application.Current.Shutdown()
            Exit Sub
        End If

        Dim DiscoOTM As String = InfoOTM.DiscoOTM
        CargaListaCarpetasDDW(InfoOTM.Carpetas)
        CargaListaPerfilesCBX(InfoOTM.Perfiles)

        ' en princpio CNT
        chkCNT.IsChecked = True ' si cnt
        chkIniCal.IsChecked = True ' no calculatin inical
        chkPreAna.IsChecked = True
        ' Fecha de entrega
        dpiFechaEntrega.SelectedDate = Date.Today

        ' valores

        '    Application.Current.Properties("MyApplicationScopeProperty") = "myApplicationScopePropertyValue"
        ' cargo los parámetros para visualizar el directorio de salida
        Dim Par As structParametros
        Par = CargaParametros()
        ' lblDirPaT.Text = Environment.ExpandEnvironmentVariables(Par.ParDirSalidaPaT)
        txtPatOutput.Text = Environment.ExpandEnvironmentVariables(Par.ParDirSalidaPaT)
        txtTarifa.Text = Par.ParTarifaPredeterminada
        ' ahora cargo los destinos
        Dim listadestinos As structListaDestinos
        listadestinos = CargaDestinos()
        cbxDestinoPaT.Items.Add("NONE")
        For Each destino In listadestinos.ListaDestino
            cbxDestinoPaT.Items.Add(destino.DestNombre)
        Next
        cbxDestinoPaT.SelectedIndex = 0 ' selecciono none
        ' miro si existe el directorio temporal, si no existe lo intento crear
        Dim auxS As String : auxS = Environment.ExpandEnvironmentVariables(Par.ParDirTemporal)
        Try
            If (Not Directory.Exists(auxS)) Then
                Directory.CreateDirectory(auxS)
                txtSalida.AppendText(String.Format("Created temporary dir {0}", auxS) & nl)
            End If
        Catch ex As Exception

        End Try

        ' 
        ' 
    End Sub

    Private Sub txtPatOutput_PreviewDragOver(sender As Object, e As DragEventArgs) Handles txtPatOutput.PreviewDragOver
        ' este evento hace falta en un text box según stack overflow
        e.Handled = True
    End Sub
    Private Sub txtPatOutput_DragEnter(sender As Object, e As DragEventArgs) Handles txtPatOutput.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop, False) Then
            e.Effects = DragDropEffects.All
        Else
            e.Effects = DragDropEffects.None
        End If
    End Sub

    Private Sub txtPatOutput_Drop(sender As Object, e As DragEventArgs) Handles txtPatOutput.Drop
        Dim files As String()
        files = e.Data.GetData(DataFormats.FileDrop, False)
        Dim filename As String

        For Each filename In files
            txtPatOutput.Text = Path.GetFullPath(filename)

            Exit For

        Next
    End Sub


    Private Sub lbxReferencia_DragEnter(sender As Object, e As DragEventArgs) Handles lbxReferencia.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop, False) Then
            e.Effects = DragDropEffects.All
        Else
            e.Effects = DragDropEffects.None
        End If
    End Sub

    Private Sub lbxReferencia_Drop(sender As Object, e As DragEventArgs) Handles lbxReferencia.Drop
        Dim files As String()
        files = e.Data.GetData(DataFormats.FileDrop, False)
        Dim filename As String


        For Each filename In files
            lbxReferencia.Items.Add(filename)
            ListaReferencias.Anyade(filename) ' ' actulizo la clase
        Next
    End Sub
    Private Sub btnBorrarReferencias_Click(sender As Object, e As RoutedEventArgs) Handles btnBorrarReferencias.Click
        ' borro y creo nuevo instancia
        lbxReferencia.Items.Clear()
        ListaReferencias = New ClaseListaReferencias
    End Sub

    Private Sub lbxCarpetas_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles lbxCarpetas.MouseDoubleClick
        ' solo adminto eventos del TextBlock del TextBock
        If Not TypeOf (e.OriginalSource) Is TextBlock Then
            Exit Sub
        End If
        ' Necesito los parámetros
        Dim Par As structParametros
        Par = CargaParametros()
        Dim carpeta As String = lbxCarpetas.SelectedValue
        Dim Perfil As String = cbxPerfiles.SelectedItem ' perfil
        ' primero voy a ver que no está duplicada
        Dim row As DataRow
        For Each row In ListaCarpetasTraducir.tCarpetasTraducir.Rows
            If carpeta = row("Carpeta") Then ' quiere decir que ya la tengo no la puedo añadir
                MsgBox(String.Format("Folder {0} Is already In the folder To be translated. ", carpeta), MsgBoxStyle.Exclamation)
                Exit Sub
            End If
        Next
        ' Carpetas que ya tengo en la lista
        ' Voy a crear un diccionario con todas las carpetas
        If Len(carpeta) < 5 Then ' demasiado pequeño el nombre
            MsgBox(String.Format("Folder {0} has less than 5 Char. We cannot process folders With less than 5 Char. ", carpeta), MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Dim DictProyectos As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)
        DictProyectos.Add(carpeta.Substring(0, 5), 0)
        For Each row In ListaCarpetasTraducir.tCarpetasTraducir.Rows
            Dim auxs As String
            auxs = row("Carpeta").Substring(0, 5)
            If Not (DictProyectos.ContainsKey(auxs)) Then
                DictProyectos.Add(auxs, 0)
            End If
        Next
        ' al llegar aquí si DictProyectos contiene el número de proyectos 
        If DictProyectos.Count() > 1 And Par.ParMasDeUnPB = False Then ' hay más de un proyecto y no está permitido por mas de un proyecto
            MsgBox("There are more than one folder proyect but only Is allowed. Check parameters To allow multiproject packages", MsgBoxStyle.Exclamation, "A PaT file can have folders only from base project ")
            Exit Sub
        End If
        ' voy a crear el nombre a partir de las entradas del diccionario.
        Dim proyectoBase As String = ""
        For Each item In DictProyectos
            proyectoBase += item.Key & "_"
        Next

        txtNombrePaT.Text = proyectoBase & DateTime.Now.ToString("yyyyMMdd_HHmm")
        'DateTime.Now.ToString("yyyyMMddHHmmss")
        Dim tarifa As Single = CType(txtTarifa.Text, Single) ' respeta el encoding
        ' primero añadiré la carpeta
        ' voy a intentar scacar el proycto de proveedor

        Dim proProveedor As String = "N/D"
        '
        Dim idioma As String = "N/D"
        Dim proIBM As String = "N/D"

        Dim inst As New strucProIdiomaCarpeta(carpeta)
        idioma = inst.LANG
        proIBM = inst.IU




        txtNotas.Text = Trim(txtNotas.Text) ' importante xml si no me da problemas

        ListaCarpetasTraducir.AnyadeCarpeta(carpeta, carpeta.Substring(0, 5), Perfil,
                                            chkCNT.IsChecked, 0, chkIniCal.IsChecked, 0,
                                            chkPreAna.IsChecked, 0,
                                            idioma, txtEnvio.Text, tarifa, proProveedor, proIBM)
        ' ahora acutalizao el DG

        'dgCarpetas.Items.Clear()
        dgCarpetas.ItemsSource = ListaCarpetasTraducir.tCarpetasTraducir.DefaultView
        dgCarpetas.IsReadOnly = True

    End Sub
    '
    ' Eventos drag-drop

    ' Para inicializar
    Sub CargaListaCarpetasDDW(LC As List(Of String))
        Dim s As String
        For Each s In LC
            lbxCarpetas.Items.Add(s)
        Next
    End Sub

    Sub CargaListaPerfilesCBX(LC As List(Of String))
        Dim s As String
        For Each s In LC
            cbxPerfiles.Items.Add(s)
        Next
        ' selecciono el primero
        cbxPerfiles.SelectedIndex = 0
    End Sub

    Function CompruebaAntibarra(s As String) As String
        If Len(s) > 1 Then
            If s.Substring(Len(s) - 1, 1) <> "\" Then
                s &= "\"
            End If
        End If
        Return s
    End Function

    Private Sub btnWCT_Click(sender As Object, e As RoutedEventArgs) Handles btnWCT.Click
        ' inicio algún contr

        AnyadetxtSalida("")

        Dim debuga As Boolean = False ' crea un log.txt en el directorio de ejecución con mandatos MainWindow
        If debuga Then logea("", True) ' borro el log

        ' antes de hacer nada, miro las condiciones
        ' 1) Necesito nombre
        Dim auxS As String = ""
        Dim PreZip As String = ""
        Dim DirEje = System.Reflection.Assembly.GetExecutingAssembly().Location
        DirEje = Path.GetDirectoryName(DirEje)  ' el directorio ejecución
        If debuga Then logea(String.Format("Dir ejecución ->'{0}'", DirEje))
        PreZip = String.Format("zip.exe") ' el path se añade al llamarlo
        ' ahora me interesa la parte ejecutable 
        Dim bCrearPaT As Boolean = True
        Dim NombrePaT As String = txtNombrePaT.Text
        Dim RC As Integer = 0
        ' ParDirSalidaPaT  es el directorio de trabajo
        Dim Par As structParametros
        Par = CargaParametros()
        ' voy a ver si me han especificado un nuevo directorio

        auxS = CompruebaAntibarra(Trim(txtPatOutput.Text))
        If Not Directory.Exists(auxS) Then
            auxS = "Output directory Is Not valid Or does Not exist. Pls correct it."
            MsgBox(auxS, MsgBoxStyle.Exclamation, "Wrong PaT output directory")
            txtSalida.AppendText(auxS & nl)
            Exit Sub
        End If
        If Par.ParDirSalidaPaT <> auxS Then
            ' El directorio del Pat ha cambiado lo guardo
            Par.ParDirSalidaPaT = auxS : txtPatOutput.Text = auxS
            GuardaParametros(Par)
        End If

        If bCrearPaT Then
            ' para evitar problema no quiero caracteres especiales


            If NombrePaT = "" Then
                'txtSalida.AppendText("Si se va a crear un PaT, debe haber un nombre " & nl)
                auxS = "If are going To create a PaT file, you must specify a filename."
                MsgBox(auxS, MsgBoxStyle.Exclamation, "No PaT filename")
                Exit Sub
                txtSalida.AppendText(auxS & nl)
            Else ' si existe lo borra
                Dim specialchar = " []#$%^&*()=|{}'<>/\?" ' allow dot
                If NombrePaT.IndexOfAny(specialchar.ToCharArray) <> -1 Then
                    auxS = "Pat name cannot contain blanks nor special list char " & specialchar
                    MsgBox(auxS, MsgBoxStyle.Exclamation, "PaT contains forbidden chars")
                    Exit Sub
                    txtSalida.AppendText(auxS & nl)
                Else
                    NombrePaT &= ".PaT.zip"
                    ' borro el paquete si existe. Aquí el path sale con comillas. No hace falta hacer nada.
                    If File.Exists(Environment.ExpandEnvironmentVariables(Par.ParDirSalidaPaT) & NombrePaT) Then
                        File.Delete(Environment.ExpandEnvironmentVariables(Par.ParDirSalidaPaT) & NombrePaT)
                    End If
                End If

            End If
            ' PaT has an standar name with no blanks


        End If
        txtSalida.AppendText("PaT filename: " & Environment.ExpandEnvironmentVariables(Par.ParDirSalidaPaT) & NombrePaT & nl)
        ' Necesitaré el traductor
        Dim NombreTraductor As String = cbxDestinoPaT.SelectedItem


        ' para actualizar la IU
        ' http://stackoverflow.com/questions/5192169/update-gui-using-backgroundworker
        ' 
        If ListaCarpetasTraducir.nCarpetas <= 0 Then
            txtSalida.AppendText("OPPS -> There are no folders" & nl)
            auxS = "There are not any folders to be translated. Select one o more folders."
            MsgBox(auxS, MsgBoxStyle.Exclamation, "Empty folders list")
            txtSalida.AppendText(auxS & nl)
            Exit Sub
        End If
        txtSalida.AppendText("Running folder calculating reports " & nl)
        Dim carpeta As String = ""
        Dim perfil As String = ""
        Dim PathPrefijo As String = Environment.ExpandEnvironmentVariables(Par.ParDirTemporal)
        Dim fila As DataRow
        Dim idioma As String = ""
        Dim ProIBM As String = ""
        Dim ProMSS As String = "N/A"
        Dim TengoProIBM As Boolean = False
        Dim archivoExcel = Par.ParArchivoExcel
        '        For Each fila In ListaCarpetasTraducir.tCarpetasTraducir.Rows
        '            ProIBM = fila("Carpeta") : TengoProIBM = False
        '            ' las carpetas son _SPA o _CAT '
        '            idioma = "N/D"
        '            ' el idioma supondré que hay _ y luego hasta el final.
        '            Dim posiGuion As Integer = 0
        '            For i As Integer = Len(ProIBM) - 1 To 0 Step -1
        '                If ProIBM.Substring(i, 1) = "_" Then
        '                    posiGuion = i
        '                    Exit For
        '                End If
        '            Next
        '            If posiGuion <> Len(ProIBM) - 1 Then
        '                idioma = ProIBM.Substring(posiGuion + 1, Len(ProIBM) - 1 - posiGuion)
        '                auxS = ProIBM.Substring(posiGuion, Len(ProIBM) - posiGuion)
        '                ProIBM = Replace(ProIBM, auxS, "")
        '                TengoProIBM = True
        '            End If
        '            If idioma = "SPA" Then idioma = "ES"
        '            If idioma = "CAT" Then idioma = "CA"
        '            fila("Idioma") = idioma
        '            ProMSS = "N/D"
        '            If TengoProIBM Then
        '                'ProMSS = LeeProyectoMS(ProIBM, idioma, archivoExcel, Par.ParSQL)
        '                fila("ProIBM") = ProIBM
        '
        '            Else
        '                fila("ProIBM") = "N/A"
        '            End If
        '            'fila("ProMSS") = ProMSS
        '        Next
        'ListaCarpetasTraducir.tCarpetasTraducir.AcceptChanges()


        For Each fila In ListaCarpetasTraducir.tCarpetasTraducir.Rows

            carpeta = UCase(fila("Carpeta"))
            perfil = fila("Perfil")
            auxS = String.Format("Folder {0} Profile {1}" & nl, carpeta, perfil)
            txtSalida.AppendText(auxS)
            Dim mandato As String = ""
            Dim archivosalida = ""
            Dim opcion As String = ""
            ' si tengo que crear el paquete exportaré la carpeta y haré el zip
            ' Voy a crear los directorios temporales 
            Try
                ' verificado para quote path. debería estar ok.
                If (Not System.IO.Directory.Exists(Environment.ExpandEnvironmentVariables(Par.ParDirTemporal))) Then
                    System.IO.Directory.CreateDirectory(Environment.ExpandEnvironmentVariables(Par.ParDirTemporal))
                    txtSalida.AppendText(String.Format("Created temporary dir {0}", Environment.ExpandEnvironmentVariables(Par.ParDirTemporal)) & nl)
                End If
                If (Not System.IO.Directory.Exists(Environment.ExpandEnvironmentVariables(Par.ParDirSalidaPaT))) Then
                    System.IO.Directory.CreateDirectory(Environment.ExpandEnvironmentVariables(Par.ParDirSalidaPaT))
                    txtSalida.AppendText(String.Format("Created output PaT dir {0}", Environment.ExpandEnvironmentVariables(Par.ParDirSalidaPaT)) & nl)
                End If
            Catch ex As Exception

            End Try
            If bCrearPaT Then
                ' el directorio de salida es ParDirTemporal p.e. C:\kk\kk2
                Dim unidad As String = Path.GetPathRoot(Environment.ExpandEnvironmentVariables(Par.ParDirTemporal)).Substring(0, 1) ' la letra
                auxS = Path.GetPathRoot(Environment.ExpandEnvironmentVariables(Par.ParDirTemporal)) ' C:\
                auxS = auxS.Replace("\", "") ' C: que es lo que debo quitar
                Dim DirTemporalSinUnidad As String = Environment.ExpandEnvironmentVariables(Par.ParDirTemporal).Replace(auxS, "")
                ' nota -> OTMBATCH maneja correcatamente los paths con espacios 
                mandato = " /TAsk=FLDEXP /FLD={0} /TOdrive={1} /ToPath={2} /OPtions=(MEM,ROMEM,DOCMEM) /OVerwrite=YES /QUIET=NOMSG  "
                mandato = String.Format(mandato, carpeta, unidad, DirTemporalSinUnidad) ' exportaré al directorio temporal
                ejecuta(mandato, RC, debuga) ' debuga es opcional
                If RC <> 0 Then
                    auxS = "Folder " & carpeta & " cannot be exported in directory " & Environment.ExpandEnvironmentVariables(Par.ParDirTemporal) & ". Verify all memories are in place. " &
                           "RC=" & RC.ToString
                    MsgBox(auxS, MsgBoxStyle.Exclamation, "Folder cannot be exported" & nl)
                    txtSalida.AppendText(auxS & nl)
                    Exit Sub
                End If
                'txtSalida.AppendText("Carpeta " & carpeta & " exporada en " & Par.ParDirTemporal & nl)
                AnyadetxtSalida("Folder " & carpeta & " exported in " & Environment.ExpandEnvironmentVariables(Par.ParDirTemporal))
                ' ahora el zip
                auxS = PreZip & " -j " & String.Format("""{0}{1}"" ""{2}{3}"" ",
                                     Environment.ExpandEnvironmentVariables(Par.ParDirSalidaPaT),
                                     NombrePaT,
                                     Environment.ExpandEnvironmentVariables(Par.ParDirTemporal),
                                     carpeta & ".FXP")

                If debug Then logea("Zipping FXP ->  " & auxS)
                mandatoZip(auxS, RC, , debuga) ' DirEjecución y debuga son opcionales
                ' lo borro
                File.Delete(Environment.ExpandEnvironmentVariables(Par.ParDirTemporal) & carpeta & ".FXP")
            End If


            If fila("bCNT") Then
                archivosalida = PathPrefijo & carpeta & ".CNT"
                Dim ContajeCT As New structContajeCNT(archivosalida, carpeta)
                ContajeCT.calcula()
                If ContajeCT.todoOk Then
                    fila("CNT") = ContajeCT.contaje
                    If bCrearPaT Then
                        ' zip kkk.zip arhivo (añade el archivo al zip) D es para que no ponga el path
                        auxS = PreZip & " -j " & String.Format("""{0}{1}"" ""{2}"" ",
                                     Environment.ExpandEnvironmentVariables(Par.ParDirSalidaPaT),
                                     NombrePaT,
                                     archivosalida)
                        If debug Then logea("Zipping CNT ->  " & auxS)
                        mandatoZip(auxS, RC, , debuga) ' el rc no hace nada
                    End If
                    File.Delete(archivosalida)

                Else
                    MsgBox(ContajeCT.MsgCorpus, MsgBoxStyle.Information, ContajeCT.MsgTitle)
                End If
                txtSalida.AppendText(ContajeCT.salida)
            End If

            archivosalida = PathPrefijo & String.Format("{0}_{1}_ini_cal.rpt", carpeta, perfil)
            Dim ContajeCalculating As New structContajeCalculating(archivosalida, carpeta, perfil)
            ' el CNT detallado es opcion = "TMMATCH"
            If fila("bIniCal") Then
                ContajeCalculating.calcula()
                If ContajeCalculating.todoOk Then
                    fila("IniCal") = ContajeCalculating.contaje
                    If bCrearPaT Then
                        ' zip kkk.zip arhivo (añade el archivo al zip) D es para que no ponga el path
                        auxS = PreZip & " -j " & String.Format("""{0}{1}"" ""{2}"" ",
                                     Environment.ExpandEnvironmentVariables(Par.ParDirSalidaPaT),
                                     NombrePaT,
                                     archivosalida)
                        If debug Then logea("zipping INI CAL->  " & auxS)
                        mandatoZip(auxS, RC, , debuga) ' el rc no hace nada

                    End If
                    File.Delete(archivosalida)
                Else
                    MsgBox(ContajeCalculating.MsgCorpus, MsgBoxStyle.Information, ContajeCalculating.MsgTitle)
                End If
                txtSalida.AppendText(ContajeCalculating.salida)
            End If


            ' ahora voy con los calculating 
            ' el preanalisis
            archivosalida = PathPrefijo & String.Format("{0}_{1}_preana.rpt", carpeta, perfil)
            Dim ContajePreAnalisis As New structContajePreAnalisis(archivosalida, carpeta, perfil)
            If fila("bPreAna") Then
                ContajePreAnalisis.calcula()
                If ContajePreAnalisis.todoOk Then
                    fila("PreAna") = ContajePreAnalisis.contaje
                    If bCrearPaT Then
                        ' zip kkk.zip arhivo (añade el archivo al zip) D es para que no ponga el path
                        auxS = PreZip & " -j " & String.Format("""{0}{1}"" ""{2}"" ",
                                     Environment.ExpandEnvironmentVariables(Par.ParDirSalidaPaT),
                                     NombrePaT,
                                     archivosalida)
                        If debug Then logea("zipping PREANA->  " & auxS)
                        mandatoZip(auxS, RC, , debuga) ' el rc no hace nada
                    End If
                    File.Delete(archivosalida)
                Else
                    MsgBox(ContajePreAnalisis.MsgCorpus, MsgBoxStyle.Information, ContajePreAnalisis.MsgTitle)
                End If
                txtSalida.AppendText(ContajePreAnalisis.salida)
            End If



        Next
        ListaCarpetasTraducir.tCarpetasTraducir.AcceptChanges()
        dgCarpetas.ItemsSource = ListaCarpetasTraducir.tCarpetasTraducir.DefaultView
        dgCarpetas.IsReadOnly = True

        ' acutlizao la ListaCarpetasTraducir

        With ListaCarpetasTraducir
            .NombrePaT = NombrePaT
            If dpiFechaEntrega.SelectedDate Is Nothing Then
                .FechaEntrega = Nothing
            Else
                .FechaEntrega = dpiFechaEntrega.SelectedDate
            End If
            .Notas = txtNotas.Text.Replace(Environment.NewLine, "<br>")
            .ListaReferencias = ListaReferencias
            .FechaCreacionPaT = Now
            .CorreoGestor = Par.ParCorreoGestor
            auxS = Environment.ExpandEnvironmentVariables(Par.ParDirSalidaPaT) : If Len(auxS) > 1 And auxS.EndsWith("\") Then auxS = auxS.TrimEnd("\")
            .DirectorioGestor = auxS
            .NombreTraductor = NombreTraductor
            ' Debo encapsular y cifrar la info de envío de correo
            Dim InfoEnvioCorreo As New classInfoEnvioCorreo
            InfoEnvioCorreo.MailContrase = Par.ParMailContrase
            InfoEnvioCorreo.MailHost = Par.ParMailHost
            InfoEnvioCorreo.MailPuerto = Par.ParMailPuerto
            InfoEnvioCorreo.MailSSL = Par.ParMailSSL
            InfoEnvioCorreo.MailUsuario = Par.ParMailUsuario
            auxS = InfoEnvioCorreo.Serialize
            .InfoEnvioCorreo = EncryptWithKey(auxS, Par.ParLongInterno)
            .FirmaInstancia = .ObtenHash(Par.ParLongInterno) ' esta línea DEBE ser la última
        End With

        ' El informe de TXT y la consola

        ' voy a por el informe
        Dim sbSalidaMF As New StringBuilder("")
        Dim lin() As String = {
            "PaT naame: {0}" & nl,
            "Project manager: {0}" & nl,
            "Due daate: {0}" & nl,
            "Notes:" & nl & " {0}" & nl,
            "Folder references:" & nl & "{0}" & nl,
            "Base project references:" & nl & "{0}" & nl,
            "" & nl
           }
        Dim s As Integer = 0
        lin(s) = String.Format(lin(s), ListaCarpetasTraducir.NombrePaT) : s += 1 ' nombre Pat
        lin(s) = String.Format(lin(s), ListaCarpetasTraducir.CorreoGestor) : s += 1 '
        ' necesitamos fecha4
        Dim kk As String = ListaCarpetasTraducir.FechaEntrega.ToString("dd/MMM/yyyy")
        lin(s) = String.Format(lin(s), kk) : s += 1
        lin(s) = String.Format(lin(s), ListaCarpetasTraducir.Notas) : s += 1
        ' las referencias las busco en la clase
        ' 
        Dim auxS2 As String = "" : auxS = ""
        For Each auxS2 In ListaReferencias.Referencias
            auxS &= auxS2 & nl
        Next
        lin(s) = String.Format(lin(s), auxS) : s += 1

        auxS = ""
        '
        ' ahora la lista de carpetas
        ' 
        ' cabecera
        For s = 0 To lin.Count - 1
            sbSalidaMF.Append(lin(s))
        Next
        ' 
        ' ahora tablas
        auxS = ""
        auxS &= "Folder".PadRight(20)
        auxS &= "VendPro".PadRight(15)
        auxS &= "BP".PadRight(10)
        auxS &= "Id".PadRight(5)
        auxS &= "IBMPro".PadRight(15)
        auxS &= "Shipment".PadRight(15)
        auxS &= "Prof".PadRight(10)
        auxS &= "CNT".PadLeft(10)
        auxS &= "PreAna".PadLeft(10)
        auxS &= "IniCal".PadLeft(10)
        auxS &= "Rate".PadLeft(10)
        auxS &= "Total".PadLeft(10)
        sbSalidaMF.Append(auxS & nl) ' añado cabecera
        For Each fila In ListaCarpetasTraducir.tCarpetasTraducir.Rows
            auxS = ""
            auxS &= CType(fila("Carpeta"), String).PadRight(20)
            auxS &= CType(fila("ProMSS"), String).PadRight(15)
            auxS &= CType(fila("ProBase"), String).PadRight(10)
            auxS &= CType(fila("Idioma"), String).PadRight(5)
            auxS &= CType(fila("ProIBM"), String).PadRight(15)
            auxS &= CType(fila("Envio"), String).PadRight(15)
            auxS &= CType(fila("Perfil"), String).PadRight(10)
            If fila("bCNT") Then
                auxS &= CType(fila("CNT"), String).PadLeft(10)
            Else
                auxS &= CType("N/D", String).PadLeft(10)
            End If
            If fila("bPreAna") Then
                auxS &= CType(fila("PreAna"), String).PadLeft(10)
            Else
                auxS &= CType("N/D", String).PadLeft(10)
            End If
            If fila("bIniCal") Then
                auxS &= CType(fila("IniCal"), String).PadLeft(10)
            Else
                auxS &= CType("N/D", String).PadLeft(10)
            End If

            auxS &= CType(String.Format("{0:n3}", fila("Tarifa")), String).PadLeft(10)
            If fila("bTotal") Then
                auxS &= CType(fila("Total"), String).PadLeft(10)
            Else
                auxS &= CType("N/D", String).PadLeft(10)
            End If
            auxS &= CType(String.Format("{0:n2}", fila("Total")), String).PadLeft(10)

            sbSalidaMF.Append(auxS & nl)
        Next
        txtSalida.AppendText(sbSalidaMF.ToString) ' en al consola
        ' no lo añado al zip
        ' ahora voy a crear txt y lo añadiré a mi zipo
        'Dim ArchivoMF As String = ParDirTemporal & NombrePaT.Replace(".PaT", "") & "_MFT.TXT"
        'txtSalida.AppendText("Creo y zipeo " & ArchivoMF & "  ...")
        'Dim sw As System.IO.StreamWriter
        'sw = My.Computer.FileSystem.OpenTextFileWriter(ArchivoMF, True)
        'sw.Write(sbSalidaMF.ToString)
        ' sw.Close()
        ' lo anayado al zip
        'auxS = PreZip & " -j " & ParDirSalidaPaT & NombrePaT & " " & ArchivoMF
        'mandatoZip(auxS, RC)
        ' lo borro
        'File.Delete(ArchivoMF)
        'txtSalida.AppendText("Hecho" & nl)

        Dim ArchivoMFT_HTML As String = Environment.ExpandEnvironmentVariables(Par.ParDirTemporal) & NombrePaT.Replace(".PaT.zip", "") & "_MFT.HTM"
        txtSalida.AppendText("Creating and zipping " & ArchivoMFT_HTML & " ..." & nl)
        If ListaCarpetasTraducir.GeneraMFT_HTML(ArchivoMFT_HTML) = False Then
            txtSalida.AppendText(String.Format(
                "Fatal error. HTML manifest cannot be written: {0}" & nl,
                ArchivoMFT_HTML))
            Exit Sub
        End If
        ' lo anayado al zip
        auxS = PreZip & " -j " & String.Format("""{0}{1}"" ""{2}"" ",
                                     Environment.ExpandEnvironmentVariables(Par.ParDirSalidaPaT),
                                     NombrePaT,
                                     ArchivoMFT_HTML)
        If debug Then logea("zipping MFT HTML->  " & auxS)
        mandatoZip(auxS, RC, , debuga) ' el rc no hace nada


        ' lo borro
        'File.Delete(ArchivoMFT_HTML)
        'txtSalida.AppendText("Done" & nl)

        ' Ahora la parte de la confiración voy a leer el HTML
        Dim ContenidoHTML As String = ""
        Try
            ' En realidad lo importante que necesito está en la definición de la clase
            ' pero hay 2 cosa que no están y que quiero avisar al usuario. 
            ' if If chkMail.IsChecked is false, Email will not be sent
            ContenidoHTML = My.Computer.FileSystem.ReadAllText(ArchivoMFT_HTML)
            Dim SendEmail As Boolean = chkMail.IsChecked
            If ListaCarpetasTraducir.NombreTraductor = "NONE" Then SendEmail = False ' If "NONE" has no sense to send any mail
            Dim VR As New VentanaResumen(ContenidoHTML, SendEmail)
            Dim Resultado As Boolean = VR.ShowDialog()
            If Resultado = False Then
                txtSalida.AppendText("**** User has cancel sent ****" & nl)
                ' voy a intentar eliminar el PaT
                Try
                    Dim PatFullName As String = Environment.ExpandEnvironmentVariables(Par.ParDirSalidaPaT) & NombrePaT
                    If File.Exists(PatFullName) Then
                        File.Delete(PatFullName)
                    End If
                Catch ex As Exception

                End Try

                Exit Sub

            End If
        Catch ex As Exception
            txtSalida.AppendText(String.Format("WARNING. Cannot read {0} to display confirmation dialgo. Process continues.", ArchivoMFT_HTML) & nl)

        End Try




        ' serializo y guardo la contraseña
        Dim ArchivoMFXML As String = Environment.ExpandEnvironmentVariables(Par.ParDirTemporal) & NombrePaT.Replace(".PaT.zip", "") & "_MFT.XML"
        txtSalida.AppendText("Creating and zipping " & ArchivoMFXML & "  ..." & nl)

        Dim auxB As Boolean = False
        auxB = ListaCarpetasTraducir.Serializate(ArchivoMFXML) 'Serializate mejora el codigo no respetaba CRLF
        If auxB = False Then ' error
            txtSalida.AppendText(nl)
            txtSalida.AppendText(String.Format("Fatal error. Error in ListaCarpetasTraducir.Serializate({0})" & nl, ArchivoMFXML))
            Exit Sub
        End If
        'Dim ObjSW As New StreamWriter(ArchivoMFXML) ' lo guardaré en mi ejecutuable
        'Dim x As New XmlSerializer(ListaCarpetasTraducir.GetType) ' serializo mi estructura
        'Try
        'x.Serialize(ObjSW, ListaCarpetasTraducir) ' guardo el par
        'Catch ex As Exception
        'ex = ex.InnerException ' la normal no dice nada
        '' de http://msdn.microsoft.com/en-us/library/aa302290.aspx
        'txtSalida.AppendText(nl)
        'txtSalida.AppendText(String.Format("Message: {0}" & nl, ex.Message))
        ' txtSalida.AppendText(String.Format("Exception Type: {0}" & nl, ex.GetType().FullName))
        ' txtSalida.AppendText(String.Format("Source: {0}" & nl, ex.Source))
        'txtSalida.AppendText(String.Format("StrackTrace: {0}" & nl, ex.StackTrace))
        'txtSalida.AppendText(String.Format("TargetSite: {0}" & nl, ex.TargetSite))
        'Exit Sub  ' no hago nada mas
        'Finally
        ''    ObjSW.Close()
        'End Try
        ' 

        ' lo anayado al zip
        ' lo anayado al zip
        auxS = PreZip & " -j " & String.Format("""{0}{1}"" ""{2}"" ",
                                     Environment.ExpandEnvironmentVariables(Par.ParDirSalidaPaT),
                                     NombrePaT,
                                     ArchivoMFXML)
        If debug Then logea("zipping MFT XML->  " & auxS)
        mandatoZip(auxS, RC, , debuga) ' el rc no hace nada



        ' lo borro
        File.Delete(ArchivoMFXML)
        txtSalida.AppendText("Done" & nl)

        ' ahora voy a por las referencias 
        'Dim archivoReferencia As String
        'For Each archivoReferencia In ListaReferencias.Referencias
        '    auxS = PreZip & " -j " & Par.ParDirSalidaPaT & NombrePaT & " " & archivoReferencia
        '    txtSalida.AppendText(String.Format("Zipeo referencia {0} ... ", archivoReferencia))
        '    mandatoZip(auxS, RC)
        '    txtSalida.AppendText("Hecho" & nl)
        'Next
        'For Each archivoReferencia In ListaReferenciasProyectoBase.ReferenciasProyectoBase
        '    txtSalida.AppendText(String.Format("Zipeo referencia PB {0} ... ", archivoReferencia))
        '    auxS = PreZip & " -j " & Par.ParDirSalidaPaT & NombrePaT & " " & archivoReferencia
        '    mandatoZip(auxS, RC)
        '    txtSalida.AppendText("Hecho" & nl)
        'Next
        ' las referencias las voy a guardar en /reference
        ' por lo que no me queda otro remedio que copiar 
        ' Intento crear directorio referencia

        'Par.ParDirTemporal ' por ejemplo C:\u\tra\tmp\
        Dim dirRef As String = Environment.ExpandEnvironmentVariables(Par.ParDirTemporal) & "Ref\"
        Try
            If Not Directory.Exists(dirRef) Then
                Directory.CreateDirectory(dirRef)
            End If
            Dim DirInfo As New DirectoryInfo(dirRef)
            ' borro lo que haya
            For Each File In DirInfo.GetFiles
                File.Delete()
            Next
            ' ahora copio los archivos de mis referencias
            Dim ArchivoDestino As String = ""

            For Each ArchivoOrigen In ListaReferencias.Referencias  ' los copio
                ArchivoDestino = dirRef & Path.GetFileName(ArchivoOrigen)
                File.Copy(ArchivoOrigen, ArchivoDestino)
            Next

            ' ahora lo añado con path al zip
            auxS = PreZip & " -u " & String.Format("""{0}{1}"" ""{2}"" ",
                                     Environment.ExpandEnvironmentVariables(Par.ParDirSalidaPaT),
                                     NombrePaT,
                                     "Ref\*") ' Para guardar el path relativo DEBO ejecutar zip desde el directorio padre
            If debug Then logea("zipping references ->  " & auxS)
            mandatoZip(auxS, RC, Environment.ExpandEnvironmentVariables(Par.ParDirTemporal), debuga) ' el directorio de trabajo debe ser dir padre de Ref para empaquetar
            'txtSalida.AppendText("Referencias zipeadas")
            AnyadetxtSalida("Zipped" & nl)
        Catch ex As Exception
            MsgBox("We have not been able to zip the references into de path file, otherwise, the pat file is fine.", MsgBoxStyle.Exclamation)
        End Try

        AnyadetxtSalida("PaT file created succesfully in: " & Environment.ExpandEnvironmentVariables(Par.ParDirSalidaPaT) & NombrePaT & nl)
        'txtSalida.AppendText("Archivo PaT creado satisfactoriamente en: " & Par.ParDirSalidaPaT & NombrePaT & nl)



        ' ahora la parte de automatización:
        If NombreTraductor <> "NONE" Then ' tengo que copiar
            Dim TengoErrorAlProcesarDestino As Boolean = False ' si hay algún error
            AnyadetxtSalida(String.Format("Translator name: " & NombreTraductor & " -> "))
            Dim destino As New structDestino
            Dim listadestinos As structListaDestinos
            listadestinos = CargaDestinos()
            ' boy a buscar la info de mi destino
            Dim encontrado As Boolean = False
            For Each destino In listadestinos.ListaDestino
                If destino.DestNombre = NombreTraductor Then
                    encontrado = True
                    Exit For
                End If
            Next
            If encontrado = False Then ' algo mal, no puede ser que tenga destino que no encuentre
                MsgBox("Internal error: Cannot find target info for " & NombreTraductor & ". PaT file is fine, we have not been able retrive target info. Looks like an internal error ", MsgBoxStyle.Exclamation)
                TengoErrorAlProcesarDestino = True
            Else
                ' hay info, ahora lo envio de acuerdo a cada caso.
                Select Case destino.DestTipoDestino
                    Case "PATH"
                        AnyadetxtSalida(String.Format("PATH " & nl))
                        ' tan solo tengo que copiar 
                        Dim source As String = Environment.ExpandEnvironmentVariables(Par.ParDirSalidaPaT) & NombrePaT
                        Dim target As String = destino.DestPath
                        If Not target.EndsWith("\") Then target &= "\"
                        target &= NombrePaT
                        Try

                            AnyadetxtSalida(String.Format("Directory Target: Copying {0} to {1}...", source, target))
                            File.Copy(source, target, True)
                            AnyadetxtSalida("Done!" & nl)
                        Catch ex As Exception
                            ' no hemos podido copy 
                            auxS = nl & String.Format("We have not been able to copy {0} in {1}. Verify target. You can send manually the PaT file {0}", source, target) & nl
                            AnyadetxtSalida(auxS)
                            AnyadetxtSalida("Process continues..." & nl)
                            MsgBox(auxS, MsgBoxStyle.Exclamation)
                            TengoErrorAlProcesarDestino = True
                        Finally

                        End Try
                    Case "SHARED_DRIVE"
                        AnyadetxtSalida(String.Format("SHARED_DRIVE" & nl))

                        Dim source As String = Environment.ExpandEnvironmentVariables(Par.ParDirSalidaPaT) & NombrePaT
                        Dim target As String = destino.DestSharedDriveNombre
                        If Not target.EndsWith("\") Then target &= "\"
                        target &= NombrePaT
                        Try

                            AnyadetxtSalida(String.Format("Target directory: Copying {0} to {1}...", source, target))
                            File.Copy(source, target, True)
                            AnyadetxtSalida("Done!" & nl)
                        Catch ex As Exception
                            ' no hemos podido copy 
                            auxS = nl & String.Format("We have not been able to copy {0} in {1}. Verify target. You can send manually the PaT file {0}", source, target) & nl
                            AnyadetxtSalida(auxS)
                            AnyadetxtSalida("Process continues..." & nl)
                            MsgBox(auxS, MsgBoxStyle.Exclamation)
                            TengoErrorAlProcesarDestino = True

                        Finally

                        End Try
                    Case "FTP", "FTPES"
                        ' no va con archivos grandes
                        ' https://msdn.microsoft.com/en-us/library/vstudio/ms229715%28v=vs.100%29.aspx
                        ' ' con buffer pero no acaba de funcionar 
                        ' http://stackoverflow.com/questions/26653619/error-while-sending-large-file-throug-ftp
                        ' este es el url 
                        ' https://social.msdn.microsoft.com/Forums/en-US/4deb4625-0e06-4b7e-966b-1161c6791069/ftpwebrequest-and-large-files?forum=vblanguage
                        ' Mirar la respuesta de:
                        ''https://social.msdn.microsoft.com/Forums/en-US/246ffc07-1cab-44b5-b529-f1135866ebca/exception-the-underlying-connection-was-closed-the-connection-was-closed-unexpectedly?forum=netfxnetcom en relación a problemas que tiene

                        Dim TIPO_FTP = destino.DestTipoDestino ' FTP O FTPES
                        AnyadetxtSalida(TIPO_FTP & nl)
                        Dim Contrasenya As String = DecryptWithKey(destino.DestFTPContrasenya, Par.ParLongInterno)
                        Contrasenya = Contrasenya.Replace(destino.DestNombre, "")
                        Dim source As String = Environment.ExpandEnvironmentVariables(Par.ParDirSalidaPaT) & NombrePaT
                        Dim hostFTP As String = "ftp://" & destino.DestFTPHost
                        If Not hostFTP.EndsWith("/") Then hostFTP &= "/"

                        Dim target As String = hostFTP & destino.DestFTPNombre
                        If Not target.EndsWith("/") Then target &= "/"
                        target &= NombrePaT & ".TMP"


                        Dim ftprequest As FtpWebRequest = DirectCast(System.Net.WebRequest.Create(target), FtpWebRequest)
                        ftprequest.Credentials = New System.Net.NetworkCredential(destino.DestFTPUsuario, Contrasenya)
                        ftprequest.Method = WebRequestMethods.Ftp.UploadFile
                        ftprequest.EnableSsl = False
                        ftprequest.UsePassive = True
                        ftprequest.UseBinary = True
                        ftprequest.KeepAlive = False
                        ftprequest.ReadWriteTimeout = 300000 '3600000
                        ftprequest.Timeout = 600000 ' 3600000
                        ftprequest.ContentLength = source.Length

                        If TIPO_FTP = "FTPES" Then
                            ftprequest.EnableSsl = True
                            ftprequest.KeepAlive = False
                        End If
                        ServicePointManager.ServerCertificateValidationCallback =
                        Function(se As Object,
                        cert As System.Security.Cryptography.X509Certificates.X509Certificate,
                        chain As System.Security.Cryptography.X509Certificates.X509Chain,
                        sslerror As System.Net.Security.SslPolicyErrors) True
                        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
                        ' end secure



                        Dim file As System.IO.FileInfo = New System.IO.FileInfo(source)
                        Dim estaLongiArchivo As Long = file.Length
                        Dim esta5Porciento As Long = estaLongiArchivo * 0.05
                        Dim estaSalidaInfo As Long = esta5Porciento
                        Dim estaProgreso As Long = 0
                        Dim estaN5Porciento As Long = 0
                        ' he dejado el tamaño de buffer en 2040 quizás resiste más
                        ' esto no lo he mirado http://stackoverflow.com/questions/25015094/uploading-big-file-to-ftp-using-ftpwebrequest-results-in-the-underlying-connect
                        ' aquí en cambio recomienda aumentarlo
                        ' http://stackoverflow.com/questions/10871943/c-sharp-ftp-upload-buffer-size
                        ' y aquí hay otro ejemplo que parece interesante:
                        ' http://stackoverflow.com/questions/16059157/c-sharp-how-to-upload-large-files-to-ftp-site
                        Dim estaTamBuffer As Integer
                        estaTamBuffer = 8192 ' '2048 ' 8192 ' en ejemmplos normalemnte ponen menos, en teoria MTU 8K
                        Dim buffer As Integer = estaTamBuffer
                        Dim content(buffer - 1) As Byte, dataread As Integer

                        Try


                            Using ftpstream As System.IO.FileStream = file.OpenRead()
                                Using request As System.IO.Stream = ftprequest.GetRequestStream
                                    Do
                                        dataread = ftpstream.Read(content, 0, buffer)
                                        request.Write(content, 0, dataread)
                                        estaProgreso += estaTamBuffer ' lo transferido
                                        'AnyadetxtSalida(".")
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
                        Catch ex As Exception
                            auxS = "PaT is fine, but ZZ-Pat cannot read ftp dir. Correct the problem or send PaT manually. " & nl
                            auxS &= ex.Message.ToString & nl
                            auxS &= "source ->  " & source & nl
                            auxS &= "target ->  " & target & nl
                            auxS &= nl
                            MsgBox(auxS, MsgBoxStyle.Exclamation, "Fatal FTP Error." & nl)
                            txtSalida.AppendText(auxS & nl)
                            Exit Sub
                        End Try

                        'Dim response As FtpWebResponse = CType(ftprequest.GetResponse(), FtpWebResponse)
                        'AnyadetxtSalida(String.Format("Upload File Complete, status {0}", response.StatusDescription))
                        AnyadetxtSalida(String.Format("Upload File Complete"))
                        Try
                            ' antes de hacer el rename intento borrarlo si existe.
                            ftprequest = DirectCast(WebRequest.Create(New Uri(target.Replace(".PaT.zip.TMP", ".PaT.zip"))), FtpWebRequest)
                            ftprequest.Credentials = New System.Net.NetworkCredential(destino.DestFTPUsuario, Contrasenya)
                            ftprequest.Method = WebRequestMethods.Ftp.DeleteFile
                            ' secure
                            If TIPO_FTP = "FTPES" Then
                                ftprequest.EnableSsl = True
                                ftprequest.KeepAlive = False
                            End If
                            ' end secure
                            Dim response As FtpWebResponse
                            response = DirectCast(ftprequest.GetResponse(), FtpWebResponse)
                        Catch ex As Exception
                            Dim auxs3 As String = ex.Message
                            ' es normal que de excepción, normalmente el target nunca existirá.
                        End Try

                        Try
                            ' ahora el rename
                            ftprequest = DirectCast(WebRequest.Create(New Uri(target)), FtpWebRequest)
                            ftprequest.Credentials = New System.Net.NetworkCredential(destino.DestFTPUsuario, Contrasenya)
                            auxS = Path.GetFileName(target) : auxS = auxS.Replace(".PaT.zip.TMP", ".Pat.zip")
                            ftprequest.RenameTo = auxS
                            ftprequest.Method = WebRequestMethods.Ftp.Rename
                            ' secure
                            If TIPO_FTP = "FTPES" Then
                                ftprequest.EnableSsl = True
                                ftprequest.KeepAlive = False
                            End If
                            ' end secure

                            Dim response As FtpWebResponse = CType(ftprequest.GetResponse(), FtpWebResponse)

                            response = DirectCast(ftprequest.GetResponse(), FtpWebResponse)


                        Catch ex As Exception
                            auxS = "PaT is fine, but ZZ-Pat cannot rename in ftp dir. Correct the problem or send PaT manually. " & nl
                            auxS &= ex.Message.ToString & nl
                            auxS &= "Rename targeget ->  " & auxS & nl
                            auxS &= nl
                            MsgBox(auxS, MsgBoxStyle.Exclamation, "Fatal FTP Error." & nl)
                            txtSalida.AppendText(auxS & nl)
                            Exit Sub

                        End Try


                    Case Else
                        auxS = "Internal error. Target has not target type PATH, SHARED_DRIVE or FTP"
                        AnyadetxtSalida(auxS)
                        MsgBox(auxS, MsgBoxStyle.Exclamation)
                        AnyadetxtSalida("Process continues..." & nl)
                        TengoErrorAlProcesarDestino = True

                End Select

                If TengoErrorAlProcesarDestino = False And Trim(destino.DestCorreo) <> "" _
                    And chkMail.IsChecked Then
                    ' ahora envío el correo si no hay error
                    Dim destinatario As String = destino.DestCorreo
                    AnyadetxtSalida(String.Format("Sending note to {0}... ", destino.DestCorreo & nl))
                    Dim mailssl As Boolean = Par.ParMailSSL ' True
                    Dim mailpuerto As Integer = Par.ParMailPuerto ' 587
                    Dim mailcontrase As String = DecryptWithKey(Par.ParMailContrase, Par.ParLongInterno) ' "la de manolo..."
                    Dim mailhost As String = Par.ParMailHost ' "smtp.gmail.com"
                    Dim mailusuario As String = Par.ParMailUsuario ' "xxxxx@gmail.com"
                    Dim client As New SmtpClient
                    client.DeliveryMethod = SmtpDeliveryMethod.Network
                    client.Host = mailhost
                    client.Port = mailpuerto
                    client.EnableSsl = mailssl
                    ' autenticaciÃ³n
                    Dim credentials As New System.Net.NetworkCredential(mailusuario, mailcontrase)
                    client.UseDefaultCredentials = False
                    client.Credentials = credentials

                    Dim msg As New MailMessage
                    msg.From = New MailAddress(mailusuario)
                    msg.To.Add(New MailAddress(destinatario))
                    Dim RtL As New List(Of MailAddress)
                    msg.ReplyTo = New MailAddress("miguelcanals@hotmail.com")
                    msg.Subject = String.Format("PaT name: {0} / Recipient: {1}", Path.GetFileName(ArchivoMFT_HTML), destino.DestNombre)
                    msg.IsBodyHtml = "true"

                    Select Case destino.DestTipoDestino
                        Case "PATH"
                            auxS = String.Format("Target directory: {0}", destino.DestPath)
                        Case "SHARED_DRIVE"
                            auxS = String.Format("Target directory: {0}", destino.DestSharedDriveNombre)
                        Case "FTP", "FTPES"
                            auxS = String.Format("{0}: {1}<br>User: {2}<br>Target: {2}", destino.DestTipoDestino, destino.DestFTPHost, destino.DestFTPUsuario, destino.DestFTPNombre)
                    End Select

                    Dim lin2() As String = {
                    "<HTML><BODY>",
                    "PaT location information",
                    "Pat recipient: " & destino.DestNombre,
                    "Target type:" & destino.DestTipoDestino,
                    auxS}
                    For s = 0 To lin2.Count - 1
                        msg.Body &= lin2(s) & "<br>"
                    Next
                    msg.Body &= "</BODY></HTML>"

                    msg.Body &= File.ReadAllText(ArchivoMFT_HTML)
                    ' log estÃ¡ en 
                    'Dim ArchivoLog As String = CargaCte("mantenimiento", "log")
                    Dim datos As New Attachment(ArchivoMFT_HTML)
                    msg.Attachments.Add(datos)

                    Try
                        client.Send(msg)
                        AnyadetxtSalida("Done!" & nl)
                    Catch ex As Exception
                        AnyadetxtSalida(String.Format("Error sending note: {0}", ex.Message) & nl & "Proces continues..." & NombreTraductor & nl)
                    End Try
                    msg.Dispose()

                End If
            End If

        End If
        ' ahora el correo
        txtSalida.AppendText("Some clean up..." & nl)
        txtSalida.AppendText("Udating log..." & nl)
        logea(String.Format("{0:yyyy-MM-dd HH:mm:ss} {1} {2}", Now,
                            ListaCarpetasTraducir.CorreoGestor,
                            ListaCarpetasTraducir.NombrePaT))
        File.Delete(ArchivoMFT_HTML)
        txtSalida.AppendText("Done" & nl)


        'If Par.ParRecopilarUso Then
        'Dim EnvioUDP As New strucEnvioUDP(Par.ParRecopilarUso, Par.ParHostUDP, Par.ParPuertoUDP, Par.ParProveedor, txtNombrePaT.Text)
        'EnvioUDP.envia()
        ' If EnvioUDP.TodoOK Then
        ' txtSalida.AppendText("Stat info OK" & nl)
        ' Else
        ' txtSalida.AppendText("Stat info NOK" & nl)
        ' End If
        ' End If
        AnyadetxtSalida("************************  END  ************************ " & nl)

    End Sub

    Private Sub btnParametros_Click(sender As Object, e As RoutedEventArgs) Handles btnParametros.Click
        Dim VP As New VentanaParametros
        VP.ShowDialog()

        ' debo cargar los parámetros
        Dim Par As structParametros
        Par = CargaParametros()
        'lblDirPaT.Text = Environment.ExpandEnvironmentVariables(Par.ParDirSalidaPaT)
        txtPatOutput.Text = Environment.ExpandEnvironmentVariables(Par.ParDirSalidaPaT)
        txtTarifa.Text = Par.ParTarifaPredeterminada
    End Sub


    Private Sub Button_Click_1(sender As Object, e As RoutedEventArgs)
        Dim VP As New ProcesaManifiesto
        VP.ShowDialog()


    End Sub


    Private Sub txtBorrarCarpetasTraducir_Click(sender As Object, e As RoutedEventArgs) Handles txtBorrarCarpetasTraducir.Click
        ' voy a borrar si hay alguna seleccionada

        Dim kk As Integer = dgCarpetas.SelectedItems.Count
        If kk = 0 Then ' se debe seleccionar al menos una
            MsgBox("At least you must select a folder", MsgBoxStyle.Information)
        End If
        Dim auxS As String = ""
        Dim listaABorrar As New List(Of String)
        For Each row In dgCarpetas.SelectedItems
            listaABorrar.Add(row("Carpeta"))
        Next
        ' una vez la lista borro uno a uno
        For Each auxS In listaABorrar
            ListaCarpetasTraducir.BorraCarpeta(auxS)
        Next
        ' una vez borrado, vuelvo a vincular
        dgCarpetas.ItemsSource = ListaCarpetasTraducir.tCarpetasTraducir.DefaultView
        dgCarpetas.IsReadOnly = True
        '
        ' es una putada, pero al elminar, puede cambiar el nombre 
        ' 
        ' Necesito los parámetros
        Dim Par As structParametros
        Par = CargaParametros()
        Dim PBCorregidoOLD As String = "NOSEPB" ' txtProyectoBase.Text
        Dim PBCorregidoNEW As String = "NOSEPB" 'txtProyectoBase.Text ' en principio supongo el mismo


        If Par.ParMasDeUnPB Then ' si dejo multiproyectos, puede haber cambiado
            Dim ListaParaNombrePaT As New List(Of String)
            For Each row In ListaCarpetasTraducir.tCarpetasTraducir.Rows
                Dim PBaux As String = row("ProBase")
                ' miro que no esté en la lista
                Dim Encontrado As Boolean = False
                For Each elelista In ListaParaNombrePaT
                    If elelista = PBaux Then Encontrado = True : Exit For ' no busco más ya lo tengo
                Next
                If Encontrado = False Then ListaParaNombrePaT.Add(PBaux)
            Next
            ' al llegar aquí tengo una lista de PB de todas las carpetas
            ListaParaNombrePaT.Sort()
            ' ahora creo el nombre
            ' si solo queda uno ya no es MB
            PBCorregidoNEW = ""
            If ListaParaNombrePaT.Count > 1 Then
                PBCorregidoNEW = "MB_"
            End If
            For Each elelista In ListaParaNombrePaT
                PBCorregidoNEW &= elelista & "_"
            Next
            PBCorregidoNEW = PBCorregidoNEW.TrimEnd("_")
        End If
        ' pueden pasar 3 cosas que no quede nada, restablezco
        If ListaCarpetasTraducir.tCarpetasTraducir.Rows.Count = 0 Then
            ' borror
            txtNombrePaT.Text = ""

        Else ' quedan carpetas
            If PBCorregidoNEW <> PBCorregidoOLD Then ' ha cambiado
                txtNombrePaT.Text = PBCorregidoNEW & "_" & DateTime.Now.ToString("yyyyMMdd") ' cambia el PaT

            End If

        End If
    End Sub

    Private Sub txtBorrarTodasCarpetasTraducir_Click(sender As Object, e As RoutedEventArgs) Handles txtBorrarTodasCarpetasTraducir.Click
        ' voy a borrar si hay alguna seleccionada

        Dim kk As Integer = dgCarpetas.SelectedItems.Count
        Dim auxS As String = ""
        Dim listaABorrar As New List(Of String)
        For Each row In dgCarpetas.Items
            ListaCarpetasTraducir.BorraCarpeta(row("Carpeta"))
        Next
        ' una vez borrado, vuelvo a vincular
        dgCarpetas.ItemsSource = ListaCarpetasTraducir.tCarpetasTraducir.DefaultView
        dgCarpetas.IsReadOnly = True
        '
        ' es una putada, pero al elminar, puede cambiar el nombre 
        ' 
        ' Necesito los parámetros
        Dim Par As structParametros
        Par = CargaParametros()
        Dim PBCorregidoOLD As String = "NOSEPB" ' txtProyectoBase.Text
        Dim PBCorregidoNEW As String = "NOSEPB" 'txtProyectoBase.Text ' en principio supongo el mismo


        ' pueden pasar 3 cosas que no quede nada, restablezco
        If ListaCarpetasTraducir.tCarpetasTraducir.Rows.Count = 0 Then
            ' borror
            txtNombrePaT.Text = ""
        End If

    End Sub


    Private Sub btnTest_Click(sender As Object, e As RoutedEventArgs) Handles btnTest.Click
        'Dim Par As structParametros
        'Par = CargaParametros()
        'If Par.ParRecopilarUso Then
        '    Dim EnvioUDP As New strucEnvioUDP(Par.ParRecopilarUso, Par.ParHostUDP, Par.ParPuertoUDP, Par.ParProveedor, txtNombrePaT.Text)
        '    EnvioUDP.envia()
        '    If EnvioUDP.TodoOK Then
        '        txtSalida.AppendText("UPD info OK" & nl)
        '    Else
        '        txtSalida.AppendText("UPD info NOK" & nl)
        '    End If
        'End If

    End Sub


    Function AnyadetxtSalida(s As String)
        txtSalida.AppendText(s)
        txtSalida.SelectionStart = txtSalida.Text.Length
        txtSalida.ScrollToEnd()
        txtSalida.Refresh() '' creo no sirve de nada
        Return 0
    End Function


    Private Sub btnAyuda_Click(sender As Object, e As RoutedEventArgs) Handles btnAyuda.Click
        System.Diagnostics.Process.Start("http://www.mknals.com/02_ZZPaT_Overview.html")
    End Sub


    Private Sub btnDestinos_Click(sender As Object, e As RoutedEventArgs) Handles btnDestinos.Click
        ' este link interesante para pasar información de child a parent
        ' http://www.codeproject.com/Tips/548160/Transferring-information-between-two-forms-Part
        ' pero esto lo saqué de otro sitio 
        ' http://stackoverflow.com/questions/280579/how-do-i-pass-a-value-from-a-child-back-to-the-parent-form


        Dim VD As New VentanaDestinos(cbxDestinoPaT.SelectedItem)
        VD.ShowDialog()
        Dim NombreDestino As String = VD.ObtenDestino
        Dim destino As New structDestino
        Dim listadestinos As structListaDestinos
        listadestinos = CargaDestinos()
        ' me pueden haber añadido o borrado alguno, así que tengo que recrear
        cbxDestinoPaT.Items.Clear()
        cbxDestinoPaT.Items.Add("NONE")
        For Each destino In listadestinos.ListaDestino
            cbxDestinoPaT.Items.Add(destino.DestNombre)
        Next
        Dim encontrado As Boolean = False
        Dim posi As Integer = 1 ' pq ya tenemos a NONE
        For Each destino In listadestinos.ListaDestino
            If destino.DestNombre = NombreDestino Then encontrado = True : cbxDestinoPaT.SelectedIndex = posi : Exit For
            posi += 1
        Next
        If encontrado = False Then Exit Sub ' no puedo hacer nada, debería haber uno
        ' voy a cargar la variable txtInfoDestino
        Dim auxS As String = "Type:{0} "
        auxS = String.Format(auxS, destino.DestTipoDestino)
        ' resto de información segun el tipo de destino
        With destino
            Select Case .DestTipoDestino
                Case "PATH"
                    auxS &= String.Format(" Target: {0} ", .DestPath)
                Case "SHARED_DRIVE"
                    auxS &= String.Format(" Target:  {0} ({1})", .DestSharedDriveNombre, .DestSharedDriveUsuario)
                Case "FTP", "FTPES"
                    auxS &= String.Format(" {0}:  {1} ({2} / {3})", .DestTipoDestino, .DestFTPNombre, .DestFTPHost, .DestSharedDriveUsuario)
                Case Else
            End Select
        End With

        txtInfoDestino.Text = auxS

        ' debo cargar los parámetros
        'Dim Par As structParametros
        'Par = CargaParametros()
        'lblDirPaT.Text = Par.ParDirSalidaPaT
        'txtTarifa.Text = Par.ParTarifaPredeterminada
    End Sub

    Private Sub cbxDestinoPaT_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles cbxDestinoPaT.SelectionChanged
        ' cambia el cbxDestinoPat
        If cbxDestinoPaT.SelectedIndex < 0 Then Exit Sub ' nada que hacer

        Dim NuevoDestino As String = cbxDestinoPaT.SelectedItem.ToString
        If NuevoDestino = "NONE" Then txtInfoDestino.Text = "" : Exit Sub

        ' ahora cargo los destinos
        Dim listadestinos As structListaDestinos
        listadestinos = CargaDestinos()
        For Each destino In listadestinos.ListaDestino
            If NuevoDestino = destino.DestNombre Then
                Dim auxS As String = "Type:{0} "
                auxS = String.Format(auxS, destino.DestTipoDestino)
                ' resto de información segun el tipo de destino
                With destino
                    Select Case .DestTipoDestino
                        Case "PATH"
                            auxS &= String.Format(" Target: {0} ", .DestPath)
                        Case "SHARED_DRIVE"
                            auxS &= String.Format(" Target:  {0} ({1})", .DestSharedDriveNombre, .DestSharedDriveUsuario)
                        Case "FTP", "FTPES"
                            auxS &= String.Format(" {0}:  {1} ({2} / {3})", .DestTipoDestino, .DestFTPNombre, .DestFTPHost, .DestSharedDriveUsuario)
                        Case Else
                    End Select
                End With

                txtInfoDestino.Text = auxS
            End If

        Next


        ' 
    End Sub

    Function ObtenDestinoenMain() As String
        Return cbxDestinoPaT.SelectedItem
    End Function

    Structure structContajeCNT
        Dim todoOk As Boolean
        Dim MsgTitle As String
        Dim MsgCorpus As String
        Dim archivo As String
        Dim carpeta As String
        Dim salida As String
        Dim contaje As Long
        Public Sub New(archivo, carpeta)
            Me.archivo = archivo : Me.carpeta = carpeta
            todoOk = True
            salida = "" : MsgTitle = "" : MsgCorpus = ""
        End Sub
        Public Sub calcula()
            Dim opcion, archivosalida, mandato, auxS As String
            Dim RC As Integer
            opcion = "TARGET"
            Dim debug As Boolean : debug = False

            'archivosalida = UCase(PathPrefijo & carpeta & ".CNT")
            archivosalida = archivo
            salida &= "CNT -> " & vbCrLf
            mandato = "/TAsk=WORDCNT /FLD={0} /OUT={1} /OV=YES /OP={2} /QUIET=NOMSG"
            mandato = String.Format(mandato, carpeta, archivosalida, opcion)
            ejecuta(mandato, RC)
            ' Shell(mandato, AppWinStyle.MinimizedFocus, True, -1)
            If RC <> 0 Then
                Dim debuginfo As String = "" : If debug Then debuginfo = vbCrLf & mandato
                auxS = "CNT report cannot be done. Otmbarch return code RC=" & RC.ToString & debuginfo & ". Process continues."
                salida &= auxS & vbCrLf
                MsgCorpus = auxS : MsgTitle = "CNT report cannot be done"
                ' MsgBox(auxS, MsgBoxStyle.Information, "CNT report cannot be done")
                todoOk = False
            Else
                If debug Then ' enseño salida si hay debug
                    Dim myProcess As New Process()
                    myProcess.StartInfo.FileName = "notepad.exe"
                    myProcess.StartInfo.Arguments = archivosalida
                    ' " " & displayArray(masterIndex).rom & " -volume -" & My.Settings.mame_volume & " -skip_gameinfo"
                    myProcess.Start()
                    myProcess.WaitForExit()
                    myProcess.Close()
                End If
                ' tengo que leer la última fila del cnt
                ' voy a saco
                Dim lines As String() = IO.File.ReadAllLines(archivosalida)
                Dim ultimalinea = lines(lines.Length - 1)
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
                ' actualizo 0 'Total' 1 Translated 2 Untraslated
                contaje = Int(ValColSal(2))
                salida &= contaje.ToString() & vbCrLf

            End If
            ' si he lllado aquí rc =0
            ' voy a guardarlo en el zip

        End Sub
    End Structure
    Structure structContajeCalculating
        Dim todoOk As Boolean
        Dim MsgTitle As String
        Dim MsgCorpus As String
        Dim archivo As String
        Dim carpeta As String
        Dim perfil As String
        Dim salida As String
        Dim contaje As Long
        Public Sub New(archivo, carpeta, perfil)
            Me.archivo = archivo : Me.carpeta = carpeta : Me.perfil = perfil
            todoOk = True
            salida = "" : MsgTitle = "" : MsgCorpus = ""
        End Sub
        Public Sub calcula()
            Dim archivosalida, mandato, auxS As String
            Dim RC As Integer
            Dim debug As Boolean : debug = False

            'archivosalida = UCase(PathPrefijo & carpeta & ".CNT")
            archivosalida = archivo
            salida &= "CALCULATING -> " & vbCrLf
            mandato = "/TAsk=CNTRPT /FLD={0} /OUT={1} /RE=CALCULATING /TYPE=BASE_SUMMARY_FACT /PROFILE={2} /OV=YES /QUIET=NOMSG"
            mandato = String.Format(mandato, carpeta, archivosalida, perfil)
            ejecuta(mandato, RC)
            ' Shell(mandato, AppWinStyle.MinimizedFocus, True, -1)
            If RC <> 0 Then
                Dim debuginfo As String = "" : If debug Then debuginfo = vbCrLf & mandato
                auxS = "Cannot create calculating report. Otmbarch return code RC=" & RC.ToString & debuginfo & ". Process continues."
                salida &= auxS & vbCrLf
                MsgCorpus = auxS : MsgTitle = "Calculating report cannot be done"
                ' MsgBox(auxS, MsgBoxStyle.Information, "CNT report cannot be done")
                todoOk = False
            Else
                If debug Then ' enseño salida si hay debug
                    Dim myProcess As New Process()
                    myProcess.StartInfo.FileName = "notepad.exe"
                    myProcess.StartInfo.Arguments = archivosalida
                    ' " " & displayArray(masterIndex).rom & " -volume -" & My.Settings.mame_volume & " -skip_gameinfo"
                    myProcess.Start()
                    myProcess.WaitForExit()
                    myProcess.Close()
                End If
                ' tengo que leer la última fila del cnt
                ' voy a saco

                Dim lines As String() = IO.File.ReadAllLines(archivosalida)
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

                ' Payable 0 Words 1 : 2 mn 2
                contaje = Int(Val(ValColSal(3)))
                'txtSalida.AppendText(fila("IniCal") & nl)
                salida &= contaje.ToString() & vbCrLf


            End If
            ' si he lllado aquí rc =0
            ' voy a guardarlo en el zip

        End Sub
    End Structure
    Structure structContajePreanalisis
        Dim todoOk As Boolean
        Dim MsgTitle As String
        Dim MsgCorpus As String
        Dim archivo As String
        Dim carpeta As String
        Dim perfil As String
        Dim salida As String
        Dim contaje As Long
        Public Sub New(archivo, carpeta, perfil)
            Me.archivo = archivo : Me.carpeta = carpeta : Me.perfil = perfil
            todoOk = True
            salida = "" : MsgTitle = "" : MsgCorpus = ""
        End Sub
        Public Sub calcula()
            Dim archivosalida, mandato, auxS As String
            Dim RC As Integer
            Dim debug As Boolean : debug = False

            'archivosalida = UCase(PathPrefijo & carpeta & ".CNT")
            archivosalida = archivo
            salida &= "PREANALISIS -> " & vbCrLf
            mandato = "/TAsk=CNTRPT /FLD={0} /OUT={1} /RE=PREANALYSIS /TYPE=BASE_SUMMARY_FACT /PROFILE={2} /OV=YES /QUIET=NOMSG"
            mandato = String.Format(mandato, carpeta, archivosalida, perfil)
            ejecuta(mandato, RC)
            ' Shell(mandato, AppWinStyle.MinimizedFocus, True, -1)
            If RC <> 0 Then
                Dim debuginfo As String = "" : If debug Then debuginfo = vbCrLf & mandato
                auxS = "Cannot create preanalysis report. Otmbarch return code RC=" & RC.ToString & debuginfo & ". Process continues."
                salida &= auxS & vbCrLf
                MsgCorpus = auxS : MsgTitle = "Preanalysis report cannot be done"
                ' MsgBox(auxS, MsgBoxStyle.Information, "CNT report cannot be done")
                todoOk = False
            Else
                If debug Then ' enseño salida si hay debug
                    Dim myProcess As New Process()
                    myProcess.StartInfo.FileName = "notepad.exe"
                    myProcess.StartInfo.Arguments = archivosalida
                    ' " " & displayArray(masterIndex).rom & " -volume -" & My.Settings.mame_volume & " -skip_gameinfo"
                    myProcess.Start()
                    myProcess.WaitForExit()
                    myProcess.Close()
                End If
                ' tengo que leer la última fila del cnt
                ' voy a saco

                Dim lines As String() = IO.File.ReadAllLines(archivosalida)
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
                contaje = Int(Val(ValColSal(3)))
                'txtSalida.AppendText(fila("IniCal") & nl)
                salida &= contaje.ToString & vbCrLf
            End If
            ' si he lllado aquí rc =0
            ' voy a guardarlo en el zip

        End Sub
    End Structure

    Private Sub btnObtenInfo_Click(sender As Object, e As RoutedEventArgs) Handles btnObtenInfo.Click
        ' inicio algún contr

        AnyadetxtSalida("")

        Dim auxS As String = ""
        Dim debuga As Boolean = True
        Dim boolCrea As Boolean = chkCreateFiles.IsChecked
        Dim dirCrea As String = ""
        auxS = Trim(txtSufijo.Text)
        Dim sufijo As String = ""
        Dim c As Char
        For Each c In auxS
            If c = "0" Or Val(c) > 0 Then
                sufijo = sufijo & c.ToString
            ElseIf UCase(c) > "A" And UCase(c) < "Z" Then
                sufijo = sufijo & c.ToString
            End If
        Next



        If boolCrea Then
            dirCrea = Trim(txtPatOutput.Text)
            If Not Directory.Exists(dirCrea) Then
                auxS = "Looks like directory {0} does not exist. " & nl
                auxS &= "Create one or fix the entry or do not select the Create report file option. "
                MsgBox(auxS, MsgBoxStyle.Critical, "Error directory")
                Exit Sub
            Else
                dirCrea = CompruebaAntibarra(dirCrea)
            End If
        End If
        '

        ' antes de hacer nada, miro las condiciones
        ' 1) Necesito nombre
        Dim Par As structParametros
        Par = CargaParametros()
        Dim carpeta As String = ""
        Dim perfil As String = ""
        Dim PathPrefijo As String = Environment.ExpandEnvironmentVariables(Par.ParDirTemporal)
        If boolCrea Then PathPrefijo = dirCrea
        Dim fila As DataRow
        Dim idioma As String = ""
        Dim ProIBM As String = ""
        Dim ProMSS As String = "N/A"
        Dim TengoProIBM As Boolean = False

        For Each fila In ListaCarpetasTraducir.tCarpetasTraducir.Rows

            carpeta = UCase(fila("Carpeta"))
            perfil = fila("Perfil")
            auxS = String.Format("Folder {0} Profile {1}" & nl, carpeta, perfil)
            txtSalida.AppendText(auxS)
            Dim mandato As String = ""
            Dim archivosalida = ""
            Dim opcion As String = ""

            If fila("bCNT") Then
                archivosalida = PathPrefijo & carpeta & ".CNT"
                Dim ContajeCT As New structContajeCNT(archivosalida, carpeta)
                ContajeCT.calcula()
                If ContajeCT.todoOk Then
                    fila("CNT") = ContajeCT.contaje
                    If Not boolCrea Then File.Delete(archivosalida)

                Else
                    MsgBox(ContajeCT.MsgCorpus, MsgBoxStyle.Information, ContajeCT.MsgTitle)
                End If
                txtSalida.AppendText(ContajeCT.salida)
            End If

            archivosalida = PathPrefijo & String.Format("{0}_{1}_{2}_cal.rpt", carpeta, perfil, sufijo)
            Dim ContajeCalculating As New structContajeCalculating(archivosalida, carpeta, perfil)
            ' el CNT detallado es opcion = "TMMATCH"
            If fila("bIniCal") Then
                ContajeCalculating.calcula()
                If ContajeCalculating.todoOk Then
                    fila("IniCal") = ContajeCalculating.contaje
                    If Not boolCrea Then File.Delete(archivosalida)
                Else
                    MsgBox(ContajeCalculating.MsgCorpus, MsgBoxStyle.Information, ContajeCalculating.MsgTitle)
                End If
                txtSalida.AppendText(ContajeCalculating.salida)
            End If


            ' ahora voy con los calculating 
            ' el preanalisis
            archivosalida = PathPrefijo & String.Format("{0}_{1}_preana.rpt", carpeta, perfil)
            Dim ContajePreAnalisis As New structContajePreanalisis(archivosalida, carpeta, perfil)
            If fila("bPreAna") Then
                ContajePreAnalisis.calcula()
                If ContajePreAnalisis.todoOk Then
                    fila("PreAna") = ContajePreAnalisis.contaje
                    If Not boolCrea Then File.Delete(archivosalida)
                Else
                    MsgBox(ContajePreAnalisis.MsgCorpus, MsgBoxStyle.Information, ContajePreAnalisis.MsgTitle)
                End If
                txtSalida.AppendText(ContajePreAnalisis.salida)
            End If


        Next
        ListaCarpetasTraducir.tCarpetasTraducir.AcceptChanges()
        dgCarpetas.ItemsSource = ListaCarpetasTraducir.tCarpetasTraducir.DefaultView
        dgCarpetas.IsReadOnly = True


        ' voy a por el informe
        AnyadetxtSalida("************************  END  ************************ " & nl)
    End Sub


End Class




' Ojo con esta clase, se basa copiando la que hay en la última verisón de z-pat
' es como la que hay en z-pat, PERO sin menos procedimientos, por eso no está en el
'´módulo común
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



' Ojo con esta clase, se basa copiando la que hay en la última verisón de z-pat
' es como la que hay en z-pat, PERO sin menos procedimientos, por eso no está en el
'´módulo común
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
    Function GeneraMFT_HTML(archivo As String) As Boolean
        Dim Par As structParametros
        Par = CargaParametros()

        ' Esto es adicional para complementar la información de la confirmación
        Dim NombreTraductor As String = Me.NombreTraductor
        Dim CorreoTraductor As String = "N/A"
        Dim TipoDest As String = "N/A"
        Dim Target As String = "N/A"
        Dim Source As String = Me.DirectorioGestor
        If Not Source.EndsWith("\") Then Source &= "\"
        Source &= NombrePaT
        If NombreTraductor <> "NONE" Then ' tengo que copiar
            Dim destino As New structDestino
            Dim listadestinos As structListaDestinos
            listadestinos = CargaDestinos()
            ' boy a buscar la info de mi destino
            Dim encontrado As Boolean = False
            For Each destino In listadestinos.ListaDestino
                If destino.DestNombre = NombreTraductor Then
                    encontrado = True
                    Exit For
                End If
            Next
            If encontrado = False Then ' algo mal, no puede ser que tenga destino que no encuentre
                ' es un error interno no se bien qué hacer
            Else
                CorreoTraductor = destino.DestCorreo
                TipoDest = destino.DestTipoDestino
                ' hay info, ahora lo envio de acuerdo a cada caso.
                Select Case TipoDest
                    Case "PATH"
                        Target = destino.DestPath
                        If Not target.EndsWith("\") Then target &= "\"
                        Target &= NombrePaT
                    Case "SHARED_DRIVE"
                        Target = destino.DestSharedDriveNombre
                        If Not Target.EndsWith("\") Then Target &= "\"
                        Target &= NombrePaT
                    Case "FTP", "FTPES"
                        Dim TIPO_FTP = destino.DestTipoDestino ' FTP O FTPES
                        Dim hostFTP As String = "ftp://" & destino.DestFTPHost
                        If Not hostFTP.EndsWith("/") Then hostFTP &= "/"

                        Target = hostFTP & destino.DestFTPNombre
                        If Not Target.EndsWith("/") Then Target &= "/"
                        Target &= NombrePaT & ".TMP"

                    Case Else
                        ' Other internal error
                        ' auxS = "Internal error. Target has not target type PATH, SHARED_DRIVE or FTP"
                        ' Exit Sub

                End Select
            End If
        End If
        ' Al llegar tengo NombreTraductor CorreoTraductor TipoDest Target 


        ' Esto es lo relamente importante
        Dim todoOK As Boolean = True ' optimismo
        Dim sbSalidaMF As New StringBuilder(String.Format("<HTML><HEAD><TITLE>{0}</TITLE></HEAD><BODY>", Me.NombrePaT))
        Dim nl As String = "<br>" & Environment.NewLine
        Dim lin() As String = {
            "<H2><b>PaT name:</b> {0}</H2>" & nl,
            "<b>Project Manager:</b> {0}" & nl,
            "<b>Source PaT:</b> {0}" & nl,
            "<b>Target PaT:</b> ({0}) {1}" & nl & nl,
            "<b>Translator:</b> {0} ({1})" & nl & nl,
            "<b>Due date:</b> {0}" & nl,
            "" & nl
           }
        Dim s As Integer = 0
        lin(s) = String.Format(lin(s), Me.NombrePaT) : s += 1 ' Nombre de pat
        lin(s) = String.Format(lin(s), Me.CorreoGestor) : s += 1 ' Nombre de pat
        ' necesitamos fecha        
        lin(s) = String.Format(lin(s), Source) : s += 1 ' Directorio gestor
        lin(s) = String.Format(lin(s), TipoDest, Target) : s += 1 ' (TipoDest) Target
        If NombreTraductor = "NONE" Then
            lin(s) = String.Format(lin(s), "PaT will not be sent, only created in source directory", "") : s += 1 ' Nombre traductor (correo)
        Else
            lin(s) = String.Format(lin(s), Me.NombreTraductor, CorreoTraductor) : s += 1 ' Nombre traductor (correo)
        End If

        Dim kk As String = Me.FechaEntrega.ToString("dd/MMM/yyyy")
        lin(s) = String.Format(lin(s), kk) : s += 1 ' Due date

        ' Las referencias y las notas, lo dejamos para el final
        Dim lin_2() As String = {
                    nl & "<b>Folder references:</b>" & nl & "{0}" & nl,
                             nl & "<b>Notes:</b>" & nl & " {0}" & nl,
            "" & nl
           }
        Dim s_2 As Integer = 0
        ' las referencias las busco en la clase
        Dim auxS2 As String = "" : Dim auxS As String = ""
        For Each auxS2 In ListaReferencias.Referencias
            auxS &= auxS2 & nl
        Next
        lin_2(s_2) = String.Format(lin_2(s_2), auxS) : s_2 += 1 ' Reference list
        lin_2(s_2) = String.Format(lin_2(s_2), Me.Notas) : s_2 += 1 ' Notes


        '
        ' ahora la lista de carpetas
        ' 
        ' cabecera

        For s = 0 To lin.Count - 1
            sbSalidaMF.Append(lin(s))
        Next
        Dim IniTabla As String = ""
        IniTabla = "<style type='text/css' media='screen'> td { padding-left: 4px; padding-right: 4px;  padding-top: 0px; padding-bottom: 0px; } </style>"
        IniTabla &= "<table border='1' cellspacing='0' >"
        Dim IniFila As String = "<tr>"
        Dim FinFila As String = "</tr>"
        Dim FinTabla As String = "</table>"
        sbSalidaMF.Append(IniTabla) ' inicio tabla
        auxS = ""
        auxS &= "<td><b>Folder</b></td>"
        auxS &= String.Format("<td><b>{0}</b></td>", Par.ParProveedor)
        auxS &= "<td><b>BP</b></td>"
        auxS &= "<td><b>Id</b></td>"
        auxS &= "<td><b>IBMPro</b></td>"
        auxS &= "<td><b>Shipment</b></td>"
        auxS &= "<td><b>Profile</b></td>"
        auxS &= "<td><b>CNT</b></td>"
        auxS &= "<td><b>PreAna</b></td>"
        auxS &= "<td><b>IniCal</b></td>"
        auxS &= "<td><b>Rate</b></td>"
        sbSalidaMF.Append(IniFila & auxS & FinFila) ' añado cabecera
        For Each fila In Me.tCarpetasTraducir.Rows
            auxS = ""
            auxS &= String.Format("<td><b>{0}</b></td>", fila("Carpeta"))
            auxS &= String.Format("<td>{0}</td>", fila("ProMSS"))
            auxS &= String.Format("<td>{0}</td>", fila("ProBase"))
            auxS &= String.Format("<td>{0}</td>", fila("Idioma"))
            auxS &= String.Format("<td>{0}</td>", fila("ProIBM"))
            auxS &= String.Format("<td>{0}</td>", fila("Envio"))
            auxS &= String.Format("<td><b>{0}</b></td>", fila("Perfil"))
            If fila("bCNT") Then
                auxS &= String.Format("<td><b>{0}</b></td>", fila("CNT"))
            Else
                auxS &= String.Format("<td>{0}</td>", "N/D")
            End If
            If fila("bPreAna") Then
                auxS &= String.Format("<td>{0}</td>", fila("PreAna"))
            Else
                auxS &= String.Format("<td>{0}</td>", "N/D")
            End If
            If fila("bIniCal") Then
                auxS &= String.Format("<td>{0}</td>", fila("IniCal"))
            Else
                auxS &= String.Format("<td>{0}</td>", "N/D")
            End If

            auxS &= String.Format("<td>{0}</td>", String.Format("{0:n3}", fila("Tarifa")))
            sbSalidaMF.Append(IniFila & auxS & FinFila)
        Next
        sbSalidaMF.Append(FinTabla)

        ' ahora la parte final
        For s_2 = 0 To lin_2.Count - 1
            sbSalidaMF.Append(lin_2(s_2))
        Next
        sbSalidaMF.Append("</BODY></HTML>")


        Dim sw As System.IO.StreamWriter
        sw = My.Computer.FileSystem.OpenTextFileWriter(archivo, False) ' False es no append
        Try
            sw.Write(sbSalidaMF.ToString)
        Catch ex As Exception
            todoOK = False
        Finally
            sw.Close()
        End Try



        Return todoOK
    End Function
    Function GeneraOC_HTML(archivo As String, TipoI As String) As Boolean
        ' TipoI si es "FAC" es para generar la factura
        ' Cuando lo borryo

        Dim suma As Single = 0
        Dim IVA As Single = 0
        Dim IRPF As Single = 0
        Dim Total As Single = 0
        ' los parámetros
        Dim Par As structParametros
        Par = CargaParametros()


        Select Case TipoI
            Case "FAC"
            Case Else ' intento borrar
                If File.Exists(archivo) Then
                    File.Delete(archivo)
                End If
        End Select

        Dim todoOK As Boolean = True ' optimismo
        Dim sbSalidaMF As New StringBuilder(String.Format("<HTML><HEAD><TITLE>{0}</TITLE></HEAD><BODY>", Me.NombrePaT))
        Dim nl As String = "<br>" & Environment.NewLine
        If TipoI <> "FAC" Then
            Dim lin() As String = {
                "<H2><b>PaT name:</b> {0}</H2>" & nl,
                "" & nl
               }
            Dim s As Integer = 0
            lin(s) = String.Format(lin(s), Me.NombrePaT) : s += 1
            ' 
            ' cabecera
            For s = 0 To lin.Count - 1
                sbSalidaMF.Append(lin(s))
            Next
        Else
            sbSalidaMF.Append(nl)
            sbSalidaMF.Append(nl)
            sbSalidaMF.Append("<b>Economic/billing information:</b>" & nl)

        End If

        Dim auxS As String = "" : Dim auxS2 = ""
        Dim IniTabla As String = ""
        IniTabla = "<style type='text/css' media='screen'> td { padding-left: 4px; padding-right: 4px;  padding-top: 0px; padding-bottom: 0px; } </style>"
        IniTabla &= "<table border='1' cellspacing='0' >"
        Dim IniFila As String = "<tr>"
        Dim FinFila As String = "</tr>"
        Dim FinTabla As String = "</table>"
        sbSalidaMF.Append(IniTabla) ' inicio tabla
        If TipoI <> "FAC" Then auxS &= "<td><b>Folder</b></td>"
        auxS = String.Format("<td><b>{0}</b></td>", Par.ParProveedor)
        auxS &= "<td><b>BP</b></td>"
        auxS &= "<td><b>Id</b></td>"
        auxS &= "<td><b>IBMPro</b></td>"
        auxS &= "<td><b>Shipment</b></td>"
        If TipoI <> "FAC" Then auxS &= "<td><b>Profile</b></td>"
        If TipoI <> "FAC" Then auxS &= "<td><b>CNT</b></td>"
        If TipoI <> "FAC" Then auxS &= "<td><b>PreAna</b></td>"
        If TipoI <> "FAC" Then auxS &= "<td><b>IniCal</b></td>"
        If TipoI <> "FAC" Then auxS &= "<td><b>FinCal</b></td>"
        auxS &= "<td><b>WCT</b></td>"
        auxS &= "<td><b>Rate</b></td>"
        auxS &= "<td><b>Total</b></td>"
        sbSalidaMF.Append(IniFila & auxS & FinFila) ' añado cabecera
        For Each fila In Me.tCarpetasTraducir.Rows
            If TipoI <> "FAC" Then auxS &= String.Format("<td>{0}</td>", fila("Carpeta"))
            auxS = String.Format("<td>{0}</td>", fila("ProMSS"))
            auxS &= String.Format("<td>{0}</td>", fila("ProBase"))
            auxS &= String.Format("<td>{0}</td>", fila("Idioma"))
            auxS &= String.Format("<td>{0}</td>", fila("ProIBM"))
            auxS &= String.Format("<td>{0}</td>", fila("Envio"))
            If TipoI <> "FAC" Then auxS &= String.Format("<td>{0}</td>", fila("Perfil"))
            If TipoI <> "FAC" Then
                If fila("bCNT") Then
                    auxS &= String.Format("<td>{0}</td>", fila("CNT"))
                Else
                    auxS &= String.Format("<td>{0}</td>", "N/D")
                End If
                If fila("bPreAna") Then '
                    auxS &= String.Format("<td>{0}</td>", fila("PreAna"))
                Else
                    auxS &= String.Format("<td>{0}</td>", "N/D")
                End If
                If fila("bIniCal") Then
                    auxS &= String.Format("<td>{0}</td>", fila("IniCal"))
                Else
                    auxS &= String.Format("<td>{0}</td>", "N/D")
                End If
                If fila("bFinCal") Then
                    auxS &= String.Format("<td>{0}</td>", fila("FinCal"))
                Else
                    auxS &= String.Format("<td>{0}</td>", "N/D")
                End If
            End If
            auxS &= String.Format("<td><b>{0}</b></td>", fila("Contaje"))
            auxS &= String.Format("<td>{0}</td>", String.Format("{0:n3}", fila("Tarifa")))
            If fila("bTotal") Then
                auxS &= String.Format("<td>{0}</td>", fila("Total"))
                suma += fila("Total")
            Else
                auxS &= String.Format("<td>{0}</td>", "N/D")
            End If
            If TipoI = "FAC" And fila("bTotal") = False Then
                '
                ' no escribo si es factruua y no tengo inof facturación
            Else
                sbSalidaMF.Append(IniFila & auxS & FinFila)
            End If

        Next
        If TipoI = "FAC" Then
            ' fila de suma ( ya está redondeada)
            IVA = suma * Par.ParIVA / 100 : IVA = Math.Round(IVA, 2, MidpointRounding.ToEven)
            IRPF = -suma * Par.ParIRPF / 100 : IRPF = Math.Round(IRPF, 2, MidpointRounding.ToEven)
            Total = suma + IVA + IRPF ' ya estará redondeado
            sbSalidaMF.Append(IniFila & "<td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td>" & FinFila)
            sbSalidaMF.Append(IniFila & String.Format("<td></td><td></td><td></td><td></td><td></td><td></td><td>Sum</td><td>{0}</td>" & FinFila, suma))
            sbSalidaMF.Append(IniFila & String.Format("<td></td><td></td><td></td><td></td><td></td><td></td><td>Iva {1}%</td><td>{0}</td>" & FinFila, IVA, Par.ParIVA))
            sbSalidaMF.Append(IniFila & String.Format("<td></td><td></td><td></td><td></td><td></td><td></td><td>IRPF {1}%</td><td>{0}</td>" & FinFila, IRPF, Par.ParIRPF))
            sbSalidaMF.Append(IniFila & String.Format("<td></td><td></td><td></td><td></td><td></td><td></td><td><b>Total</b></td><td>{0}</td>" & FinFila, Total))

        End If
        sbSalidaMF.Append(FinTabla)
        sbSalidaMF.Append("</BODY></HTML>")
        Dim boolappend As Boolean = False ' en principio sin append
        If TipoI = "FAC" Then boolappend = True
        Dim sw As System.IO.StreamWriter
        sw = My.Computer.FileSystem.OpenTextFileWriter(archivo, boolappend) ' False es no append
        Try
            sw.Write(sbSalidaMF.ToString)
        Catch ex As Exception
            todoOK = False
        Finally
            sw.Close()
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