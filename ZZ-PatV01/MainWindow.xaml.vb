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




Class MainWindow
    Dim debug As Boolean = False
    Dim nl As String = Environment.NewLine
    Dim ListaCarpetasTraducir As New ClaseListaCarpetasTraducir
    Dim ListaReferenciasProyectoBase As New ClaseListaReferenciasProyectoBase
    Dim ListaReferencias As New ClaseListaReferencias

    
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
        VentanaPrincipal.Title = "ZZ-Pat " & lastMod & "/" & lastMod2 & "  (" & ViaEjecutable & ")"
        ' voy a buscar el path de OTM
        Dim ViaSistema = Environment.GetEnvironmentVariable("PATH")
        'voy a buscar C:\OTM\WIN
        Dim auxi As Integer = InStr(ViaSistema, ":\OTM\WIN")
        ' 
        btnTest.Visibility = Windows.Visibility.Hidden

        ' solo visible para hacer pruebas

        If auxi = 0 Then ' OTM parece no instalado 
            ' para que el splashscreen no se me coma el diálogo
            Dim temp As Window = New Window()
            temp.Visibility = Windows.Visibility.Hidden
            temp.Show()
            MessageBox.Show(temp, "Cannot find :\OTM\WIN in PATH variable. Looks like OTM is no installed.", "OTM not in path", MessageBoxButton.OK, MessageBoxImage.Error)
            Application.Current.Shutdown()
            Exit Sub
        End If

        Dim DiscoOTM As String = ViaSistema.Substring(auxi - 2, 1)
        Dim ListaCarpetas As New ClaseListaCarpetas(String.Format("{0}:\OTM\PROPERTY", DiscoOTM))
        ListaCarpetas.Carpetas.Sort()
        CargaListaCarpetasDDW(ListaCarpetas.Carpetas)
        Dim ListaPerfiles As New ClaseListaPerfiles(String.Format("{0}:\OTM\PROPERTY", DiscoOTM))
        ListaPerfiles.Perfiles.Sort()
        CargaListaPerfilesCBX(ListaPerfiles.Perfiles)

        ' en princpio CNT
        chkCNT.IsChecked = True ' si cnt
        chkIniCal.IsChecked = False ' no calculatin inical
        ' valores
        
        '    Application.Current.Properties("MyApplicationScopeProperty") = "myApplicationScopePropertyValue"
        ' cargo los parámetros para visualizar el directorio de salida
        Dim Par As structParametros
        Par = CargaParametros()

        lblDirPaT.Text = Par.ParDirSalidaPaT
        txtTarifa.Text = Par.ParTarifaPredeterminada

        ' 
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
        Dim ProyectoBase As String = txtProyectoBase.Text
        Dim Perfil As String = cbxPerfiles.SelectedItem ' perfil
        ' primero voy a ver que no está duplicada
        Dim row As DataRow
        For Each row In ListaCarpetasTraducir.tCarpetasTraducir.Rows
            If carpeta = row("Carpeta") Then ' quiere decir que ya la tengo no la puedo añadir
                MsgBox(String.Format("Folder {0} is already in the folder to be translated. ", carpeta), MsgBoxStyle.Exclamation)
                Exit Sub
            End If
        Next
        ' Carpetas que ya tengo en la lista
        Dim eleListaCarpeta As Integer = ListaCarpetasTraducir.tCarpetasTraducir.Rows.Count
        Dim PBCarpetaSeleccionada As String = ""
        'txbCarpeta.Text = carpeta
        Dim PBCorregido As String
        If Len(carpeta) >= 5 Then
            PBCorregido = carpeta.Substring(0, 5) : PBCarpetaSeleccionada = PBCorregido ' en principio iguales
            If Trim(ProyectoBase) <> "" Then ' tengo uno
                If PBCorregido <> ProyectoBase Then ' solo dejo crear un PaT del mismo proyecto
                    ' vale no coincide, así que la única posibilidad es que no haya
                    ' carpetas
                    If eleListaCarpeta = 0 Then ' voy a borrar la info base
                        'ListaReferenciasProyectoBase = New ClaseListaReferenciasProyectoBase
                    Else ' hay un PB y la auxS no coincide
                        ListaReferenciasProyectoBase.TengoProyectoBase = False ' ya no tengo
                        If Par.ParMasDeUnPB Then ' ojo que cambios en este if afectan al borrar la fila del dg
                            Dim ListaParaNombrePaT As New List(Of String)
                            ListaParaNombrePaT.Add(PBCarpetaSeleccionada) ' el que voy a añadir seguro
                            ' busco en la lista
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
                            PBCorregido = "MB_"
                            For Each elelista In ListaParaNombrePaT
                                PBCorregido &= elelista & "_"
                            Next
                            PBCorregido = PBCorregido.TrimEnd("_")
                        Else  ' no permito
                            MsgBox("A PaT file can have folders only from base project. ", MsgBoxStyle.Exclamation, "Folder does not belong to the same base project")
                            Exit Sub
                        End If
                    End If
                End If
            End If
            ProyectoBase = PBCarpetaSeleccionada
            txtProyectoBase.Text = PBCorregido ' nombre corregido
            txtNombrePaT.Text = PBCorregido & "_" & DateTime.Now.ToString("yyyyMMdd")
            'DateTime.Now.ToString("yyyyMMddHHmmss")
        Else
            MsgBox("Only folders with a 5 characteres word lenght can be processed.")
            Exit Sub
        End If
        Dim tarifa As Single = CType(txtTarifa.Text, Single) ' respeta el encoding
        ' primero añadiré la carpeta
        ' ahor añado la carpeta

        txtNotas.Text = Trim(txtNotas.Text) ' importante xml si no me da problemas
        ListaCarpetasTraducir.AnyadeCarpeta(carpeta, ProyectoBase, Perfil, _
                                            chkCNT.IsChecked, 0, chkIniCal.IsChecked, _
                                            0, "N/D", txtEnvio.Text, tarifa)
        ' ahora acutalizao el DG

        'dgCarpetas.Items.Clear()
        dgCarpetas.ItemsSource = ListaCarpetasTraducir.tCarpetasTraducir.DefaultView
        dgCarpetas.IsReadOnly = True



        ' tengo que ver si he cargado o no un proyeto
        If ListaReferenciasProyectoBase.TengoProyectoBase Then
            ' ojo que ya tengo uno
            If ListaReferenciasProyectoBase.ProyectoBase = ProyectoBase Then ' 
                ' no tengo que hacer nada más salgo continuo normal
                Exit Sub
            Else
                ' no me sirve lo que tengo
                ListaReferenciasProyectoBase = New ClaseListaReferenciasProyectoBase ' 
            End If
        End If

        '
        Dim archivo As String = Par.ParDirTPot & PBCorregido & "_RefBase.XML"
        If ListaReferenciasProyectoBase.TengoProyectoBase = False Then ' lo voy a cargar
            ListaReferenciasProyectoBase.CargarProyecto(archivo, PBCorregido) 'puede ser multibase
        End If

        ' borro la lista
        lbxReferenciasProyectoBase.Items.Clear() ' borro las referencias
        Dim referencia As String
        For Each referencia In ListaReferenciasProyectoBase.ReferenciasProyectoBase
            lbxReferenciasProyectoBase.Items.Add(referencia)
        Next




    End Sub
    '
    ' Eventos drag-drop
    '
    Private Sub lbxReferenciasProyectoBase_DragEnter(sender As Object, e As DragEventArgs) Handles lbxReferenciasProyectoBase.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop, False) Then
            e.Effects = DragDropEffects.All
        Else
            e.Effects = DragDropEffects.None
        End If

    End Sub

    Private Sub lbxReferenciasProyectoBase_Drop(sender As Object, e As DragEventArgs) Handles lbxReferenciasProyectoBase.Drop
        Dim files As String()
        files = e.Data.GetData(DataFormats.FileDrop, False)

        For Each filename In files
            lbxReferenciasProyectoBase.Items.Add(filename)
            ListaReferenciasProyectoBase.anyade(filename)
        Next
    End Sub

    Private Sub btnBorrarReferenciasProyectoBase_Click(sender As Object, e As RoutedEventArgs) Handles btnBorrarReferenciasProyectoBase.Click
        lbxReferenciasProyectoBase.Items.Clear()
        ListaReferenciasProyectoBase = New ClaseListaReferenciasProyectoBase()

    End Sub


    Private Sub btnGuardarReferencias_Click(sender As Object, e As RoutedEventArgs) Handles btnGuardarReferencias.Click
        txtSalida.AppendText("-> Saving references")
        txtSalida.AppendText(nl)
        Dim Par As structParametros
        Par = CargaParametros()
        Dim PB As String = txtProyectoBase.Text
        Dim archivoRefBase As String = Par.ParDirTPot & PB & "_RefBase.xml"
        If IO.File.Exists(archivoRefBase) = True Then ' lo borro
            File.Delete(archivoRefBase)
        End If
        ' miro si hay alguna referencia
        If lbxReferenciasProyectoBase.Items.Count = 0 Then
            txtSalida.AppendText("ERROR: There are not references" & nl)
            Exit Sub ' salgo
        End If

        ' ahora lo voy a guardar como xML
        ' la forma más fácil, serializar
        Dim RefPB As structReferenciasProyectoBase
        RefPB.ProyectoBase = PB
        ' ahora la lista
        RefPB.Referencias = New List(Of String)
        Dim auxS As String = "" : Dim i As Integer
        For i = 0 To lbxReferenciasProyectoBase.Items.Count - 1
            auxS = lbxReferenciasProyectoBase.Items(i).ToString
            RefPB.Referencias.Add(auxs)
        Next
        ' ya tengo my estructura

        Dim ObjSW As New StreamWriter(archivoRefBase) ' lo guardaré en mi ejecutuable
        Dim x As New XmlSerializer(RefPB.GetType) ' serializo mi estructura
        x.Serialize(ObjSW, RefPB) ' guardo el par
        ObjSW.Close()
        txtSalida.AppendText(String.Format("-> Base project references file saved {0}", PB))
        txtSalida.AppendText(nl)
        ' ahora debería cargar mi instancia por si se utiliza
        ListaReferenciasProyectoBase.CargarProyecto(archivoRefBase, PB)

    End Sub
   

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




    Private Sub btnWCT_Click(sender As Object, e As RoutedEventArgs) Handles btnWCT.Click
        ' inicio algún contr

        AnyadetxtSalida("")

        Dim debuga As Boolean = True

        ' antes de hacer nada, miro las condiciones
        ' 1) Necesito nombre
        Dim auxS As String = ""
        Dim PreZip As String = System.Reflection.Assembly.GetExecutingAssembly().Location
        PreZip = Path.GetDirectoryName(PreZip) & "\zip.exe "
        ' ahora me interesa la parte ejecutable 
        Dim bCrearPaT As Boolean = True
        Dim NombrePaT As String = txtNombrePaT.Text
        Dim RC As Integer = 0
        ' ParDirSalidaPaT  es el directorio de trabajo
        Dim Par As structParametros
        Par = CargaParametros()
        If bCrearPaT Then
            If NombrePaT = "" Then
                'txtSalida.AppendText("Si se va a crear un PaT, debe haber un nombre " & nl)
                auxS = "If are going to create a PaT file, you must specify a filename."
                MsgBox(auxS, MsgBoxStyle.Exclamation, "No PaT filename")
                Exit Sub
                txtSalida.AppendText(auxS & nl)
            Else ' si existe lo borra
                NombrePaT &= ".PaT.zip"
                ' borro el paquete si existe
                If File.Exists(Par.ParDirSalidaPaT & NombrePaT) Then
                    File.Delete(Par.ParDirSalidaPaT & NombrePaT)
                End If
            End If
        End If
        txtSalida.AppendText("PaT filename: " & Par.ParDirSalidaPaT & NombrePaT & nl)

        ' para actualizar la IU
        ' http://stackoverflow.com/questions/5192169/update-gui-using-backgroundworker
        ' 
        If ListaCarpetasTraducir.nCarpetas <= 0 Then
            txtSalida.AppendText("OPPS -> No hay carpetas" & nl)
            auxS = "There are not any folders to be translated. Select one o more folders."
            MsgBox(auxS, MsgBoxStyle.Exclamation, "Empty folders list")
            txtSalida.AppendText(auxS & nl)
            Exit Sub
        End If
        txtSalida.AppendText("Running folder calculating reports " & nl)
        Dim carpeta As String = ""
        Dim perfil As String = ""
        Dim PathPrefijo As String = Par.ParDirTemporal
        Dim fila As DataRow
        Dim idioma As String = ""
        Dim ProIBM As String = ""
        Dim ProMSS As String = "N/A"
        Dim TengoProIBM As Boolean = False
        Dim archivoExcel = Par.ParArchivoExcel
        For Each fila In ListaCarpetasTraducir.tCarpetasTraducir.Rows
            ProIBM = fila("Carpeta") : TengoProIBM = False
            ' las carpetas son _SPA o _CAT '
            idioma = "N/D"
            ' el idioma supondré que hay _ y luego hasta el final.
            Dim posiGuion As Integer = 0
            For i As Integer = Len(ProIBM) - 1 To 0 Step -1
                If ProIBM.Substring(i, 1) = "_" Then
                    posiGuion = i
                    Exit For
                End If
            Next
            If posiGuion <> Len(ProIBM) - 1 Then
                idioma = ProIBM.Substring(posiGuion + 1, Len(ProIBM) - 1 - posiGuion)
                auxS = ProIBM.Substring(posiGuion, Len(ProIBM) - posiGuion)
                ProIBM = Replace(ProIBM, auxS, "")
                TengoProIBM = True
            End If
            If idioma = "SPA" Then idioma = "ES"
            If idioma = "CAT" Then idioma = "CA"

            fila("Idioma") = idioma
            ProMSS = "N/D"

            If TengoProIBM Then
                ProMSS = LeeProyectoMS(ProIBM, idioma, archivoExcel, Par.ParSQL)
                fila("ProIBM") = ProIBM

            Else
                fila("ProIBM") = "N/A"
            End If
            fila("ProMSS") = ProMSS
        Next
        ListaCarpetasTraducir.tCarpetasTraducir.AcceptChanges()


        For Each fila In ListaCarpetasTraducir.tCarpetasTraducir.Rows

            carpeta = UCase(fila("Carpeta"))
            perfil = fila("Perfil")
            auxS = String.Format("Folder {0} Profile {1}" & nl, carpeta, perfil)
            txtSalida.AppendText(auxS)
            Dim mandato As String = ""
            Dim archivosalida = ""
            Dim opcion As String = ""
            ' si tengo que crear el paquete exportaré la carpeta y haré el zip
            If bCrearPaT Then
                ' el directorio de salida es ParDirTemporal p.e. C:\kk\kk2
                Dim unidad As String = Path.GetPathRoot(Par.ParDirTemporal).Substring(0, 1) ' la letra
                auxS = Path.GetPathRoot(Par.ParDirTemporal) ' C:\
                auxS = auxS.Replace("\", "") ' C: que es lo que debo quitar
                Dim DirTemporalSinUnidad As String = Par.ParDirTemporal.Replace(auxS, "")
                mandato = " /TAsk=FLDEXP /FLD={0} /TOdrive={1} /ToPath={2} /OPtions=(MEM,ROMEM,DOCMEM) /OVerwrite=YES /QUIET=NOMSG  "
                mandato = String.Format(mandato, carpeta, unidad, DirTemporalSinUnidad) ' exportaré al directorio temporal
                ejecuta(mandato, RC)
                If RC <> 0 Then
                    auxS = "Folder " & carpeta & " cannot be exported in directory " & Par.ParDirTemporal & ". " & _
                           "RC=" & RC.ToString
                    MsgBox(auxS, MsgBoxStyle.Exclamation, "Folder cannot be exported" & nl)
                    txtSalida.AppendText(auxS & nl)
                    Exit Sub
                End If
                'txtSalida.AppendText("Carpeta " & carpeta & " exporada en " & Par.ParDirTemporal & nl)
                AnyadetxtSalida("Folder " & carpeta & " exported in " & Par.ParDirTemporal)
                ' ahora el zip
                auxS = PreZip & " -j " & Par.ParDirSalidaPaT & NombrePaT & " " & Par.ParDirTemporal & carpeta & ".FXP"
                mandatoZip(auxS, RC)
                ' lo borro
                File.Delete(Par.ParDirTemporal & carpeta & ".FXP")
            End If

            ' el CNT detallado es opcion = "TMMATCH"
            If fila("bCNT") Then
                opcion = "TARGET"
                'archivosalida = UCase(PathPrefijo & carpeta & ".CNT")
                archivosalida = PathPrefijo & carpeta & ".CNT"
                txtSalida.AppendText("CNT -> ")
                mandato = "/TAsk=WORDCNT /FLD={0} /OUT={1} /OV=YES /OP={2} /QUIET=NOMSG"
                mandato = String.Format(mandato, carpeta, archivosalida, opcion)
                ejecuta(mandato, RC)
                ' Shell(mandato, AppWinStyle.MinimizedFocus, True, -1)
                If RC <> 0 Then
                    Dim debuginfo As String = "" : If debug Then debuginfo = vbCrLf & mandato
                    auxS = "CNT report cannot be done. RC=" & RC.ToString & debuginfo & ". Process continues."
                    txtSalida.AppendText(auxS & nl)
                    MsgBox(auxS, MsgBoxStyle.Information, "CNT report cannot be done")
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
                End If
                ' si he lllado aquí rc =0
                ' voy a guardarlo en el zip
                If bCrearPaT Then
                    ' zip kkk.zip arhivo (añade el archivo al zip) D es para que no ponga el path
                    auxS = PreZip & " -j " & Par.ParDirSalidaPaT & NombrePaT & " " & archivosalida
                    mandatoZip(auxS, RC) ' el rc no hace nada
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
                fila("CNT") = Int(ValColSal(2))
                AnyadetxtSalida(fila("CNT") & nl)
                File.Delete(archivosalida)
            End If
            ' ahora voy con los calculating 
            If fila("bIniCal") Then
                archivosalida = PathPrefijo & String.Format("{0}_{1}_ini_cal.rpt", carpeta, perfil)
                txtSalida.AppendText("CAL_INI -> ")
                mandato = "/TAsk=CNTRPT /FLD={0} /OUT={1} /RE=CALCULATING /TYPE=BASE_SUMMARY_FACT /PROFILE={2} /OV=YES /QUIET=NOMSG"
                mandato = String.Format(mandato, carpeta, archivosalida, perfil)
                ejecuta(mandato, RC)
                ' Shell(mandato, AppWinStyle.MinimizedFocus, True, -1)
                If RC <> 0 Then
                    Dim debuginfo As String = "" : If debug Then debuginfo = vbCrLf & mandato
                    auxS = "Cannot create calculating report. RC=" & RC.ToString & debuginfo
                    MsgBox(auxS)
                    'txtSalida.AppendText(auxS & nl)
                    AnyadetxtSalida(auxS & nl)
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
                End If
                ' si he lllado aquí rc =0
                ' voy a guardarlo en el zip
                If bCrearPaT Then
                    ' zip kkk.zip arhivo (añade el archivo al zip) D es para que no ponga el path
                    auxS = PreZip & " -j " & Par.ParDirSalidaPaT & NombrePaT & " " & archivosalida
                    mandatoZip(auxS, RC) ' el rc no hace nada

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
                fila("IniCal") = Int(Val(ValColSal(3)))
                'txtSalida.AppendText(fila("IniCal") & nl)
                AnyadetxtSalida(fila("IniCal") & nl)
                File.Delete(archivosalida)
            End If
            ' acutalizo fila
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
            .ListaReferenciasProyectoBase = ListaReferenciasProyectoBase
            .FechaCreacionPaT = Now
            .FirmaInstancia = .ObtenHash(Par.ParSecretoServidor)

        End With
        ' El informe de TXT y la consola

        ' voy a por el informe
        Dim sbSalidaMF As New StringBuilder("")
        Dim lin() As String = { _
            "PaT name: {0}" & nl, _
            "Due date: {0}" & nl, _
            "Notes:" & nl & " {0}" & nl, _
            "Folder references:" & nl & "{0}" & nl, _
            "Base project references:" & nl & "{0}" & nl, _
            "" & nl
           }
        Dim s As Integer = 0
        lin(s) = String.Format(lin(s), ListaCarpetasTraducir.NombrePaT) : s += 1
        ' necesitamos fecha

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
        For Each auxS2 In ListaReferenciasProyectoBase.ReferenciasProyectoBase
            auxS &= auxS2 & nl
        Next
        lin(s) = String.Format(lin(s), auxS) : s += 1
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

        Dim ArchivoMFT_HTML As String = Par.ParDirTemporal & NombrePaT.Replace(".PaT.zip", "") & "_MFT.HTM"
        txtSalida.AppendText("Creating and zipping " & ArchivoMFT_HTML & "  ...")
        If ListaCarpetasTraducir.GeneraMFT_HTML(ArchivoMFT_HTML) = False Then
            txtSalida.AppendText(String.Format( _
                "Fatal error. HTML manifest cannot be written: {0}" & nl, _
                ArchivoMFT_HTML))
            Exit Sub
        End If
        ' lo anayado al zip
        auxS = PreZip & " -j " & Par.ParDirSalidaPaT & NombrePaT & " " & ArchivoMFT_HTML
        mandatoZip(auxS, RC)
        ' lo borro
        File.Delete(ArchivoMFT_HTML)
        txtSalida.AppendText("Done" & nl)




        ' serializo y guardo la contraseña
        Dim ArchivoMFXML As String = Par.ParDirTemporal & NombrePaT.Replace(".PaT.zip", "") & "_MFT.XML"
        txtSalida.AppendText("Creating and zipping " & ArchivoMFXML & "  ...")

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
        auxS = PreZip & " -j " & Par.ParDirSalidaPaT & NombrePaT & " " & ArchivoMFXML
        mandatoZip(auxS, RC)
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
        Dim dirRef As String = Par.ParDirTemporal & "Ref\"
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
            For Each ArchivoOrigen In ListaReferenciasProyectoBase.ReferenciasProyectoBase ' los copio
                ArchivoDestino = dirRef & Path.GetFileName(ArchivoOrigen)
                File.Copy(ArchivoOrigen, ArchivoDestino)

            Next
            For Each ArchivoOrigen In ListaReferencias.Referencias  ' los copio
                ArchivoDestino = dirRef & Path.GetFileName(ArchivoOrigen)
                File.Copy(ArchivoOrigen, ArchivoDestino)
            Next

            ' ahora lo añado con path al zip
            auxS = PreZip & " -u " & Par.ParDirSalidaPaT & NombrePaT & " " & "Ref\*"
            'txtSalida.AppendText("Zipeando referencias ...")
            AnyadetxtSalida("Zipping references...")
            mandatoZip(auxS, RC, Par.ParDirTemporal)
            'txtSalida.AppendText("Referencias zipeadas")
            AnyadetxtSalida("Zipped" & nl)
        Catch ex As Exception
            MsgBox("We have not been able to zip the references into de path file, otherwise, the pat file is fine.", MsgBoxStyle.Exclamation)
        End Try

        AnyadetxtSalida("PaT file created succesfully in: " & Par.ParDirSalidaPaT & NombrePaT & nl)
        'txtSalida.AppendText("Archivo PaT creado satisfactoriamente en: " & Par.ParDirSalidaPaT & NombrePaT & nl)

        If Par.ParRecopilarUso Then
            Dim EnvioUDP As New strucEnvioUDP(Par.ParRecopilarUso, Par.ParHostUDP, Par.ParPuertoUDP, Par.ParProveedor, txtNombrePaT.Text)
            EnvioUDP.envia()
            If EnvioUDP.TodoOK Then
                txtSalida.AppendText("Stat info OK" & nl)
            Else
                txtSalida.AppendText("Stat info NOK" & nl)
            End If
        End If
        AnyadetxtSalida("************************  END  ************************ " & nl)

    End Sub

    Private Sub btnParametros_Click(sender As Object, e As RoutedEventArgs) Handles btnParametros.Click
        Dim VP As New VentanaParametros
        VP.ShowDialog()
        ' debo cargar los parámetros
        Dim Par As structParametros
        Par = CargaParametros()
        lblDirPaT.Text = Par.ParDirSalidaPaT
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
        Dim PBCorregidoOLD As String = txtProyectoBase.Text
        Dim PBCorregidoNEW As String = txtProyectoBase.Text ' en principio supongo el mismo


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
            txtProyectoBase.Text = ""
            txtNombrePaT.Text = ""
            ListaReferenciasProyectoBase = New ClaseListaReferenciasProyectoBase ' 
            lbxReferenciasProyectoBase.Items.Clear() ' brro
        Else ' quedan carpetas
            If PBCorregidoNEW <> PBCorregidoOLD Then ' ha cambiado
                txtProyectoBase.Text = PBCorregidoNEW ' cambia el proyecto base
                txtNombrePaT.Text = PBCorregidoNEW & "_" & DateTime.Now.ToString("yyyyMMdd") ' cambia el PaT
                ' cambian las referencias
                ListaReferenciasProyectoBase = New ClaseListaReferenciasProyectoBase ' 
                Dim archivo As String = Par.ParDirTPot & PBCorregidoNEW & "_RefBase.XML"
                lbxReferenciasProyectoBase.Items.Clear() ' borro las referencias
                ListaReferenciasProyectoBase.CargarProyecto(archivo, PBCorregidoNEW) 'puede ser multibase
                Dim referencia As String
                For Each referencia In ListaReferenciasProyectoBase.ReferenciasProyectoBase
                    lbxReferenciasProyectoBase.Items.Add(referencia)
                Next
                ' otra cosa que tengo que cambiar es el nombre del archivo
            End If

        End If
        

    End Sub


 
    Private Sub btnTest_Click(sender As Object, e As RoutedEventArgs) Handles btnTest.Click
        Dim Par As structParametros
        Par = CargaParametros()
        If Par.ParRecopilarUso Then
            Dim EnvioUDP As New strucEnvioUDP(Par.ParRecopilarUso, Par.ParHostUDP, Par.ParPuertoUDP, Par.ParProveedor, txtNombrePaT.Text)
            EnvioUDP.envia()
            If EnvioUDP.TodoOK Then
                txtSalida.AppendText("UPD info OK" & nl)
            Else
                txtSalida.AppendText("UPD info NOK" & nl)
            End If
        End If

    End Sub

   
    Function AnyadetxtSalida(s As String)
        txtSalida.AppendText(s)
        txtSalida.SelectionStart = txtSalida.Text.Length
        txtSalida.ScrollToEnd()
        txtSalida.Refresh()
        Return 0
    End Function

    
    Private Sub btnAyuda_Click(sender As Object, e As RoutedEventArgs) Handles btnAyuda.Click
        System.Diagnostics.Process.Start("http://www.mknals.com/010_1_TMT_ZZ-Pat.html")
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
    Public Sub CargarProyecto(archivo As String, PB As String)
        TengoProyectoBase = True ' algo tengo
        ArchivoConPB = archivo
        ProyectoBase = PB
        Dim RefPB As structReferenciasProyectoBase
        RefPB.ProyectoBase = ProyectoBase
        RefPB.Referencias = New List(Of String)
        If IO.File.Exists(ArchivoConPB) = True Then
            ' no existe no puedo hacer nada salgo todoOK = False
            Try
                Dim objStreamReader As New StreamReader(ArchivoConPB)
                Dim x As New XmlSerializer(RefPB.GetType)
                RefPB = x.Deserialize(objStreamReader)
                objStreamReader.Close()
            Catch ex As Exception
                ' no puedo hacer nada, el archivo está mal
            End Try
        End If
        ' tenga lo que tenga acutlizao los valores de la clae
        todoOK = True
        TengoProyectoBase = True
        ReferenciasProyectoBase = RefPB.Referencias

    End Sub
End Class

Class ClaseListaCarpetas
    Public todoOK As Boolean
    Public Carpetas As List(Of String)
    Public Sub New(Oricarpetas)
        todoOK = False ' pesimimo
        'Dim Carpetas As New List(Of String) ' donde guardaré las carpetas
        Carpetas = New List(Of String)
        Dim dDir As New DirectoryInfo(Oricarpetas)
        Dim archivo As FileSystemInfo

        For Each archivo In dDir.GetFiles("*.f00")
            Dim aux As String = ""
            Dim carpetaLong As String = ""
            ' 
            Dim input As New FileStream(archivo.FullName, FileMode.Open, FileAccess.Read)
            Dim reader As New BinaryReader(input)
            ' por si acaso leeo hasta punto

            Dim ncorto As String = "" : Dim posi As Integer = "0"
            Dim b As Byte
            While reader.PeekChar <> Asc(".")
                b = reader.ReadByte()
                ncorto &= Chr(b)
                posi += 1
            End While
            '
            ' busco el string 
            reader.ReadBytes(Convert.ToInt32(&HA79) - posi)

            While reader.PeekChar <> 0
                b = reader.ReadByte()
                carpetaLong &= Chr(b)
            End While
            reader.Close()
            If Len(carpetaLong) = 0 Then
                aux = ncorto
            Else
                aux = carpetaLong
            End If
            Carpetas.Add(aux)
            ' 
        Next
        todoOK = True
    End Sub
End Class
Class ClaseListaPerfiles
    Public todoOK As Boolean
    Public Perfiles As List(Of String)
    Public Sub New(OriPerfiles As String)
        todoOK = True
        Perfiles = New List(Of String)
        Dim dDir As New DirectoryInfo(OriPerfiles)
        Dim archivo As FileSystemInfo
        For Each archivo In dDir.GetFiles("*.R00")
            Dim auxS As String = archivo.Name
            auxS = auxS.Replace(".R00", "") ' quito la extensión
            Perfiles.Add(auxS)
        Next
    End Sub
End Class
Public Class ClaseListaCarpetasTraducir
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
    Function GeneraMFT_HTML(archivo As String) As Boolean
        Dim Par As structParametros
        Par = CargaParametros()
        Dim todoOK As Boolean = True ' optimismo
        Dim sbSalidaMF As New StringBuilder(String.Format("<HTML><HEAD><TITLE>{0}</TITLE></HEAD><BODY>", Me.NombrePaT))
        Dim nl As String = "<br>" & Environment.NewLine
        Dim lin() As String = { _
            "<H2><b>PaT name:</b> {0}</H2>" & nl, _
            "<b>Due date:</b> {0}" & nl, _
            "<b>Notes:</b>" & nl & " {0}" & nl, _
            "<b>Folder references:</b>" & nl & "{0}" & nl, _
            "<b>Base project references:</b>" & nl & "{0}" & nl, _
            "" & nl
           }
        Dim s As Integer = 0
        lin(s) = String.Format(lin(s), Me.NombrePaT) : s += 1
        ' necesitamos fecha

        Dim kk As String = Me.FechaEntrega.ToString("dd/MMM/yyyy")
        lin(s) = String.Format(lin(s), kk) : s += 1
        lin(s) = String.Format(lin(s), Me.Notas) : s += 1
        ' las referencias las busco en la clase
        Dim auxS As String = ""
        Dim auxS2 As String = "" : auxS = ""
        For Each auxS2 In ListaReferencias.Referencias
            auxS &= auxS2 & nl
        Next
        lin(s) = String.Format(lin(s), auxS) : s += 1

        auxS = ""
        For Each auxS2 In ListaReferenciasProyectoBase.ReferenciasProyectoBase
            auxS &= auxS2 & nl
        Next
        lin(s) = String.Format(lin(s), auxS) : s += 1
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
        auxS &= "<td><b>IniCal</b></td>"
        auxS &= "<td><b>Rate</b></td>"
        sbSalidaMF.Append(IniFila & auxS & FinFila) ' añado cabecera
        For Each fila In Me.tCarpetasTraducir.Rows
            auxS = ""
            auxS &= String.Format("<td>{0}</td>", fila("Carpeta"))
            auxS &= String.Format("<td>{0}</td>", fila("ProMSS"))
            auxS &= String.Format("<td>{0}</td>", fila("ProBase"))
            auxS &= String.Format("<td>{0}</td>", fila("Idioma"))
            auxS &= String.Format("<td>{0}</td>", fila("ProIBM"))
            auxS &= String.Format("<td>{0}</td>", fila("Envio"))
            auxS &= String.Format("<td>{0}</td>", fila("Perfil"))
            If fila("bCNT") Then
                auxS &= String.Format("<td>{0}</td>", fila("CNT"))
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
        sbSalidaMF.Append("</BODY></HTML>")
        Dim sw As System.IO.StreamWriter
        sw = My.Computer.FileSystem.OpenTextFileWriter(archivo, True)
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
            Dim lin() As String = { _
                "<H2><b>PaT name:</b> {0}</H2>" & nl, _
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
            sbSalidaMF.Append("<b>Billing information:</b>" & nl)
            sbSalidaMF.Append("In order to ease the billing process, your invoice should be sent with this note. ")
            sbSalidaMF.Append("If you do not do so, it could delay such billing process.")
            sbSalidaMF.Append("So, if you agree with the billing information, include the following information in you invoice. " & nl)
            sbSalidaMF.Append("Thanks a lot!" & nl & nl)


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
            auxS &= String.Format("<td>{0}</td>", fila("Contaje"))
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
        Dim sw As System.IO.StreamWriter
        sw = My.Computer.FileSystem.OpenTextFileWriter(archivo, True)
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

    <Extension()> _
    Public Sub Refresh(ByVal uiElement As UIElement)
        uiElement.Dispatcher.Invoke(DispatcherPriority.Render, New Action(AddressOf EmptyMethod))
    End Sub

End Module