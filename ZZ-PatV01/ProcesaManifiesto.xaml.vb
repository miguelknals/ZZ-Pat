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


Public Class ProcesaManifiesto
    Dim nl As String = Environment.NewLine
    Dim ListaCarpetasTraducir As New ClaseListaCarpetasTraducir
    ' aquí los parametros para toda la apliación
    Dim ParDirTemporal As String = ""
    Dim ParDirSalidaPaT As String = ""
    Dim ParDirTPot As String = ""
    Dim ParSecretoServidor As String = ""
    Dim ViaArchivoMF As String = ""
    Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Dim Par As structParametros
        Par = CargaParametros()
        ParDirTemporal = Par.ParDirTemporal
        ParDirSalidaPaT = Par.ParDirSalidaPaT
        ParDirTPot = Par.ParDirTPot
        ParSecretoServidor = Par.ParSecretoServidor


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
        ' vamos a jugar
        ' primero voy a leer el manifiesto
        txtSalida.Clear()
        ' voy a ver si existe el archivo
        If Not File.Exists(txtArchivoMF.Text) Then
            txtSalida.AppendText(String.Format("'' does not seem a file or is not found" & nl, txtArchivoMF.Text))
            MsgBox("Pls, drag and drop a PaT manifest file (*_MFT.XML)", MsgBoxStyle.Critical, "No manifiest file")
            Exit Sub
        End If

        Dim auxS As String = "" ' variable auxiliar
        ListaCarpetasTraducir = New ClaseListaCarpetasTraducir
        Dim ArchivoMF As String = txtArchivoMF.Text
        ViaArchivoMF = Path.GetDirectoryName(ArchivoMF)
        ' miro si acaba en \
        If ViaArchivoMF.Substring(Len(ViaArchivoMF) - 1, 1) <> "\" Then ViaArchivoMF &= "\"


        txtSalida.AppendText("Loading... " & ArchivoMF & " ....")
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

        Dim OriI As Integer = 0

        If ListaCarpetasTraducir.FirmaInstancia <> ListaCarpetasTraducir.ObtenHash(ParSecretoServidor) Then
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
        ' voy a ver que tengo al menos las carpetas
        Dim row As DataRow : Dim auxB As Boolean = True ' en principio suponemos que todo ok.
        Dim unidad As String = Path.GetPathRoot(ParDirTemporal).Substring(0, 1) ' la letra
        auxS = Path.GetPathRoot(ParDirTemporal) ' C:\
        auxS = auxS.Replace("\", "") ' C: que es lo que debo quitar
        Dim DirTemporalSinUnidad As String = ParDirTemporal.Replace(auxS, "")
        Dim mandato As String = "" : Dim opcion As String = ""
        For Each row In ListaCarpetasTraducir.tCarpetasTraducir.Rows
            Dim carpeta As String = row("Carpeta")
            Dim carpetaConPath = ViaArchivoMF & carpeta & ".FXP"
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
                    Dim carpetaWCT As String = carpeta & "_WCT" '
                    ' primero suprmiré la carpeta por si acaso
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
                    Dim ArchivoFinCal As String = ViaArchivoMF & String.Format("{0}_{1}_fin_cal.rpt", carpeta, row("Perfil"))
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


                    ' al llegar aquí ArchivoFinCal tiene el calculating final
                    ' voy a saco
                    Dim lines As String() = IO.File.ReadAllLines(ArchivoFinCal)
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
        ' tengo que gernar la OS

        Dim archivoOC As String = ViaArchivoMF & ListaCarpetasTraducir.NombrePaT.Replace(".PaT", "") & "_PO.HTM"

        If ListaCarpetasTraducir.GeneraOC_HTML(archivoOC, "") = False Then
            txtSalida.AppendText("Fatal error. We have not been able to generate the purchase order (1st Part)" & nl)
            Exit Sub
        Else
            txtSalida.AppendText("Generated 1st part purchase order!" & nl)

        End If
        If ListaCarpetasTraducir.GeneraOC_HTML(archivoOC, "FAC") = False Then
            txtSalida.AppendText("Fatal error. We have not been able to generate the purchase order (2nd Part)" & nl)
        Else
            txtSalida.AppendText("Generated 1st part purchase order!" & nl)
        End If
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
