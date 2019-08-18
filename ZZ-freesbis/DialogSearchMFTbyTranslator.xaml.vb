Imports System.Xml
Imports System.Xml.Serialization
Imports System.IO
Imports System.Data
Public Class DialogSearchMFTbyTranslator
    Dim DirectorioGestor As String = ""
    Dim ArchivoMFTconPathlocal As String = ""
    Dim listadestinos As structListaDestinos
    Dim nl As String = Environment.NewLine

    Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        listadestinos = CargaDestinos()
        If listadestinos.ListaDestino.Count = 0 Then ' no tengo nada
            ddwDestinos.SelectedIndex = 0
        Else ' tengo destinos
            Dim destino As New structDestino
            Dim posiDestino As Integer = 0 ' 
            Dim auxi As Integer = 0
            ' cargo la lista de usuarios
            For Each destino In listadestinos.ListaDestino
                ddwDestinos.Items.Add(destino.DestNombre)

            Next
            ' en principio selecciono el primero
            ddwDestinos.SelectedIndex = posiDestino ' desencadena Changed
            Dim DestinoSeleccionado As String = ddwDestinos.SelectedItem

            ActualizaVentaDestinosCon(DestinoSeleccionado)

        End If
    End Sub
    Function ActualizaVentaDestinosCon(destinoseleccionado As String)
        Dim destino As New structDestino
        If listadestinos.ListaDestino.Count = 0 Then ' no tengo nada
            'ddwTipoDestino.SelectedIndex = 0
        Else ' tengo destinos voy a buscarlo
            For Each destino In listadestinos.ListaDestino
                If destino.DestNombre = destinoseleccionado Then  ' te encontré
                    Exit For
                End If
            Next
            'ddwDestinos.SelectedIndex = posiSele ' lo selecciono
            Dim AuxS As String = "" ' aquí construyo la info del destino

            AuxS &= "Translator -> " & destino.DestNombre
            AuxS &= nl & "Email -> " & destino.DestCorreo
            AuxS &= nl & "Target type -> " & destino.DestTipoDestino
            ' Es en el traductor no lo visualizo
            ' AuxS &= nl & destino.DestPath 
            Select Case destino.DestTipoDestino
                Case "SHARED_DRIVE"
                    AuxS &= nl & "Share -> " & destino.DestSharedDriveNombre
                    AuxS &= nl & "Share host (N/D) -> " & destino.DestSharedDriveHost
                    AuxS &= nl & "Share user (N/D) -> " & destino.DestSharedDriveUsuario
                    ' no hace falta enseñar la contraseña.
                    'Dim auxs2 As String
                    'auxs2 = Decrypt(destino.DestSharedDriveContrasenya)
                    'AuxS &= nl & Replace(auxs2, destino.DestNombre, "")
                Case "FTP"
                    AuxS &= nl & "FTP path -> " & destino.DestFTPNombre
                    AuxS &= nl & "FTP host -> " & destino.DestFTPHost
                    AuxS &= nl & "FTP user -> " & destino.DestFTPUsuario
                    'auxs2 = Decrypt(destino.DestFTPContrasenya)
                    'AuxS &= nl & Replace(auxs2, destino.DestNombre, "")
                Case "PATH"
                    AuxS &= nl & "Diretory -> " & destino.DestPath
                Case Else
                    AuxS &= nl & "*** Target type not supported ****"

            End Select
            txtInfoTranslator.Text = AuxS
            ' Ahora voy a intentar leer los MFT si hay alguno
            lbxMFT.Items.Clear()
            Select Case destino.DestTipoDestino
                Case "SHARED_DRIVE", "PATH"

                    ' SHARED DRIVE voy a leer los archivos que hay destino.DestSharedDriveNombre
                    ' PATH voy a leer los archivos que hay destino.DestPath
                    Dim dire As String = destino.DestSharedDriveNombre
                    If destino.DestTipoDestino = "PATH" Then dire = destino.DestPath
                    Dim dirinfo As DirectoryInfo
                    Try
                        If Not Directory.Exists(dire) Then
                            AuxS = "Target dir &" & dire & " does not exist. Looks like target is not correctly defined. " _
                                & nl & "Cancel the dialog and specify the MFT file manually."
                            MsgBox(AuxS, MsgBoxStyle.Exclamation, "Ops..shared drive/directory does not exist")
                            ' no hago nada más
                        Else ' el directorio existe así que voy a leer
                            dirinfo = New DirectoryInfo(dire)
                            For Each File In dirinfo.GetFiles("*_MFT.XML")
                                ' debo abrir y leerlo para encontrar el traductor.
                                Dim objStreamReader As StreamReader
                                Dim strlinea As String
                                objStreamReader = New StreamReader(File.FullName)
                                strlinea = objStreamReader.ReadLine
                                Do While Not strlinea Is Nothing
                                    If InStr(strlinea, "<NombreTraductor>") <> 0 Then ' la linea del traductor
                                        If InStr(strlinea, destino.DestNombre) <> 0 Then ' bingo
                                            lbxMFT.Items.Add(File.Name)
                                            Exit Do ' he acabado
                                        End If
                                    End If
                                    strlinea = objStreamReader.ReadLine ' siguiente linea
                                Loop
                                objStreamReader.Close()
                            Next
                        End If

                    Catch ex As Exception
                        AuxS = "Cannot read files in &" & dire & ". " & nl & "Cancel the dialog and specify the MFT file manually"
                        MsgBox(AuxS, MsgBoxStyle.Exclamation, "Ops..cannot read directory")
                    End Try
                Case "FTP"
                    MsgBox("Right now MFT retrivial from a FTP directory is not suported", _
                           MsgBoxStyle.Information, "FTP repository not yet supported")

                Case "PATH"



                Case Else
                    ' ya he avisado que era un tipo no soportado
            End Select

        End If
        Return Nothing
    End Function

    Private Sub btnCancel_Click(sender As Object, e As RoutedEventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub ddwDestinos_DropDownClosed(sender As Object, e As EventArgs) Handles ddwDestinos.DropDownClosed
        Dim DestinoSeleccionado As String = ddwDestinos.SelectedItem
        ActualizaVentaDestinosCon(DestinoSeleccionado)
    End Sub

    Private Sub ddwDestinos_KeyDown(sender As Object, e As KeyEventArgs) Handles ddwDestinos.KeyDown
        'Dim DestinoSeleccionado As String = ddwDestinos.SelectedItem
        'ActualizaVentaDestinosCon(DestinoSeleccionado)
    End Sub

    Private Sub ddwDestinos_KeyUp(sender As Object, e As KeyEventArgs) Handles ddwDestinos.KeyUp
        'Dim DestinoSeleccionado As String = ddwDestinos.SelectedItem
        ' ActualizaVentaDestinosCon(DestinoSeleccionado)
    End Sub

    Private Sub btnSelectAndClose_Click(sender As Object, e As RoutedEventArgs) Handles btnSelectAndClose.Click
        ' Primero tengo que elegir el seleccionado.
        Dim auxS As String = ""
        If lbxMFT.SelectedIndex < 0 Then
            AuxS = "You must select a MFT file. If you do not want or there are no MFT files, pelase cancel this dialog. "
            MsgBox(AuxS, MsgBoxStyle.Information, "You must select a MFT file")
            Exit Sub
        End If

        ' primero el archivo MFT seleccionado
        Dim destino As New structDestino
        ' Tengo el nombre del path, necesito destino (traductor)
        listadestinos = CargaDestinos()
        If listadestinos.ListaDestino.Count = 0 Then ' no tengo nada
            auxS = "Internal error. No targets in lbxMFT_SelectionChanged "
            MsgBox(auxS, MsgBoxStyle.Information, "Internal error")
            Exit Sub
        Else ' tengo destinos
            For Each destino In listadestinos.ListaDestino
                If destino.DestNombre = ddwDestinos.SelectedItem Then
                    Exit For ' tengo el destino 
                End If
            Next
        End If

        ' quiere decir que tenemos un candidato
        Dim ArchivoMF As String = lbxMFT.SelectedItem
        Dim ArchivoMFconPath As String = ""
        ' necesitaré las careptas
        Dim ListaCarpetasTraducir As New ClaseListaCarpetasTraducir
        '
        ' debo leerlo, pero me falta el path que tengo que obtner
        ' a partir de la info del 
        Select Case destino.DestTipoDestino
            Case "SHARED_DRIVE", "PATH"

                ' SHARED DRIVE voy a leer los archivos que hay destino.DestSharedDriveNombre
                ' PATH voy a leer los archivos que hay destino.DestPath
                Dim dire As String = destino.DestSharedDriveNombre
                If destino.DestTipoDestino = "PATH" Then dire = destino.DestPath
                If Not Directory.Exists(dire) Then
                    AuxS = "Target dir &" & dire & " does not exist. Looks like target is not correctly defined. " _
                        & nl & "Cancel the dialog and specify the MFT file manually."
                    MsgBox(AuxS, MsgBoxStyle.Exclamation, "Ops..shared drive/directory does not exist")
                    ' no hago nada más
                Else ' el directorio existe así que voy a leer el archivo
                    ArchivoMFconPath = dire
                    If Not ArchivoMFconPath.EndsWith("\") Then ArchivoMFconPath &= "\"
                    ArchivoMFconPath &= ArchivoMF
                    ArchivoMFTconPathlocal = DirectorioGestor
                    If Not ArchivoMFTconPathlocal.EndsWith("\") Then ArchivoMFTconPathlocal &= "\"
                    ArchivoMFTconPathlocal &= ArchivoMF

                    ' copio
                    Try
                        If Not Directory.Exists(DirectorioGestor) Then
                            auxS = "Project manager PAT directory does not exist in the local file. The automatic copy cannot be done. Copy the MFT manually. Cancel this dialog or select another MFT file."
                            MsgBox(auxS, MsgBoxStyle.Information, "PM PAT directory does not exist.")
                            Directory.CreateDirectory(DirectorioGestor)
                        End If

                        ' ahora copio los archivos de mis referencias
                        File.Copy(ArchivoMFconPath, ArchivoMFTconPathlocal, True)

                    Catch ex As Exception
                        auxS = "Internal error in btnSelectAndClose_Click: Cannot copy from {0} to {1}"
                        auxS = String.Format(auxS, ArchivoMFconPath, ArchivoMFTconPathlocal)
                        MsgBox(auxS, MsgBoxStyle.Information, "Internal error")
                        MsgBox("We have not been able to zip the references into de path file, otherwise, the pat file is fine.", MsgBoxStyle.Exclamation)
                        Exit Sub
                    End Try

                End If

            Case "FTP"
                MsgBox("Right now MFT retrivial from a FTP directory is not suported", _
                       MsgBoxStyle.Information, "FTP repository not yet supported")
            Case Else
                ' ya he avisado que era un tipo no soportado
        End Select


        Close()
    End Sub
    Function RetreiveMFTFromTranslator() As String ' para recuperar el 
        Return ArchivoMFTconPathlocal
    End Function

    Private Sub lbxMFT_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles lbxMFT.SelectionChanged
        If lbxMFT.SelectedIndex < 0 Then
            ' por ejemplo al cambiar traductor, se carga la lista de elementos
            ' pero no hay nada seleccionado. ASí que salgo simplemente
            Exit Sub
        End If

        ' primero el archivo MFT seleccionado
        Dim AuxS As String = ""
        Dim destino As New structDestino

        ' Tengo el nombre del path, necesito destino (traductor)
        listadestinos = CargaDestinos()

        If listadestinos.ListaDestino.Count = 0 Then ' no tengo nada
            AuxS = "Internal error. No targets in lbxMFT_SelectionChanged "
            MsgBox(AuxS, MsgBoxStyle.Information, "Internal error")
            Exit Sub
        Else ' tengo destinos

            For Each destino In listadestinos.ListaDestino
                If destino.DestNombre = ddwDestinos.SelectedItem Then
                    Exit For ' tengo el destino 
                End If
            Next
        End If

        ' quiere decir que tenemos un candidato
        Dim ArchivoMF As String = lbxMFT.SelectedItem
        Dim ArchivoMFconPath As String = ""
        ' necesitaré las careptas
        Dim ListaCarpetasTraducir As New ClaseListaCarpetasTraducir
        DirectorioGestor = "" ' lo borro
        '
        ' debo leerlo, pero me falta el path que tengo que obtner
        ' a partir de la info del 
        Select Case destino.DestTipoDestino
            Case "SHARED_DRIVE", "PATH"

                ' SHARED DRIVE voy a leer los archivos que hay destino.DestSharedDriveNombre
                ' PATH voy a leer los archivos que hay destino.DestPath
                Dim dire As String = destino.DestSharedDriveNombre
                If destino.DestTipoDestino = "PATH" Then dire = destino.DestPath
                If Not Directory.Exists(dire) Then
                    AuxS = "Target dir &" & dire & " does not exist. Looks like target is not correctly defined. " _
                        & nl & "Cancel the dialog and specify the MFT file manually."
                    MsgBox(AuxS, MsgBoxStyle.Exclamation, "Ops..shared drive/directory does not exist")
                    ' no hago nada más
                Else ' el directorio existe así que voy a leer el archivo
                    ArchivoMFconPath = dire
                    If Not ArchivoMFconPath.EndsWith("\") Then ArchivoMFconPath &= "\"
                    ArchivoMFconPath &= ArchivoMF
                    ' ahora voy a deserializar
                    Try
                        Dim objStreamReader As New StreamReader(ArchivoMFconPath)
                        Dim x As New XmlSerializer(ListaCarpetasTraducir.GetType)
                        ListaCarpetasTraducir = x.Deserialize(objStreamReader)
                        objStreamReader.Close()
                    Catch ex As Exception
                        AuxS = "Fatal error: Looks like xml manifest is not valid. "
                        MsgBox(AuxS, MsgBoxStyle.Critical, "MFT file not valid")
                        ' no puedo hacer nada, el archivo está mal
                        Exit Sub
                    End Try
                    ' ahora simplemente cumplimento el 
                    Dim fila As DataRow
                    AuxS = "Folder / Shipment / Wordcount" & nl
                    For Each fila In ListaCarpetasTraducir.tCarpetasTraducir.Rows
                        AuxS &= String.Format("{0}/ {1} / {2}", fila("Carpeta"), fila("Envio"), fila("CNT")) & nl
                    Next
                End If
                DirectorioGestor = ListaCarpetasTraducir.DirectorioGestor ' por si tengo que enviar
                AuxS &= nl & "PM working dir -> " & DirectorioGestor & nl

                txtFoldersInfo.Text = AuxS

            Case "FTP"
                MsgBox("Right now MFT retrivial from a FTP directory is not suported", _
                       MsgBoxStyle.Information, "FTP repository not yet supported")
            Case Else
                ' ya he avisado que era un tipo no soportado
        End Select

    End Sub
    


    Private Sub ddwDestinos_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles ddwDestinos.SelectionChanged
        ' debo borrar
        txtFoldersInfo.Clear()
        txtInfoTranslator.Clear()
        lbxMFT.Items.Clear()
        DirectorioGestor = ""

    End Sub
End Class
