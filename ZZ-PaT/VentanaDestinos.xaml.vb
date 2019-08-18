Public Class VentanaDestinos

    Sub New(NombreDestino As String)

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        ' leo la propiedad


        ' cargo el cbxTipoDestino
        cargaddwTiposdeDestino()
        ' ahora los destinos
        Dim listadestinos As structListaDestinos
        listadestinos = CargaDestinos()
        ' ahora lo cargo en mis campos
        If listadestinos.ListaDestino.Count = 0 Then ' no tengo nada
            ddwTipoDestino.SelectedIndex = 0
        Else ' tengo destinos
            Dim destino As New structDestino
            Dim posiDestino As Integer = 0 ' 
            Dim auxi As Integer = 0
            ' cargo la lista de usuarios
            For Each destino In listadestinos.ListaDestino
                ddwDestinos.Items.Add(destino.DestNombre)
                If NombreDestino = destino.DestNombre Then posiDestino = auxi ' el que seleccionare
                auxi += 1
            Next
            ' 
            ' en principio selecciono el primero
            ddwDestinos.SelectedIndex = posiDestino ' desencadena Changed
            Dim DestinoSeleccionado As String = ddwDestinos.SelectedItem
            
            ActualizaVentaDestinosCon(DestinoSeleccionado)


        End If
        ' el tipo lo tengo que buscar en el cbx


    End Sub

    Private Sub Exit_Click(sender As Object, e As RoutedEventArgs) Handles [Exit].Click
        Me.Close()
    End Sub

    Function cargaddwTiposdeDestino()
        With ddwTipoDestino
            .Items.Clear()
            .Items.Add("PATH")
            .Items.Add("SHARED_DRIVE")
            .Items.Add("FTP")
        End With
        Return Nothing
    End Function


    Private Sub btnGuardarActualizar_Click(sender As Object, e As RoutedEventArgs) Handles btnGuardarActualizar.Click
        ' si se modifica este código, es posible que tenga que modificarse delete un target
        Dim NuevoDestino As New structDestino
        With NuevoDestino
            .DestNombre = txtNombreTraductor.Text
            .DestCorreo = txtCorreoTraductor.Text
            .DestTipoDestino = ddwTipoDestino.SelectedItem.ToString
            .DestPath = txtViaDestino.Text
            .DestSharedDriveHost = txtSharedDriveHost.Text
            .DestSharedDriveNombre = txtSharedDriveNombre.Text
            .DestSharedDriveUsuario = txtSharedDriveUsuario.Text
            .DestSharedDriveContrasenya = Encrypt(txtNombreTraductor.Text & txtSharedDriveContrasenya.Text)

            .DestFTPHost = txtFTPHost.Text
            .DestFTPNombre = txtFTPNombre.Text
            .DestFTPUsuario = txtFTPUsuario.Text
            .DestFTPContrasenya = Encrypt(txtNombreTraductor.Text & txtFTPContrasenya.Text)
        End With
        ' 
        If InStr(txtNombreTraductor.Text, " ") > 0 Then
            MsgBox("Spaces are not allowed in target name", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        ' paso mi paquete a la función
        Dim infoDestino As New infoGuardaDestino(NuevoDestino)
        GuardaDestino(infoDestino)
        If infoDestino.todoOk Then
            ' voy a intentar borrar y añadir
            ddwDestinos.Items.Clear()
            'ddwDestinos = New ComboBox
            ' obtengo los nuevos target
            Dim listadestinos As structListaDestinos
            listadestinos = CargaDestinos()
            ' cargo el ddw
            For Each destino In listadestinos.ListaDestino
                ddwDestinos.Items.Add(destino.DestNombre)
            Next
            ' ahora si lo selecciono
            Dim posi As Integer = 0
            For Each destino In listadestinos.ListaDestino
                If destino.DestNombre = txtNombreTraductor.Text Then
                    ddwDestinos.SelectedIndex = posi
                    Dim DestinoSeleccionado As String = ddwDestinos.SelectedItem
                    ActualizaVentaDestinosCon(DestinoSeleccionado)
                    Exit For
                End If
                posi += 1
            Next
            MsgBox("Target has been updated/saved.", MsgBoxStyle.Information)
        Else
            MsgBox(String.Format("Error saving/updating target. Error {0}", infoDestino.info), MsgBoxStyle.Exclamation)

        End If
    End Sub

    Private Sub btnSuprimir_Click(sender As Object, e As RoutedEventArgs) Handles btnSuprimir.Click
        ' este código se basa en añadir
        Dim DestinoABorrar As String = txtNombreTraductor.Text ' sera uno existente o no existente (por haberlo modificado)
        Dim listadestinos As structListaDestinos
        listadestinos = CargaDestinos()
        ' cargo el ddw
        Dim posi As Integer = 0
        For Each destino In listadestinos.ListaDestino
            If destino.DestNombre = DestinoABorrar Then
                listadestinos.ListaDestino.RemoveAt(posi)
                Exit For
            End If
            posi += 1
        Next
        ' voy a intentar borrar y añadir
        GuardaArchivoDestinos(listadestinos)
        ddwDestinos.Items.Clear()
        listadestinos = CargaDestinos()
        ' cargo el ddw
        For Each destino In listadestinos.ListaDestino
            ddwDestinos.Items.Add(destino.DestNombre)
        Next
        If ddwDestinos.Items.Count >= 0 Then
            ddwDestinos.SelectedIndex = 0
            Dim DestinoSeleccionado As String = ddwDestinos.SelectedItem
            ActualizaVentaDestinosCon(DestinoSeleccionado)
        End If
        MsgBox("Target has been deleted/removed.", MsgBoxStyle.Information)

    End Sub



    Private Sub ddwDestinos_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)
        ' si no hay nada seleccionado salgo
        'If ddwDestinos.SelectedIndex < 0 Then Exit Sub
        ' me ha cambiado el destino está copiado del inicio
        'Dim DestinoSeleccionado As String = ddwDestinos.SelectedItem
        '
        'ActualizaVentaDestinosCon(DestinoSeleccionado)

    End Sub
    Function ActualizaVentaDestinosCon(DestinoSeleccionado As String)
        Dim listadestinos As structListaDestinos
        listadestinos = CargaDestinos()
        ' ahora lo cargo en mis campos
        If listadestinos.ListaDestino.Count = 0 Then ' no tengo nada
            'ddwTipoDestino.SelectedIndex = 0
        Else ' tengo destinos voy a buscarlo
            Dim destino As New structDestino
            For Each destino In listadestinos.ListaDestino
                If destino.DestNombre = DestinoSeleccionado Then  ' te encontré
                    Exit For
                End If
            Next
            'ddwDestinos.SelectedIndex = posiSele ' lo selecciono
            For Each s In ddwTipoDestino.Items
                If s = destino.DestTipoDestino Then
                    ddwTipoDestino.SelectedItem = s
                    Exit For
                End If
            Next
            txtNombreTraductor.Text = destino.DestNombre
            txtCorreoTraductor.Text = destino.DestCorreo
            txtViaDestino.Text = destino.DestPath
            txtSharedDriveHost.Text = destino.DestSharedDriveHost
            txtSharedDriveNombre.Text = destino.DestSharedDriveNombre
            txtSharedDriveUsuario.Text = destino.DestSharedDriveUsuario
            Dim AUXs As String = ""
            AUXs = Decrypt(destino.DestSharedDriveContrasenya)
            txtSharedDriveContrasenya.Text = Replace(AUXs, destino.DestNombre, "")
            txtFTPHost.Text = destino.DestFTPHost
            txtFTPNombre.Text = destino.DestFTPNombre
            txtFTPUsuario.Text = destino.DestFTPUsuario
            AUXs = Decrypt(destino.DestFTPContrasenya)
            txtFTPContrasenya.Text = Replace(AUXs, destino.DestNombre, "")

        End If
        Return Nothing
    End Function

    Private Sub ddwDestinos_DropDownClosed(sender As Object, e As EventArgs) Handles ddwDestinos.DropDownClosed
        Dim DestinoSeleccionado As String = ddwDestinos.SelectedItem
        ActualizaVentaDestinosCon(DestinoSeleccionado)
    End Sub

    Private Sub ddwDestinos_KeyDown(sender As Object, e As KeyEventArgs) Handles ddwDestinos.KeyDown
        Dim DestinoSeleccionado As String = ddwDestinos.SelectedItem
        ActualizaVentaDestinosCon(DestinoSeleccionado)
    End Sub

    Private Sub ddwDestinos_KeyUp(sender As Object, e As KeyEventArgs) Handles ddwDestinos.KeyUp
        Dim DestinoSeleccionado As String = ddwDestinos.SelectedItem
        ActualizaVentaDestinosCon(DestinoSeleccionado)
    End Sub

    Function ObtenDestino() As String
        Dim destino As String = ""
        If ddwDestinos.SelectedIndex >= 0 Then
            destino = ddwDestinos.SelectedItem
        End If
        Return destino
    End Function

End Class
