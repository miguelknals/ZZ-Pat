Class ClaseParameters
    Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        ' cargo los valores de combobox
        ' cargo los parámetros
        cbxTipoOrigen.Items.Clear()
        cbxTipoOrigen.Items.Add("NONE")
        cbxTipoOrigen.Items.Add("EXCEL")
        cbxTipoOrigen.Items.Add("POSTGRES")

        Dim Par As structParametros
        Par = CargaParametros()
        txtDirTrabajo.Text = Par.ParDirTemporal
        txtDirSalidaPaT.Text = Par.ParDirSalidaPaT
        txtDirTPot.Text = Par.ParDirTPot
        txtArchivoExcel.Text = Par.ParArchivoExcel
        txtSQL.Text = Par.ParSQL
        txtSecretoServidor.Text = Decrypt(Par.ParSecretoServidor)
        txtIVA.Text = CType(Par.ParIVA, String)
        txtIRPF.Text = CType(Par.ParIRPF, Single)
        txtTarifaPredeterminada.Text = CType(Par.ParTarifaPredeterminada, Single)
        chkRecopilarInfo.IsChecked = Par.ParRecopilarUso
        chkParMasDeUnPB.IsChecked = Par.ParMasDeUnPB
        txtHostUDP.Text = Par.ParHostUDP
        txtPuertoUDP.Text = Par.ParPuertoUDP
        txtProveedor.Text = Par.ParProveedor
        txtCorreoGestor.Text = Par.ParCorreoGestor
        txtSQL_Alternativo.Text = Par.ParSQL_Alternativo
        txtPGHost.Text = Par.ParPGHost
        txtPGPuerto.Text = Par.ParPGPuerto
        txtPGBaseDatos.Text = Par.ParPGBaseDatos
        txtPGUsuario.Text = Par.ParPGUsuario
        txtPGContrasenya.Text = Decrypt(Par.ParPGContrasenya)
        txtPGContrasenya.Text = Replace(txtPGContrasenya.Text, txtPGUsuario.Text, "", , 1)
        ' ahora el tipo origen
        Dim posi As Integer = 0 : Dim Encontrado As Boolean = False
        For Each item In cbxTipoOrigen.Items
            If item = Par.ParTipoOrigen Then Encontrado = True : Exit For
            posi += 1
        Next
        If Encontrado = False Then
            MsgBox("The source type {0} is not allowed. NONE will be used instead", MsgBoxStyle.Exclamation)
            posi = 0
        End If
        cbxTipoOrigen.SelectedIndex = posi

    End Sub

    Private Sub btnGuardar1_Click(sender As Object, e As RoutedEventArgs) Handles btnGuardar1.Click
        Dim Par As New structParametros
        ' voy a ver que me hayan añadido el \

        Par.ParDirTemporal = CompruebaAntibarra(txtDirTrabajo.Text)
        Par.ParDirSalidaPaT = CompruebaAntibarra(txtDirSalidaPaT.Text)
        Par.ParDirTPot = CompruebaAntibarra(txtDirTPot.Text)
        Par.ParArchivoExcel = txtArchivoExcel.Text
        Par.ParSQL = txtSQL.Text
        Par.ParSecretoServidor = Encrypt(txtSecretoServidor.Text)
        Par.ParIVA = CType(txtIVA.Text, Single) ' es nls dependiente
        Par.ParIRPF = CType(txtIRPF.Text, Single) ' es nls dependiente.
        Par.ParTarifaPredeterminada = CType(txtTarifaPredeterminada.Text, Single)
        Par.ParRecopilarUso = chkRecopilarInfo.IsChecked
        Par.ParMasDeUnPB = chkParMasDeUnPB.IsChecked
        Par.ParHostUDP = txtHostUDP.Text
        Par.ParPuertoUDP = CType(txtPuertoUDP.Text, Long)
        Par.ParProveedor = txtProveedor.Text
        Par.ParCorreoGestor = txtCorreoGestor.Text
        Par.ParSQL_Alternativo = txtSQL_Alternativo.Text
        Par.ParPGHost = txtPGHost.Text
        Par.ParPGPuerto = txtPGPuerto.Text
        Par.ParPGBaseDatos = txtPGBaseDatos.Text
        Par.ParPGUsuario = txtPGUsuario.Text
        Par.ParPGContrasenya = Encrypt(txtPGUsuario.Text & txtPGContrasenya.Text)
        Par.ParTipoOrigen = cbxTipoOrigen.SelectedItem


        GuardaParametros(Par)
        '  debería actualizar SendF
        MsgBox("Done", MsgBoxStyle.Information, "Saving parameters")


    End Sub
    Function CompruebaAntibarra(s As String) As String
        If Len(s) > 1 Then
            If s.Substring(Len(s) - 1, 1) <> "\" Then
                s &= "\"
            End If
        End If
        Return s
    End Function
End Class
