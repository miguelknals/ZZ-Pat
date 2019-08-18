Public Class VentanaParametros
    ' por razones historicas (los parémetros se guardaban en un XML)
    ' hay dos niveles de los parámetros, uno a nivel de Studio
    ' y otro a nivel de ZZ-Pat. Mejor no usar nunca My.Setting...
    ' utilizar la strcuPAr. Por eso, al añadir un parámetro hay
    ' que darlo de alto en cuatro sitios
    ' Aquí
    ' Sub New
    ' Botón de guardar
    ' Y en CargaParametros
    ' y GudaraParametros
    ' ( Y darlo de alta en strucParametros)
    Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' cargo los parámetros

        Dim Par As structParametros
        Par = CargaParametros()
        txtDirTrabajo.Text = Par.ParDirTemporal
        txtDirSalidaPaT.Text = Par.ParDirSalidaPaT
        txtDirTPot.Text = Par.ParDirTPot
        txtArchivoExcel.Text = Par.ParArchivoExcel
        txtSQL.Text = Par.ParSQL
        txtSecretoServidor.Text = Par.ParSecretoServidor
        txtIVA.Text = CType(Par.ParIVA, String)
        txtIRPF.Text = CType(Par.ParIRPF, Single)
        txtTarifaPredeterminada.Text = CType(Par.ParTarifaPredeterminada, Single)
        chkRecopilarInfo.IsChecked = Par.ParRecopilarUso
        chkParMasDeUnPB.IsChecked = Par.ParMasDeUnPB
        txtHostUDP.Text = Par.ParHostUDP
        txtPuertoUDP.Text = Par.ParPuertoUDP
        txtProveedor.Text = Par.ParProveedor

        ' los leeo


    End Sub

    Private Sub btnGuardar1_Click(sender As Object, e As RoutedEventArgs) Handles btnGuardar1.Click
        Dim Par As New structParametros
        ' voy a ver que me hayan añadido el \

        Par.ParDirTemporal = CompruebaAntibarra(txtDirTrabajo.Text)
        Par.ParDirSalidaPaT = CompruebaAntibarra(txtDirSalidaPaT.Text)
        Par.ParDirTPot = CompruebaAntibarra(txtDirTPot.Text)
        Par.ParArchivoExcel = txtArchivoExcel.Text
        Par.ParSQL = txtSQL.Text
        Par.ParSecretoServidor = txtSecretoServidor.Text
        Par.ParIVA = CType(txtIVA.Text, Single) ' es nls dependiente
        Par.ParIRPF = CType(txtIRPF.Text, Single) ' es nls dependiente.
        Par.ParTarifaPredeterminada = CType(txtTarifaPredeterminada.Text, Single)
        Par.ParRecopilarUso = chkRecopilarInfo.IsChecked
        Par.ParMasDeUnPB = chkParMasDeUnPB.IsChecked
        Par.ParHostUDP = txtHostUDP.Text
        Par.ParPuertoUDP = CType(txtPuertoUDP.Text, Long)
        Par.ParProveedor = txtProveedor.Text

        GuardaParametros(Par)
        Me.Close()
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
