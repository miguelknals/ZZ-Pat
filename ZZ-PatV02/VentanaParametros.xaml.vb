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
    Dim nl As String = Environment.NewLine

    Sub New()

        ' This call is required by the designer.
        InitializeComponent()


        Dim Par As structParametros

        Par = CargaParametros()
        txtDirTrabajo.Text = Par.ParDirTemporal
        txtDirSalidaPaT.Text = Par.ParDirSalidaPaT
        txtLongInterno.Text = Par.ParLongInterno
        txtIVA.Text = CType(Par.ParIVA, String)
        txtIRPF.Text = CType(Par.ParIRPF, Single)
        txtTarifaPredeterminada.Text = CType(Par.ParTarifaPredeterminada, Single)
        chkParMasDeUnPB.IsChecked = Par.ParMasDeUnPB
        txtProveedor.Text = Par.ParProveedor
        txtCorreoGestor.Text = Par.ParCorreoGestor
        btnSLL.IsChecked = Par.ParMailSSL
        txtMailPuerto.Text = Par.ParMailPuerto
        txtMailContrase.Text = Par.ParMailContrase
        txtMailHost.Text = Par.ParMailHost
        txtMailUsuario.Text = Par.ParMailUsuario

        ' ahora el tipo origen

        ' los leeo


    End Sub

    Private Sub btnGuardar1_Click(sender As Object, e As RoutedEventArgs) Handles btnGuardar1.Click
        Dim Par As New structParametros
        Par = CargaParametros() ' cargo los valores anteriores por si quiero comparar algún valor anterio
        ' voy a ver que me hayan añadido el \

        Par.ParDirTemporal = CompruebaAntibarra(txtDirTrabajo.Text)
        Par.ParDirSalidaPaT = CompruebaAntibarra(txtDirSalidaPaT.Text)
        ' esto no lo dejo cambiar tan fácilmente
        Dim AuxLong As Long = 0
        Dim AuxS As String
        Dim HayError As Boolean = True
        Try
            AuxLong = Convert.ToInt64(txtLongInterno.Text)
            HayError = False
            If Math.Abs(AuxLong) > 10 ^ 15 Then
                HayError = True
            End If
        Catch ex As Exception
        Finally
            If HayError Then
                AuxS = String.Format("Cannot cast Long Internal for '{0}'. Use a big long lower than 999999999999999 (14 digits)", txtLongInterno.Text)
                MsgBox(AuxS, MsgBoxStyle.Critical, "Error: Cannot cast Long integer")
            End If
        End Try
        If HayError Then Exit Sub

        If AuxLong <> Par.ParLongInterno Then
                AuxS = "***  WARNING ***" & nl & nl
                AuxS &= "Looks you want to change the internal long integer. This value is use to " & nl
                AuxS &= "encrypt sensible parameters. If you change it all current sensible information " & nl
                AuxS &= "(as paswwords AND documents will become unsable.  " & nl & nl
                AuxS &= "This value only has to defined in the inital deployment of the application." & nl & nl
                AuxS &= "¿Are you sure?"
                Dim rc As Integer = MsgBox(AuxS, MsgBoxStyle.OkCancel, "*** You are about to change Long integer internal *** ")
            If rc <> 1 Then ' user not sure
                Exit Sub
            End If
        End If
        Par.ParLongInterno = AuxLong
        Par.ParIVA = CType(txtIVA.Text, Single) ' es nls dependiente
        Par.ParIRPF = CType(txtIRPF.Text, Single) ' es nls dependiente.
        Par.ParTarifaPredeterminada = CType(txtTarifaPredeterminada.Text, Single)
        Par.ParMasDeUnPB = chkParMasDeUnPB.IsChecked
        Par.ParProveedor = txtProveedor.Text
        Par.ParCorreoGestor = txtCorreoGestor.Text
        If btnSLL.IsChecked Then Par.ParMailSSL = True Else Par.ParMailSSL = False
        Par.ParMailPuerto = txtMailPuerto.Text
        ' Voy a intentar descifrarla
        AuxS = DecryptWithKey(txtMailContrase.Text, Par.ParLongInterno)
        If AuxS = "Unable to decypher" Then
            AuxS = "***  ERROR ***" & nl & nl
            AuxS &= "There is problem with you email password. Looks like: " & nl & nl
            AuxS &= "1) You have not encrypt it. Use the 'Encrypt' button to encrypt it and paste in the 'Encrypted passowrod' field. " & nl & nl
            AuxS &= "2) You have changed the 'Long internal' integer field. You need to reenter the password as in 1). " & nl & nl
            AuxS &= "In order to continue you need to fix this error."
            MsgBox(AuxS, MsgBoxStyle.Critical, "Error: Encrypted password not compatible wiht Internal Long Integer")
            Exit Sub
        End If
        Par.ParMailContrase = txtMailContrase.Text
        Par.ParMailHost = txtMailHost.Text
        Par.ParMailUsuario = txtMailUsuario.Text



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

    Private Sub btnEncript_Click(sender As Object, e As RoutedEventArgs) Handles btnEncript.Click
        Dim AuxS As String

        If txtEncrypt.Text.Trim() = "" Then
            AuxS = "Please, type a password to encrypt" & nl & nl
            MsgBox(AuxS, MsgBoxStyle.Critical, "Info: Type a password to encrypt")
            Exit Sub
        End If
        Dim AuxLong As Long = 0 : Dim HayError = True
        Try
            AuxLong = Convert.ToInt64(txtLongInterno.Text)
            HayError = False
            If Math.Abs(AuxLong) > 10 ^ 15 Then
                HayError = True
            End If
        Catch ex As Exception
        Finally
            If HayError Then
                AuxS = String.Format("Cannot cast Long Internal for '{0}'. Use a big long 14 digit max like 12345678901234", txtLongInterno.Text)
                MsgBox(AuxS, MsgBoxStyle.Critical, "Error: Cannot cast Long integer greater than 0")
            End If
        End Try
        If HayError Then Exit Sub


        txtEncrypt.Text = EncryptWithKey(txtEncrypt.Text, AuxLong)
        'Dim veri As String = DecryptWithKey(txtEncrypt.Text, AuxLong)
    End Sub
End Class
