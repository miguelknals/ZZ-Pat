Imports System.IO
Public Class VentanaParametros
    Dim nl As String = Environment.NewLine

    Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        ' cargo el ddw
        cargaddwTiposdeDestino()

        Dim Par As structParametros
        Par = CargaParametros()
        '
        ' test
        For Each s In ddwTipoDestino.Items
            If s = Par.DestTipoDestino Then
                ddwTipoDestino.SelectedItem = s
                Exit For
            End If
        Next
        txtNombreTraductor.Text = Par.DestNombre
        txtCorreoTraductor.Text = Par.DestCorreo
        txtSharedDriveHost.Text = Par.DestSharedDriveHost
        txtSharedDriveNombre.Text = Par.DestSharedDriveNombre
        txtSharedDriveUsuario.Text = Par.DestSharedDriveUsuario
        txtSharedDriveContrasenya.Text = Par.DestSharedDriveContrasenya
        txtFTPHost.Text = Par.DestFTPHost
        txtFTPNombre.Text = Par.DestFTPNombre
        txtFTPUsuario.Text = Par.DestFTPUsuario
        txtFTPContrasenya.Text = Par.DestFTPContrasenya
        txtLongInterno.Text = Par.DestLongInterno
        ' 
        txtLocalPatDir.Text = Par.DestLocalPatDir



    End Sub
    Function cargaddwTiposdeDestino()
        With ddwTipoDestino
            .Items.Clear()
            .Items.Add("PATH")
            .Items.Add("SHARED_DRIVE")
            .Items.Add("FTP")
            .Items.Add("FTPES")
            .Items.Add("SFTP")
        End With
        Return Nothing
    End Function

    Private Sub btnGuardarActualizar_Click(sender As Object, e As RoutedEventArgs) Handles btnGuardarActualizar.Click

        Dim Par As structParametros
        Par = CargaParametros()
        ' Miro la contraseña a ver si la puedo descifrar
        ' Voy a intentar descifrarla
        Dim AuxS As String = ""
        AuxS = DecryptWithKey(txtFTPContrasenya.Text, Par.DestLongInterno)
        If AuxS = "Unable to decypher" Then
            AuxS = "***  ERROR ***" & nl & nl
            AuxS &= "There is problem with you FTP password. Looks like: " & nl & nl
            AuxS &= "1) You have not encrypt it. Use the 'Encrypt' button to encrypt it and paste in the FTP encrypted password field. " & nl & nl
            AuxS &= "2) Parameter 'Long internal' integer field has recently changed. You need to reenter the password as in 1). " & nl & nl
            AuxS &= "In order to continue you need to fix this error."
            MsgBox(AuxS, MsgBoxStyle.Critical, "Error: Encrypted password not compatible wiht Internal Long Integer")
            Exit Sub
        End If
        If InStr(txtNombreTraductor.Text, " ") > 0 Then
            MsgBox("Spaces are not allowed in target name. Use dahs for instance.", MsgBoxStyle.Critical, "Translator name")
            Exit Sub
        End If
        If Not Directory.Exists(txtLocalPatDir.Text) Then
            MsgBox("Local directory must exit. Please create it", MsgBoxStyle.Critical, "Local directory does not exist")
            Exit Sub
        End If
        Dim AuxLong As Long = 0
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
                AuxS = String.Format("Cannot cast Long Internal for '{0}'. Use a big long max 14 digits like 12345678901234", txtLongInterno.Text)
                MsgBox(AuxS, MsgBoxStyle.Critical, "Error: Cannot cast Long integer")
            End If
        End Try
        If HayError Then Exit Sub


        With Par
            .DestNombre = txtNombreTraductor.Text
            .DestCorreo = txtCorreoTraductor.Text
            .DestTipoDestino = ddwTipoDestino.SelectedItem.ToString
            .DestSharedDriveHost = txtSharedDriveHost.Text
            .DestSharedDriveNombre = txtSharedDriveNombre.Text
            .DestSharedDriveUsuario = txtSharedDriveUsuario.Text
            .DestSharedDriveContrasenya = txtSharedDriveContrasenya.Text
            .DestFTPHost = txtFTPHost.Text
            .DestFTPNombre = txtFTPNombre.Text
            .DestFTPUsuario = txtFTPUsuario.Text
            .DestFTPContrasenya = txtFTPContrasenya.Text
            .DestLongInterno = AuxLong

            '
            .DestLocalPatDir = txtLocalPatDir.Text
            End With
            ' 
            GuardaParametros(Par)
            Me.Close()
    End Sub

    Private Sub btnEncrypt_Click(sender As Object, e As RoutedEventArgs) Handles btnEncrypt.Click


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
                MsgBox(AuxS, MsgBoxStyle.Critical, "Error: Cannot cast Long integer")
            End If
        End Try
        If HayError Then Exit Sub


        txtEncrypt.Text = EncryptWithKey(txtEncrypt.Text, AuxLong)
        'Dim veri As String = DecryptWithKey(txtEncrypt.Text, AuxLong)

    End Sub
End Class
