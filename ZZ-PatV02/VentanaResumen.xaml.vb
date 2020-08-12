Public Class VentanaResumen
    Sub New(ContenidoHTML As String, EnvioCorreo As Boolean)

        ' This call is required by the designer.
        InitializeComponent()


        myWebBrowser.NavigateToString(ContenidoHTML)
        If EnvioCorreo Then
            lblSendStatus.Content = "Email will be sent to the translator"
            lblSendStatus.FontWeight = FontWeights.Normal
        Else
            lblSendStatus.Content = "Warning: You have selected NOT to send an email to the translator or translator is 'NONE'."
            lblSendStatus.FontWeight = FontWeights.Bold


        End If


    End Sub
    Private Sub btnCancelar_Click(sender As Object, e As RoutedEventArgs) Handles btnCancelar.Click
        Me.DialogResult = False
        Me.Close()
    End Sub

    Private Sub btnSendOK_Click(sender As Object, e As RoutedEventArgs) Handles btnSendOK.Click
        Me.DialogResult = True
        Me.Close()
    End Sub
End Class
