Public Class VentanaPreproceso
    Sub New(lista As List(Of String))
        ' This call is required by the designer.
        InitializeComponent()
        listaPat = lista
        currentPatindex = 0
        currentPat = lista(currentPatindex)
        Dim nl As String = vbNewLine
        Dim numpat As Integer = lista.Count
        Dim auxS As String = ""
        auxS = String.Format("Númber of PaT files : {0}", numpat) & nl
        auxS &= nl & "Pats found: " & nl
        For Each s In listaPat
            auxS &= s & nl
        Next
        auxS &= nl
        lblInfo01.Content = auxS
        lblInfo02.Content = String.Format("Current PaT: {0}", currentPat)
        If numpat = 1 Then btnNext.IsEnabled = False



    End Sub
    Public listaPat As List(Of String)
    Public Resultado As String
    Public currentPat As String
    Public currentPatindex As Integer
    Private Sub btnCancel_Click(sender As Object, e As RoutedEventArgs) Handles btnCancel.Click
        Me.Resultado = "Cancel"
        Me.Close()
    End Sub

    Private Sub btnNext_Click(sender As Object, e As RoutedEventArgs) Handles btnNext.Click
        currentPatindex += 1
        If currentPatindex = listaPat.Count Then currentPatindex = 0
        currentPat = listaPat(currentPatindex)
        lblInfo02.Content = String.Format("Current PaT: {0}", currentPat)
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As RoutedEventArgs) Handles btnDelete.Click

        Dim result As Integer =
            MsgBox(String.Format("Are you sure you want to delete {0}?", currentPat),
                   MsgBoxStyle.OkCancel, "Delete PaT file sent")
        If result = 1 Then ' OK was selected
            Me.Resultado = "Delete"
            Me.Close()
        End If

    End Sub

    Private Sub btnContinue_Click(sender As Object, e As RoutedEventArgs) Handles btnContinue.Click
        Me.Resultado = "Continue"
        Me.Close()
    End Sub
End Class
