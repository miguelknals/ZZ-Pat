Imports System
Imports System.IO
Public Class VentanaMain
    Dim SendF As New ClaseSendF
    Dim ReceiveF As New ClaseReceiveF
    Dim Parameters As New ClaseParameters
    Dim Targets As New ClaseTargets
    ' variables entre carpetas
    Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.
        Dim ViaEjecutable As String = System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase
        ViaEjecutable = ViaEjecutable.Replace("file:///", "")
        'Dim fechaEjecutablee As String = File.GetCreationTime(System.Reflection.Assembly.GetExecutingAssembly().Location)
        Dim arch As FileInfo = New System.IO.FileInfo(System.Reflection.Assembly.GetExecutingAssembly().Location)
        Dim lastMod As DateTime = arch.LastWriteTime
        Me.Title = "ZZ-PaT vBeta " & lastMod & "  (" & ViaEjecutable & ")"

        

        FrameMain.Navigate(SendF)



    End Sub

    Private Sub MenuItem_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub

    Private Sub MenuItem_Click_SendF(sender As Object, e As RoutedEventArgs)

        FrameMain.Navigate(SendF)
        ' es lo msimo FrameMain.NavigationService.Navigate(SendF) 
    End Sub

    Private Sub MenuItem_Click_ReceiveF(sender As Object, e As RoutedEventArgs)
        FrameMain.Navigate(ReceiveF)
    End Sub

    Private Sub MenuItem_Click_Targets(sender As Object, e As RoutedEventArgs)
        FrameMain.Navigate(Targets)
    End Sub

    Private Sub MenuItem_Click_Parameters(sender As Object, e As RoutedEventArgs)
        FrameMain.Navigate(Parameters)

    End Sub

    Private Sub MenuItem_Click_Web_page(sender As Object, e As RoutedEventArgs)
        System.Diagnostics.Process.Start("http://www.mknals.com/010_1_TMT_ZZ-Pat.html")
    End Sub
End Class
