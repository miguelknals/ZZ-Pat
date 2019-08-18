Module Rutinas

    Function OTMEnEjecucion() As Boolean
        Dim p() As Process
        p = Process.GetProcessesByName("OpenTM2")
        If p.Count > 0 Then Return True
        Return False
    End Function
    Function AyudaZZInstallPaT()
        Console.WriteLine("Usage: ZZ-InstallPaT -p PatFile")
        Console.WriteLine("Available options:")
        Console.WriteLine()
        Console.WriteLine("     -p PatFile             PaT file")
        Console.WriteLine()
        Return Nothing
    End Function




End Module
