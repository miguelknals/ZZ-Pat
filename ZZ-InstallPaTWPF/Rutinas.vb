Module Rutinas
    Function OTMEnEjecucion() As Boolean
        Dim p() As Process
        p = Process.GetProcessesByName("OpenTM2")
        If p.Count > 0 Then Return True
        Return False
    End Function
End Module
