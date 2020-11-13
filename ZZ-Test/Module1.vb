Module Module1

    Sub Main()
        Dim nl As String = vbNewLine
        Dim auxS As String = ""
        Console.WriteLine("Prueba")
        Dim var As String = "PATH"
        Console.WriteLine("Variable del sistema-> " & var)
        Dim ViaSistema = Environment.GetEnvironmentVariable("PATH")
        Console.WriteLine("Resultado GetEnvironmentvalriable(""{0}"") -> {1}", var, ViaSistema)
        Console.WriteLine("voy a buscar C:\OTM\WIN")
        Dim auxi As Integer = InStr(ViaSistema, ":\OTM\WIN")
        Console.WriteLine(String.Format("Poscion -> {0} ", auxi))
        Console.WriteLine("Según mecanismo oficial")
        Dim InfoOTM As New ClassInfoOTM ' obtiene info de OTM
        If InfoOTM.TodoOK = False Then ' OTM parece no instalado 
            ' para que el splashscreen no se me coma el diálogo
            auxS = String.Format("Error in OTM environment {0}", InfoOTM.Info) & nl
            Console.WriteLine(auxS)
            Dim DiscoOTM As String = InfoOTM.DiscoOTM
            Console.WriteLine("Disco OTM -> " + DiscoOTM)
        Else
            Console.WriteLine("Mecanismo oficial ha encontrado openm2")


        End If
        Console.Write("Fin")
    End Sub

End Module
