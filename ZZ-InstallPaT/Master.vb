Imports System.Threading

Module Master
    Sub main()
        Console.WriteLine("This is ZZ-InstallPaT waiting for a PaT file...")
        Dim cki As ConsoleKeyInfo
        ' Prevent example from ending if CTL+C is pressed.
        Console.TreatControlCAsInput = True

        Console.WriteLine("Press any combination of CTL, ALT, and SHIFT, and a console key.")
        Console.WriteLine("Press the Escape (Esc) key to quit: " + vbCrLf)
        Do
            cki = Console.ReadKey()
            Console.Write(" --- You pressed ")
            If (cki.Modifiers And ConsoleModifiers.Alt) <> 0 Then Console.Write("ALT+")
            If (cki.Modifiers And ConsoleModifiers.Shift) <> 0 Then Console.Write("SHIFT+")
            If (cki.Modifiers And ConsoleModifiers.Control) <> 0 Then Console.Write("CTL+")
            Console.WriteLine(cki.Key.ToString)
            Console.Write(".")
        Loop While cki.Key <> ConsoleKey.Escape


        Console.WriteLine("Press any key if you want fo finish...")

        
        Console.ReadLine()

    End Sub
End Module
