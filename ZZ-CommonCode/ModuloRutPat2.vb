Imports System.Security.Cryptography
Imports System.Text
Imports System.IO
Imports System.Xml.Serialization
Public Module ModuloRutPat2
    'Serializable()>
    Public Class classInfoEnvioCorreo
        Public MailSSL As Boolean
        Public MailPuerto As Integer
        Public MailContrase As String
        Public MailHost As String
        Public MailUsuario As String
        Public TodoOk As Boolean
        Function Serialize() As String
            Dim auxS As String = ""
            Dim x As New XmlSerializer(Me.GetType) ' serializo mi estructura
            Dim todoOK As Boolean = True ' optimismo
            Dim objSW As New StringWriter()
            Try
                x.Serialize(objSW, Me) ' guardo el par
                auxS = objSW.ToString()
            Catch ex As Exception
                todoOK = False
            Finally
                objSW.Close()
            End Try
            Return auxS
        End Function
        Function DeSerialize(s) As classInfoEnvioCorreo
            Dim InfoEnvioCorreo As New classInfoEnvioCorreo
            TodoOk = True
            Try
                Dim x As New XmlSerializer(Me.GetType) ' serializo mi estructura
                Dim string_reader As New StringReader(s)
                InfoEnvioCorreo = DirectCast(x.Deserialize(string_reader), classInfoEnvioCorreo)
                string_reader.Close()
            Catch ex As Exception
                TodoOk = False
            Finally
            End Try
            InfoEnvioCorreo.TodoOk = TodoOk
            Return InfoEnvioCorreo
        End Function

    End Class

    Public Function Encrypt(clearText As String) As String
        Dim EncryptionKey As String = "Y0URL0VEREVERSESMYS0UL"
        Dim clearBytes As Byte() = Encoding.Unicode.GetBytes(clearText)
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D,
             &H65, &H64, &H76, &H65, &H64, &H65,
             &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write)
                    cs.Write(clearBytes, 0, clearBytes.Length)
                    cs.Close()
                End Using
                clearText = Convert.ToBase64String(ms.ToArray())
            End Using
        End Using
        Return clearText
    End Function
    Public Function Decrypt(cipherText As String) As String
        Dim EncryptionKey As String = "Y0URL0VEREVERSESMYS0UL"
        Dim cipherBytes As Byte() = Convert.FromBase64String(cipherText)
        Using encryptor As Aes = Aes.Create()
            ' la clave tiene la pwd y la sal
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D,
             &H65, &H64, &H76, &H65, &H64, &H65,
             &H76})


            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write)
                    cs.Write(cipherBytes, 0, cipherBytes.Length)
                    cs.Close()
                End Using
                cipherText = Encoding.Unicode.GetString(ms.ToArray())
            End Using
        End Using
        Return cipherText
    End Function
    Public Function EncryptWithKey(clearText As String, LongiInternal As Long) As String
        Dim value As Double = Math.Sin(Math.Log(Convert.ToDouble(LongiInternal)))
        Dim expo As Double = Math.Floor(Math.Log10(Math.Abs(value)))
        Dim mant As Double = Math.Abs(value / Math.Pow(10, expo))
        Dim longmant As Double = Convert.ToInt64(Math.Abs(mant * 10 ^ 14))
        Dim onethirdbytes As Byte() = BitConverter.GetBytes(longmant) ' 8 byte lets make it longer 
        Dim EBytes(onethirdbytes.Length * 3 - 1) As Byte
        For i = 0 To 2
            onethirdbytes.CopyTo(EBytes, onethirdbytes.Length * i)
        Next
        Dim clearBytes As Byte() = Encoding.Unicode.GetBytes(clearText)
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EBytes, New Byte() {&H50, &H77, &H62, &H6F, &H21, &H4E,
             &H66, &H65, &H77, &H66, &H65, &H66, &H77}, 1000) ' clave y sal
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using ms As New MemoryStream()
                Using cs As New CryptoStream(ms, encryptor.CreateEncryptor(), CryptoStreamMode.Write)
                    cs.Write(clearBytes, 0, clearBytes.Length)
                    cs.Close()
                End Using
                clearText = Convert.ToBase64String(ms.ToArray())
            End Using
        End Using
        Return clearText
    End Function
    Public Function DecryptWithKey(cipherText As String, LongiInternal As Long) As String
        Dim value As Double = Math.Sin(Math.Log(Convert.ToDouble(LongiInternal)))
        Dim expo As Double = Math.Floor(Math.Log10(Math.Abs(value)))
        Dim mant As Double = Math.Abs(value / Math.Pow(10, expo))
        Dim longmant As Double = Convert.ToInt64(Math.Abs(mant * 10 ^ 14))
        Dim onethirdbytes As Byte() = BitConverter.GetBytes(longmant) ' 8 byte lets make it longer 
        Dim EBytes(onethirdbytes.Length * 3 - 1) As Byte
        For i = 0 To 2
            onethirdbytes.CopyTo(EBytes, onethirdbytes.Length * i)
        Next
        Using encryptor As Aes = Aes.Create()
            ' la clave tiene la pwd y la sal
            Dim pdb As New Rfc2898DeriveBytes(EBytes, New Byte() {&H50, &H77, &H62, &H6F, &H21, &H4E,
             &H66, &H65, &H77, &H66, &H65, &H66, &H77}, 1000) ' clave y sal


            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Try
                Dim cipherBytes As Byte() = Convert.FromBase64String(cipherText)

                Using ms As New MemoryStream()
                    Using cs As New CryptoStream(ms, encryptor.CreateDecryptor(), CryptoStreamMode.Write)
                        cs.Write(cipherBytes, 0, cipherBytes.Length)
                        cs.Close()
                    End Using
                    cipherText = Encoding.Unicode.GetString(ms.ToArray())
                End Using
            Catch ex As System.Security.Cryptography.CryptographicException
            Catch ex As Exception
                cipherText = "Unable to decypher"

            End Try



        End Using

        Return cipherText
    End Function
    Public Function FixHTML(HTML As String) As String
        Dim sb As New StringBuilder()
        Dim s As Char() = HTML.ToCharArray()
        For Each c As Char In s
            If Convert.ToInt32(c) > 127 Then
                sb.Append("&#" & Convert.ToInt32(c) & ";")
            Else
                sb.Append(c)
            End If
        Next
        Return sb.ToString()
    End Function


    Public Class ClassInfoOTM
        Public DiscoOTM As String
        Public TodoOK As Boolean
        Public Info As String
        Public Carpetas As List(Of String)
        Public Perfiles As List(Of String)

        Sub New()
            DiscoOTM = ""
            TodoOK = True ' optmismo
            Info = ""
            ' voy a buscar el path de OTM
            Dim ViaSistema = Environment.GetEnvironmentVariable("PATH")
            'voy a buscar C:\OTM\WIN
            Dim auxi As Integer = InStr(ViaSistema, ":\OTM\WIN")
            ' 
            If auxi = 0 Then ' OTM parece no instalado 
                TodoOK = False
                Info = "Cannot find :\OTM\WIN in PATH variable. Looks like OTM is no installed."
                Exit Sub
            End If
            DiscoOTM = ViaSistema.Substring(auxi - 2, 1)
            ' ahora la lista de carpetas
            Carpetas = New List(Of String)
            Dim dDir As New DirectoryInfo(String.Format("{0}:\OTM\PROPERTY", DiscoOTM))
            Dim archivo As FileSystemInfo

            For Each archivo In dDir.GetFiles("*.f00")
                ' no puede empezar por $$
                If InStr(archivo.Name, "$$", CompareMethod.Binary) <> 1 Then
                    Dim aux As String = ""
                    Dim carpetaLong As String = ""
                    ' 
                    Dim input As New FileStream(archivo.FullName, FileMode.Open, FileAccess.Read)
                    Dim reader As New BinaryReader(input)
                    ' por si acaso leeo hasta punto

                    Dim ncorto As String = "" : Dim posi As Integer = "0"
                    Dim b As Byte
                    While reader.PeekChar <> Asc(".")
                        b = reader.ReadByte()
                        ncorto &= Chr(b)
                        posi += 1
                    End While
                    '
                    ' busco el string 
                    reader.ReadBytes(Convert.ToInt32(&HA79) - posi)

                    While reader.PeekChar <> 0
                        b = reader.ReadByte()
                        carpetaLong &= Chr(b)
                    End While
                    reader.Close()
                    If Len(carpetaLong) = 0 Then
                        aux = ncorto
                    Else
                        aux = carpetaLong
                    End If
                    Carpetas.Add(aux)
                End If
                ' 
            Next
            Carpetas.Sort()
            ' ahora los perfiles.

            TodoOK = True
            Perfiles = New List(Of String)
            'Dim dDir As New DirectoryInfo(OriPerfiles)
            'Dim archivo As FileSystemInfo
            For Each archivo In dDir.GetFiles("*.R00")
                Dim auxS As String = archivo.Name
                auxS = auxS.Replace(".R00", "") ' quito la extensión
                Perfiles.Add(auxS)
            Next
            Perfiles.Sort()
        End Sub
    End Class

    Public Sub ejecuta(ByVal mandato As String, ByRef procEC As Integer, Optional debuga As Boolean = False)



        Dim procID As Integer
        Dim newProc As Diagnostics.Process
        If debuga Then logea(mandato)
        Dim startInfo As New ProcessStartInfo("OTMBATCH.EXE")
        startInfo.WindowStyle = ProcessWindowStyle.Minimized
        startInfo.Arguments = mandato
        startInfo.UseShellExecute = False ' si no pongo false, no recupera el path, me obligaría a indicarlos en otmbath.
        ' startInfo.CreateNoWindow = True (Efectivamente lo oculta, pero no se como hacer refresh)

        Try
            newProc = Process.Start(startInfo)
            procID = newProc.Id



            'Mouse.OverrideCursor = Cursors.Wait
            newProc.WaitForExit()
            'Mouse.OverrideCursor = Nothing
            procEC = -1
            If newProc.HasExited Then
                procEC = newProc.ExitCode
            End If
            If debuga Then logea("RC=" & procEC.ToString)

        Catch ex As Exception
            procEC = 999 ' error fatal
            If debuga Then logea("Excp= " & ex.ToString)
            Exit Sub
        End Try



    End Sub

    Sub mandatoZip(mandato As String, ByRef procEC As Integer,
                   Optional ByRef dirTra As String = "",
                   Optional ByRef debuga As Boolean = False)
        procEC = 0
        Dim newProc As New Diagnostics.Process
        ' Por un lado tengo el mandato con zip/unzip.exe xxx" debo separarlo
        Dim principio As String = "" : Dim resto As String = "" : Dim primerblanco As Integer = 0
        primerblanco = InStr(mandato, " ")
        principio = mandato.Substring(0, primerblanco - 1)
        resto = mandato.Substring(primerblanco, Len(mandato) - primerblanco)
        If debuga Then logea(String.Format("dirTra='{0}' mdto='{1}'", dirTra, mandato))


        'info proceso
        newProc.StartInfo.FileName = principio
        newProc.StartInfo.WindowStyle = ProcessWindowStyle.Minimized
        newProc.StartInfo.Arguments = resto
        newProc.StartInfo.RedirectStandardOutput = True
        newProc.StartInfo.UseShellExecute = False
        newProc.StartInfo.RedirectStandardOutput = True
        'If dirTra <> "" Then
        newProc.StartInfo.WorkingDirectory = dirTra
        'End If
        ' newProc.StartInfo.CreateNoWindow = True 'True (Efectivamente lo oculta, pero no se como hacer refresh)

        newProc.Start()
        ' Do not wait for the child process to exit before
        ' reading to the end of its redirected stream.
        ' p.WaitForExit();
        ' Read the output stream first and then wait.
        ' necesito esperar de lo contrario a veces salía de aquí sin haber acabado
        ' http://msdn.microsoft.com/en-us/library/system.diagnostics.process.standardoutput.aspx
        Dim output As String = newProc.StandardOutput.ReadToEnd()
        If debuga Then logea(String.Format("Output -> '{0}'", output))
        newProc.WaitForExit()

    End Sub

    Sub logea(s As String, Optional borra As Boolean = False)
        Dim nl As String = Environment.NewLine
        Dim logea_archivo As String = "Log.txt"
        Try
            If borra Then
                If File.Exists(logea_archivo) Then
                    File.Delete(logea_archivo)
                End If
            Else
                File.AppendAllText(logea_archivo, s & nl)
            End If

        Catch ex As Exception

        End Try
    End Sub

End Module
