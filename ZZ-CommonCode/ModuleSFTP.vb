Imports System.IO
Imports System.Threading.Tasks
Imports Renci.SshNet
Imports Renci.SshNet.Sftp


' como colocar variable públicas
' How would I declare a global variable in Visual Basic?
' https://stackoverflow.com/questions/22738243/how-would-i-declare-a-global-variable-in-visual-basic
Public Module GlobalVariables
    Public GlobalSize As ULong
End Module


Module ModuleSFTP


    Structure strucSFTPUploadFile
        Dim salida As String
        Dim todoOk As Boolean
        Dim host As String
        Dim port As Integer
        Dim username As String
        Dim password As String
        Dim localfile As String
        Dim ftpfile As String
        Public Sub New(host As String, port As Integer, username As String, password As String,
            localfile As String, ftpfile As String)
            Me.host = host : Me.username = username : Me.password = password
            Me.localfile = localfile : Me.ftpfile = ftpfile : Me.port = port
            todoOk = True
        End Sub

    End Structure
    Public Function UploadMyFile(info As strucSFTPUploadFile) As Task(Of strucSFTPUploadFile)
        info.salida = "OK" : info.todoOk = True
        Using filestream As IO.FileStream = New FileStream(info.localfile, FileMode.Open)
            Using sftp As SftpClient = New SftpClient(info.host, info.port, info.username, info.password)
                ' AddressOF es la dirección (el delegado)
                Try
                    sftp.Connect()
                    sftp.UploadFile(filestream, info.ftpfile, AddressOf DisplayPercentage)
                    sftp.Disconnect()

                Catch ex As Exception
                    'Console.WriteLine("Exception " & ex.ToString())
                    info.todoOk = False
                    info.salida = ex.ToString()
                End Try

            End Using
        End Using
        Return Task.FromResult(info)
    End Function




    Structure strucSFTPDownloadFile
        Dim salida As String
        Dim todoOk As Boolean
        Dim host As String
        Dim port As Integer
        Dim username As String
        Dim password As String
        Dim localfile As String
        Dim ftpfile As String
        Public Sub New(host As String, port As Integer, username As String, password As String,
            localfile As String, ftpfile As String)
            Me.host = host : Me.username = username : Me.password = password
            Me.port = port : Me.localfile = localfile : Me.ftpfile = ftpfile
            todoOk = True
        End Sub

    End Structure
    Public Function KKK(info As strucSFTPDownloadFile) As strucSFTPDownloadFile
        info.salida = "OK" : info.todoOk = True
        ' AddressOF es la dirección (el delegado)
        Try
            Using sftp As SftpClient = New SftpClient(info.host, info.port, info.username, info.password)

                sftp.Connect()
                Using filestream As IO.FileStream = New FileStream(info.localfile, IO.FileMode.Create, IO.FileAccess.Write)
                    sftp.DownloadFile(info.ftpfile, filestream, AddressOf DisplayPercentage)
                End Using

                sftp.Disconnect()

            End Using
        Catch ex As Exception
            'Console.WriteLine("Exception " & ex.ToString())
            info.todoOk = False
            info.salida = ex.ToString()
        End Try
        Return info
    End Function
    Public Function DownloadMyFile(info As strucSFTPDownloadFile) As Task(Of strucSFTPDownloadFile)
        info.salida = "OK" : info.todoOk = True
        ' AddressOF es la dirección (el delegado)
        Try
            Using sftp As SftpClient = New SftpClient(info.host, info.port, info.username, info.password)

                sftp.Connect()
                Using filestream As IO.FileStream = New FileStream(info.localfile, IO.FileMode.Create, IO.FileAccess.Write)
                    sftp.DownloadFile(info.ftpfile, filestream, AddressOf DisplayPercentage)
                End Using

                sftp.Disconnect()

            End Using
        Catch ex As Exception
            'Console.WriteLine("Exception " & ex.ToString())
            info.todoOk = False
            info.salida = ex.ToString()
        End Try
        Return Task.FromResult(info)
    End Function
    Public Function DisplayPercentage(u As ULong) As Action(Of ULong)
        ' no es fácil acceder a la IU pq pertenece al thread main
        ' y no vi forma fácil de modificar nada. Así que la trampa ha sido
        ' definiar una variable static global que actualizao con los bytes 
        ' transferidos
        'txtDirPaT.Text = u.ToString >> esto daba error
        ' GlobalSize = u
        GlobalSize = u ' al final conseguí compartir la variable.
        'ActualizaGlobal(u)
        'Dim handle As delegateActualizaGlobal = AddressOf ActualizaGlobal
        'handle(u)
    End Function






    Structure strucSFTPRenameFile
        Dim salida As String
        Dim todoOk As Boolean
        Dim host As String
        Dim port As Integer
        Dim username As String
        Dim password As String
        Dim oldname As String
        Dim newname As String
        Public Sub New(host As String, port As Integer, username As String, password As String,
            oldname As String, newname As String)
            Me.host = host : Me.username = username : Me.password = password : Me.oldname = oldname : Me.newname = newname
            Me.port = port
            todoOk = True
        End Sub
    End Structure
    Public Function RenameMyFile(info As strucSFTPRenameFile) As strucSFTPRenameFile
        info.salida = "OK" : info.todoOk = True
        Using sftp As SftpClient = New SftpClient(info.host, info.port, info.username, info.password)
            ' AddressOF es la dirección (el delegado)
            Try
                sftp.Connect()
                sftp.RenameFile(info.oldname, info.newname)
                sftp.Disconnect()

            Catch ex As Exception
                'Console.WriteLine("Exception " & ex.ToString())
                info.todoOk = False
                info.salida = ex.ToString()
            End Try

        End Using
        Return info
    End Function

    Structure strucSFTPDeleteFile
        Dim salida As String
        Dim todoOk As Boolean
        Dim host As String
        Dim port As Integer
        Dim username As String
        Dim password As String
        Dim file2delete As String
        Public Sub New(host As String, port As Integer, username As String, password As String,
            file2delete As String)
            Me.host = host : Me.username = username : Me.password = password : Me.file2delete = file2delete
            todoOk = True
            Me.port = port
        End Sub
    End Structure
    Public Function DeleteMyFile(info As strucSFTPDeleteFile) As strucSFTPDeleteFile
        info.salida = "OK" : info.todoOk = True
        Using sftp As SftpClient = New SftpClient(info.host, info.port, info.username, info.password)
            ' AddressOF es la dirección (el delegado)
            Try
                sftp.Connect()
                sftp.DeleteFile(info.file2delete)
                sftp.Disconnect()

            Catch ex As Exception
                'Console.WriteLine("Exception " & ex.ToString())
                info.todoOk = False
                info.salida = ex.ToString()
            End Try

        End Using
        Return info
    End Function

    Structure strucinfofile
        Dim filename As String
        Dim filesize As Integer
    End Structure
    Structure strucSFTPSFTPDirDirectory
        Dim salida As String
        Dim todoOk As Boolean
        Dim host As String
        Dim port As Integer
        Dim username As String
        Dim password As String
        Dim directory As String
        Dim filelist As List(Of strucinfofile)
        Public Sub New(host As String, port As Integer, username As String, password As String, directory As String)
            Me.host = host : Me.username = username : Me.password = password : Me.directory = directory
            Me.port = port
            todoOk = True

        End Sub
    End Structure
        Public Function ReadMyDirectory(info As strucSFTPSFTPDirDirectory) As strucSFTPSFTPDirDirectory
        info.salida = "OK" : info.todoOk = True : info.filelist = New List(Of strucinfofile)
        Using sftp As SftpClient = New SftpClient(info.host, info.port, info.username, info.password)
            ' AddressOF es la dirección (el delegado)
            Try
                sftp.Connect()
                Dim Mysftplist As List(Of SftpFile) = sftp.ListDirectory(info.directory)
                sftp.Disconnect()
                For Each Myftpfile In Mysftplist
                    If Myftpfile.IsRegularFile = True Then
                        Dim infofile As New strucinfofile
                        infofile.filename = Myftpfile.Name
                        infofile.filesize = Myftpfile.Attributes.Size
                        info.filelist.Add(infofile)
                    End If
                Next

            Catch ex As Exception
                'Console.WriteLine("Exception " & ex.ToString())
                info.todoOk = False
                info.salida = ex.ToString()
            End Try

        End Using
        Return info
    End Function



End Module
