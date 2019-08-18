Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
<Serializable()> Public Structure structParametros
    Sub New(nada As String)
        DestNombre = "your_name"
        DestTipoDestino = "SHARED_DRIVE"
        DestCorreo = "your_email@dot.com"
        DestSharedDriveHost = ""
        DestSharedDriveNombre = "\\hostname\exchange_area\mcanals_SHARED"
        DestSharedDriveUsuario = ""
        DestSharedDriveContrasenya = ""
        DestFTPHost = "ftp.dom.com"
        DestFTPNombre = "\ftpdirectory"
        DestFTPUsuario = "your_ftp_user"
        DestFTPContrasenya = ""
        DestLocalPatDir = "c:\ZZ-Pat-Client"
        DestLongInterno = 12345123451234

    End Sub
    Dim DestNombre As String
    Dim DestTipoDestino As String  'SHARED_DRIVE, FTP
    Dim DestCorreo As String
    Dim DestPath As String
    Dim DestSharedDriveHost As String
    Dim DestSharedDriveNombre As String
    Dim DestSharedDriveUsuario As String
    Dim DestSharedDriveContrasenya As String
    Dim DestFTPHost As String
    Dim DestFTPNombre As String
    Dim DestFTPUsuario As String
    Dim DestFTPContrasenya As String
    Dim DestLongInterno As Long
    '
    Dim DestLocalPatDir As String
End Structure


Module General

    Function CargaParametros() As structParametros ' devuelve los parámetors
        ' tengo que encontrar el directorio de mi ejecutable
        Dim archivoParametros As String = ObtenArchivoParametros()
        ' será ejecutable.XML
        Dim Parametros As New structParametros("") ' para inicializar las variables en caso no puedo leer.
        ' creo valores predeterminados por si no puedo leer

        Dim todoOK As Boolean = False
        If IO.File.Exists(archivoParametros) = True Then ' lo intento leer
            Try
                Dim objStreamReader As New StreamReader(archivoParametros)
                ' esta sentencia puede dar errores en el debugger
                ' que en su momento entdí que se podían ignorar y que aparencen
                ' al pulstar Control+Alt+E Hay que deseleccionar todo.
                ' el problema, que a veces algunos errores en el debugger no los ves
                ' http://stackoverflow.com/questions/294659/why-did-i-get-an-error-with-my-xmlserializer
                ' y
                ' https://social.msdn.microsoft.com/Forums/en-US/9f0c169f-c45e-4898-b2c4-f72c816d4b55/strange-xmlserializer-error?forum=asmxandxml
                ' aquí algo más
                ' http://www.iteramos.com/pregunta/4414/dando-filenotfoundexception-al-constructor-xmlserializer
                ' Hay que deseleccionar:
                ' - CLR exceptions
                ' - Mannaging Debugging Assitants
                Dim x As New XmlSerializer(Parametros.GetType)
                Parametros = x.Deserialize(objStreamReader)
                objStreamReader.Close()
                todoOK = True
            Catch ex As Exception
                ' no puedo hacer nada, el archivo está mal
            End Try
            'Deserialize text file to a new object.
        End If
        Return Parametros '
    End Function

    Sub GuardaParametros(Par As structParametros)

        ' tengo que encontrar el directorio de mi ejecutable
        Dim archivoParametros As String = ObtenArchivoParametros()

        ' será ejecutable.XML
        Dim todoOK As Boolean = False

        If IO.File.Exists(archivoParametros) = True Then ' lo intento leer
            ' lo borro
            IO.File.Delete(archivoParametros)
        End If
        Dim ObjSW As New StreamWriter(archivoParametros) ' lo guardaré en mi ejecutuable
        Dim x As New XmlSerializer(Par.GetType) ' serializo mi estructura
        x.Serialize(ObjSW, Par) ' guardo el par
        ObjSW.Close()

    End Sub
    Function ObtenArchivoParametros() As String
        Dim auxS As String = ""
        auxS = System.Reflection.Assembly.GetExecutingAssembly().Location
        ' ahora me interesa la parte ejecutable 
        Dim FullPath As String = Path.GetDirectoryName(auxS)
        Dim archivoValores As String = FullPath & "\" & Path.GetFileNameWithoutExtension(auxS) & "_PAR.XML"
        Return archivoValores

    End Function


    ' Ojo con esta clase, se basa copiando la que hay en la última verisón de z-pat
    ' es como la que hay en z-pat, PERO sin menos procedimientos, por eso no está en el

End Module

'´módulo común

