Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
<Serializable()> Public Structure structParametros
    Dim ParDirTemporal As String
    Dim ParDirSalidaPaT As String
    Dim ParArchivoExcel As String
    'Dim ParSQL As String
    'Dim ParSQL_Alternativo As String
    Dim ParLongInterno As Long
    Dim ParIVA As Single
    Dim ParIRPF As Single
    Dim ParTarifaPredeterminada As Single
    'Dim ParPuertoUDP As Long
    'Dim ParHostUDP As String
    Dim ParProveedor As String
    Dim ParRecopilarUso As Boolean
    Dim ParMasDeUnPB As Boolean
    Dim ParCorreoGestor As String
    Dim ParTipoOrigen As String
    'Dim ParPGHost As String
    'Dim ParPGPuerto As Long
    'Dim ParPGBaseDatos As String
    'Dim ParPGUsuario As String
    'Dim ParPGContrasenya As String
    Dim ParMailSSL As Boolean
    Dim ParMailPuerto As Integer
    Dim ParMailContrase As String
    Dim ParMailHost As String
    Dim ParMailUsuario As String
    Dim TodoOk As Boolean
End Structure
<Serializable()> Public Structure structReferenciasProyectoBase
    Dim ProyectoBase As String
    Dim Referencias As List(Of String)
End Structure

<Serializable()> Public Structure structListaDestinos
    Sub New(nada As String)
        ListaDestino = New List(Of structDestino)
    End Sub
    'Dim Ndestinos As Integer
    Dim ListaDestino As List(Of structDestino)
End Structure
<Serializable()> Public Structure structDestino
    Dim DestNombre As String
    Dim DestTipoDestino As String  ' PATH, SHARED_DRIVE, FTP
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
End Structure



Module General





    Function ObtenArchivoDestinos() As String
        Dim auxS As String = ""
        auxS = System.Reflection.Assembly.GetExecutingAssembly().Location


        ' ahora me interesa la parte ejecutable 
        Dim FullPath As String = Path.GetDirectoryName(auxS)


        Dim archivoDestinos As String = FullPath & "\" & "ZZ_TARGETS.XML" ' & Path.GetFileNameWithoutExtension(auxS) & "_TARGETS.XML"
        Return archivoDestinos

    End Function
    Function ObtenArchivoParametros() As String
        Dim auxS As String = ""
        auxS = System.Reflection.Assembly.GetExecutingAssembly().Location

        Dim args() As String = Environment.GetCommandLineArgs()
        Dim usuario As String = ""
        If args.Count > 1 Then
            usuario = "_" & args(1)
        End If

        ' ahora me interesa la parte ejecutable 
        Dim FullPath As String = Path.GetDirectoryName(auxS)

        Dim archivoValores As String = FullPath & "\" & Path.GetFileNameWithoutExtension(auxS) & "_PAR" & usuario & ".XML"
        Return archivoValores

    End Function
    Function CargaParametros() As structParametros ' devuelve los parámetors
        ' tengo que encontrar el directorio de mi ejecutable
        'Dim archivoParametros As String = ObtenArchivoParametros()
        ' será ejecutable.XML

        ' Convertir las variables de entorno de los path
        Dim archivoParametros As String = ObtenArchivoParametros() ' el archivo ZZ_PAR.XML
        Dim Parametros As New structParametros

        ' será ejecutable.XML
        'Dim ListaParametros As New structParametros
        ' creo valores predeterminados por si no puedo leer

        Dim todoOK As Boolean = False
        If IO.File.Exists(archivoParametros) = True Then ' lo intento leer
            Try
                Dim objStreamReader As New StreamReader(archivoParametros)
                Dim x As New XmlSerializer(Parametros.GetType)
                Parametros = x.Deserialize(objStreamReader)
                objStreamReader.Close()
            Catch ex As Exception
                ' no puedo hacer nada, el archivo está mal
            End Try
            'Deserialize text file to a new object.
        Else
            ' no existe, voy a guardarlo
            ' creo valores predeterminados por si no puedo leer
            ' estos parámetros estaban originalmente en Environment.ExpandEnvironmentVariables(My.Settings.ParDirTemporal)
            Parametros.ParDirTemporal = "%TEMP%\ZZ-Pat\PB" ' Environment.ExpandEnvironmentVariables(My.Settings.ParDirTemporal)
            Parametros.ParDirSalidaPaT = "C:\ZZ-Pat" 'My.Settings.ParDirSalida
            ' Parametros.ParDirTPot = "C:\ZZ-Pat" 'Environment.ExpandEnvironmentVariables(My.Settings.ParDirPB)
            ' No excel Parametros.ParArchivoExcel = My.Settings.ParArchivoEXCEL
            ' No mysql Parametros.ParSQL = My.Settings.ParSQL
            ' No sql Parametros.ParSQL_Alternativo = My.Settings.ParSQL_Alternativo
            Parametros.ParLongInterno = 999999999999999 ' My.Settings.ParSecreto
            Parametros.ParIVA = 21  ' My.Settings.ParIVA
            Parametros.ParIRPF = 15 ' My.Settings.ParIRPF
            Parametros.ParTarifaPredeterminada = 1.234 ' My.Settings.ParTarifaPredeterminada
            'Parametros.ParPuertoUDP = 50505 ' My.Settings.ParPuertoUDP
            'Parametros.ParHostUDP = "213.97.57.158" '  My.Settings.ParHostUDP
            Parametros.ParRecopilarUso = False ' My.Settings.ParRecopilarUso
            Parametros.ParProveedor = "ABC" ' My.Settings.ParProveedor
            Parametros.ParMasDeUnPB = False ' My.Settings.ParMasDeUnPB
            Parametros.ParCorreoGestor = "none@none.com" 'My.Settings.ParCorreoGestor
            Parametros.ParTipoOrigen = "NONE" ' My.Settings.ParTipoOrigen
            ' No postgres Parametros.ParPGHost = My.Settings.ParPGHost
            ' No postgres Parametros.ParPGPuerto = My.Settings.ParPGPuerto
            ' No postgres Parametros.ParPGBaseDatos = My.Settings.ParPGBaseDatos
            ' No postgres Parametros.ParPGUsuario = My.Settings.ParPGUsuario
            ' No postgres Parametros.ParPGContrasenya = My.Settings.ParPGContrasenya
            Parametros.ParMailSSL = True
            Parametros.ParMailPuerto = 587
            Parametros.ParMailContrase = ""
            Parametros.ParMailHost = "smtp.xxxx.com"
            Parametros.ParMailUsuario = "sender@dom.com"
            GuardaParametros(Parametros)
            CargaParametros() ' posible bucle si GuardaParametros falla
        End If

        Return Parametros '

    End Function

    Function CargaDestinos() As structListaDestinos  ' devuelve los destinos
        ' tengo que encontrar el directorio de mi ejecutable
        Dim archivoDestinos As String = ObtenArchivoDestinos() ' el archivo _TARGETS.XML
        ' será ejecutable.XML
        Dim ListaDestinos As New structListaDestinos("")
        ' creo valores predeterminados por si no puedo leer

        Dim todoOK As Boolean = False
        If IO.File.Exists(archivoDestinos) = True Then ' lo intento leer
            Try
                Dim objStreamReader As New StreamReader(archivoDestinos)
                Dim x As New XmlSerializer(ListaDestinos.GetType)
                ListaDestinos = x.Deserialize(objStreamReader)
                objStreamReader.Close()
            Catch ex As Exception
                ' no puedo hacer nada, el archivo está mal
            End Try
            'Deserialize text file to a new object.
        End If

        Return ListaDestinos

    End Function
    Structure infoGuardaDestino
        Dim Destino As structDestino
        Dim todoOk As Boolean
        Dim info As String
        Sub New(nombreDestino As structDestino)
            Destino = nombreDestino
            todoOk = True
            info = ""
        End Sub
    End Structure

    Sub GuardaDestino(infoNuevoDestino As infoGuardaDestino)

        ' lo primero tengo que recuperar los destinos
        Dim listadestinos As structListaDestinos
        listadestinos = CargaDestinos()
        ' si el destino ya existe lo borro
        Dim accion As String = ""
        Dim posi As Integer = 0
        For Each destino In listadestinos.ListaDestino
            If destino.DestNombre = infoNuevoDestino.Destino.DestNombre Then
                listadestinos.ListaDestino.RemoveAt(posi)
                Exit For
            End If
            posi += 1
        Next
        listadestinos.ListaDestino.Add(infoNuevoDestino.Destino)
        ' antes de guardarla, la voy a ordenar
        listadestinos.ListaDestino.Sort(Function(k, y) k.DestNombre.CompareTo(y.DestNombre))

        ' será ejecutable.XML
        GuardaArchivoDestinos(listadestinos)

        
    End Sub
    Sub GuardaArchivoDestinos(destinos As structListaDestinos)

        ' ' tengo que encontrar el directorio de mi ejecutable
        Dim archivoDestinos As String = ObtenArchivoDestinos()

        ' será ejecutable.XML
        Dim todoOK As Boolean = False

        If IO.File.Exists(archivoDestinos) = True Then ' lo intento leer
            ' lo borro
            ' miro si existe el bak
            If IO.File.Exists(archivoDestinos & ".BAK") = True Then
                IO.File.Delete(archivoDestinos & ".BAK")
            End If
            IO.File.Move(archivoDestinos, archivoDestinos & ".BAK") ' renombro por si hay eerror
        End If
        Try
            Dim ObjSW As New StreamWriter(archivoDestinos) ' lo guardaré en mi ejecutuable
            Dim x As New XmlSerializer(destinos.GetType) ' serializo mi estructura
            x.Serialize(ObjSW, destinos) ' guardo el par
            ObjSW.Close()
        Catch ex As Exception
            MsgBox("Error updating Target file" & archivoDestinos & ". We will try to restore de original file after you press OK", MsgBoxStyle.Exclamation)
            If IO.File.Exists(archivoDestinos) = True Then ' existe y lo borry
                IO.File.Delete(archivoDestinos)
            End If
            If IO.File.Exists(archivoDestinos & ".BAK") = True Then ' en teoría siempre lo tendré
                IO.File.Move(archivoDestinos & ".BAK", archivoDestinos) ' restauro el bak
            End If
        Finally
            ' si existe el bak, lo borro
            If IO.File.Exists(archivoDestinos & ".BAK") = True Then ' 
                IO.File.Delete(archivoDestinos & ".BAK") ' borro el bak ya no lo necesito
            End If
        End Try

    End Sub

    Sub GuardaParametros(Parametros As structParametros)

        ' ' tengo que encontrar el directorio de mi ejecutable
        Dim archivoParametros As String = ObtenArchivoParametros()

        ' será ejecutable.XML
        Dim todoOK As Boolean = False

        If IO.File.Exists(archivoParametros) = True Then ' lo intento leer
            ' lo borro
            ' miro si existe el bak
            If IO.File.Exists(archivoParametros & ".BAK") = True Then
                IO.File.Delete(archivoParametros & ".BAK")
            End If
            IO.File.Move(archivoParametros, archivoParametros & ".BAK") ' renombro por si hay eerror
        End If
        Try
            Dim ObjSW As New StreamWriter(archivoParametros) ' lo guardaré en mi ejecutuable
            Dim x As New XmlSerializer(Parametros.GetType) ' serializo mi estructura
            x.Serialize(ObjSW, Parametros) ' guardo el par
            ObjSW.Close()
            todoOK = True
        Catch ex As Exception
            MsgBox("Error updating Parameter file" & archivoParametros & ". We will try to restore de original file after you press OK", MsgBoxStyle.Exclamation)
            If IO.File.Exists(archivoParametros) = True Then ' existe y lo borry
                IO.File.Delete(archivoParametros)
            End If
            If IO.File.Exists(archivoParametros & ".BAK") = True Then ' en teoría siempre lo tendré
                IO.File.Move(archivoParametros & ".BAK", archivoParametros) ' restauro el bak
            End If
        Finally
            ' si existe el bak, lo borro
            If IO.File.Exists(archivoParametros & ".BAK") = True Then ' 
                IO.File.Delete(archivoParametros & ".BAK") ' borro el bak ya no lo necesito
            End If
        End Try


    End Sub


    Structure strucProIdiomaCarpeta
        Sub New(carpeta As String)
            Me.carpeta = carpeta : Me.LANG = "N/D" : Me.IU = "N/D"
            Dim auxS As String = "" : Dim c As Char = "" : Dim encontrado As Boolean = False
            For i = Len(Me.carpeta) To 1 Step -1
                c = Me.carpeta.Substring(i - 1, 1)
                If c = "_" Then encontrado = True : Exit For
                auxS = c & auxS
            Next
            If encontrado Then
                LANG = auxS
                IU = carpeta.Replace("_" & LANG, "")
            Else
                LANG = ""
                IU = carpeta
            End If


        End Sub
        Dim carpeta As String
        Dim LANG As String
        Dim IU As String
    End Structure


End Module
