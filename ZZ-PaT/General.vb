Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
<Serializable()> Public Structure structParametros
    Dim ParDirTemporal As String
    Dim ParDirSalidaPaT As String
    Dim ParDirTPot As String
    Dim ParArchivoExcel As String
    Dim ParSQL As String
    Dim ParSQL_Alternativo As String
    Dim ParSecretoServidor As String
    Dim ParIVA As Single
    Dim ParIRPF As Single
    Dim ParTarifaPredeterminada As Single
    Dim ParPuertoUDP As Long
    Dim ParHostUDP As String
    Dim ParProveedor As String
    Dim ParRecopilarUso As Boolean
    Dim ParMasDeUnPB As Boolean
    Dim ParCorreoGestor As String
    Dim ParTipoOrigen As String
    Dim ParPGHost As String
    Dim ParPGPuerto As Long
    Dim ParPGBaseDatos As String
    Dim ParPGUsuario As String
    Dim ParPGContrasenya As String
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


        Dim archivoDestinos As String = FullPath & "\" & Path.GetFileNameWithoutExtension(auxS) & "_TARGETS.XML"
        Return archivoDestinos

    End Function
    Function ObtenArchivoParametros() As String
        Dim auxS As String = ""
        auxS = System.Reflection.Assembly.GetExecutingAssembly().Location


        ' ahora me interesa la parte ejecutable 
        Dim FullPath As String = Path.GetDirectoryName(auxS)


        Dim archivoValores As String = FullPath & "\" & "ZZ_PAR.XML" ' Path.GetFileNameWithoutExtension(auxS) & "_PAR.XML"
        Return archivoValores

    End Function
    Function CargaParametros() As structParametros ' devuelve los parámetors
        ' tengo que encontrar el directorio de mi ejecutable
        'Dim archivoParametros As String = ObtenArchivoParametros()
        ' será ejecutable.XML
        Dim Parametros As structParametros
        ' creo valores predeterminados por si no puedo leer
        Parametros.ParDirTemporal = Environment.ExpandEnvironmentVariables(My.Settings.ParDirTemporal)
        Parametros.ParDirSalidaPaT = My.Settings.ParDirSalida
        Parametros.ParDirTPot = Environment.ExpandEnvironmentVariables(My.Settings.ParDirPB)
        Parametros.ParArchivoExcel = My.Settings.ParArchivoEXCEL
        Parametros.ParSQL = My.Settings.ParSQL
        Parametros.ParSQL_Alternativo = My.Settings.ParSQL_Alternativo
        Parametros.ParSecretoServidor = My.Settings.ParSecreto
        Parametros.ParIVA = My.Settings.ParIVA
        Parametros.ParIRPF = My.Settings.ParIRPF
        Parametros.ParTarifaPredeterminada = My.Settings.ParTarifaPredeterminada
        Parametros.ParPuertoUDP = My.Settings.ParPuertoUDP
        Parametros.ParHostUDP = My.Settings.ParHostUDP
        Parametros.ParRecopilarUso = My.Settings.ParRecopilarUso
        Parametros.ParProveedor = My.Settings.ParProveedor
        Parametros.ParMasDeUnPB = My.Settings.ParMasDeUnPB
        Parametros.ParCorreoGestor = My.Settings.ParCorreoGestor
        Parametros.ParTipoOrigen = My.Settings.ParTipoOrigen
        Parametros.ParPGHost = My.Settings.ParPGHost
        Parametros.ParPGPuerto = My.Settings.ParPGPuerto
        Parametros.ParPGBaseDatos = My.Settings.ParPGBaseDatos
        Parametros.ParPGUsuario = My.Settings.ParPGUsuario
        Parametros.ParPGContrasenya = My.Settings.ParPGContrasenya

        ' Convertir las variables de entorno de los path

        
        Return Parametros '

    End Function
    Function CargaParametrosOLD() As structParametros ' devuelve los parámetors
        ' tengo que encontrar el directorio de mi ejecutable
        Dim archivoParametros As String = ObtenArchivoParametros()
        ' será ejecutable.XML
        Dim Parametros As structParametros
        ' creo valores predeterminados por si no puedo leer
        Parametros.ParDirTemporal = ""
        Parametros.ParDirSalidaPaT = ""
        Parametros.ParDirTPot = ""
        Parametros.ParArchivoExcel = ""
        Parametros.ParSQL = ""
        Parametros.ParSecretoServidor = ""
        Parametros.ParIVA = 0
        Parametros.ParIRPF = 0

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

    Sub GuardaParametros(par As structParametros)
        With My.Settings
            .ParDirSalida = par.ParDirSalidaPaT
            .ParDirPB = par.ParDirTPot
            .ParDirTemporal = par.ParDirTemporal
            .ParArchivoEXCEL = par.ParArchivoExcel
            .ParSQL = par.ParSQL
            .ParSQL_Alternativo = par.ParSQL_Alternativo
            .ParSecreto = par.ParSecretoServidor
            .ParIVA = par.ParIVA
            .ParIRPF = par.ParIRPF
            .ParTarifaPredeterminada = par.ParTarifaPredeterminada
            .ParPuertoUDP = par.ParPuertoUDP
            .ParHostUDP = par.ParHostUDP
            .ParRecopilarUso = par.ParRecopilarUso
            .ParProveedor = par.ParProveedor
            .ParMasDeUnPB = par.ParMasDeUnPB
            .ParCorreoGestor = par.ParCorreoGestor
            .ParPGHost = par.ParPGHost
            .ParPGPuerto = par.ParPGPuerto
            .ParPGBaseDatos = par.ParPGBaseDatos
            .ParPGUsuario = par.ParPGUsuario
            .ParPGContrasenya = par.ParPGContrasenya
            .ParTipoOrigen = par.ParTipoOrigen
            .Save()
        End With

    End Sub

    Sub GuardaParametrosOLD(Par As structParametros)

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
    Structure strucInterpretaSQL
        Sub New(carpeta As String, mascaraSQL As String)
            Me.carpeta = carpeta : Me.mascaraSQL = mascaraSQL : Me.TodoOK = True
            Me.info = "" : Me.sql = "" : Me.LANG = "" : Me.IU = ""
        End Sub
        Dim carpeta As String
        Dim mascaraSQL As String
        Dim info As String
        Dim sql As String
        Dim TodoOK As Boolean
        Dim LANG As String
        Dim IU As String
        Sub interpreta()
            'PI314ABD001_SPA
            '{FULL} -> PI314ABD001_SPA
            '{IU} -> PI314ABD001
            '{LANG} -> SPA
            '{SUB 0 5} -> PI314
            '{SUB 5 5} -> ABD001
            Me.sql = Me.mascaraSQL
            Me.sql = Me.sql.Replace("{FULL}", Me.carpeta)

            If InStr(Me.sql, "{IU}") Then


                Dim posiG As Integer = 0
                posiG = InStr(Me.carpeta, "_")
                If posiG = 0 Or posiG = 1 Then
                    IU = Me.carpeta ' no hay guion o está en primera posición
                Else
                    IU = Me.carpeta.Substring(0, posiG - 1)

                End If
                Me.sql = Me.sql.Replace("{IU}", IU)


            End If
            '            If InStr(Me.sql, "{LANG}") Then
            '                Dim posiG As Integer = 0
            '                posiG = InStr(Me.carpeta, "_")
            '                If posiG = 0 Or posiG = Len(Me.carpeta) Then
            '                    LANG = "" ' no hay guion o al final
            '                Else
            '                    LANG = Me.carpeta.Substring(posiG, Len(Me.carpeta) - posiG)
            '                End If
            '                Me.sql = Me.sql.Replace("{LANG}", LANG)
            '            End If
            If InStr(Me.sql, "{LANG}") Then
                Dim auxS As String = "" : Dim c As Char = "" : Dim encontrado As Boolean = False
                For i = Len(Me.carpeta) To 1 Step -1
                    c = Me.carpeta.Substring(i - 1, 1)
                    If c = "_" Then encontrado = True : Exit For
                    auxS = c & auxS
                Next
                If encontrado Then
                    LANG = auxS
                Else
                    LANG = ""
                End If
            End If

            Try
                While InStr(Me.sql, "{SUB") <> 0
                    ' hay un sub
                    Dim posiIniSub = InStr(Me.sql, "{SUB") - 1
                    Dim i As Integer = posiIniSub
                    Dim auxS As String = "" : Dim c As Char = ""
                    For i = posiIniSub To Len(Me.sql) - 1
                        c = Me.sql.Substring(i, 1)
                        auxS &= c
                        If c = "}" Then ' fin
                            ' aux S deberia ser algo como SUB 1 3
                            Dim pal As String() = auxS.Split(" ")
                            Dim iniSub As Integer = Val(pal(1)) - 1
                            Dim lenSub As Integer = Val(pal(2))
                            Dim SubS As String = Me.carpeta.Substring(iniSub, lenSub)
                            Me.sql = Replace(Me.sql, auxS, SubS)
                            Exit For
                        End If
                    Next
                End While

            Catch ex As Exception
                Me.TodoOK = False
                Me.info = ex.Message.ToString
                Me.sql = ""
            End Try


        End Sub
    End Structure

   

End Module
