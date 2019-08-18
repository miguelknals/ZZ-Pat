
Imports System
Imports System.Net
Imports System.Net.Sockets
Imports System.Collections
Imports System.Text

Module UDP
    Structure strucEnvioUDP
        Dim TodoOK As Boolean
        Dim info As String
        Dim Puerto As String
        Dim Host As String
        Dim Proveedor As String
        Dim NombrePat As String
        Dim RecopilarUso As Boolean
        Sub New(RecopilarUso As Boolean, host As String, Puerto As String, Proveedor As String, NombrePat As String)
            Me.TodoOK = True
            Me.RecopilarUso = RecopilarUso
            Me.Host = host
            Me.Puerto = Puerto
            Me.Proveedor = Proveedor
            Me.NombrePat = NombrePat
            Me.info = ""
        End Sub
        Sub envia()
            If Me.RecopilarUso = False Then End ' si no me lo piden
            ' lo inetneto

            ' Este código es una versión simplificada de:
            ' ' este ejemplo es de http://www.winsocketdotnetworkprogramming.com/clientserversocketnetworkcommunication8s.html
            '
            ' Dim puerto As String = inst.Puerto
            ' Dim host As String = inst.Host
            '
            Dim multicastGroups As ArrayList = New ArrayList
            Dim localAddress As IPAddress = IPAddress.Any
            'My.Settings.ParHostUDP = "127.0.0.1"
            'My.Settings.ParPuertoUDP = 50505
            Dim destAddress As IPAddress = IPAddress.Parse(Me.Host)
            Dim portNumber As Long = Me.Puerto
            Dim udpSender As Boolean = False
            Dim bufferSize As Integer = 200 ' ojo importante que tiene que ser el mismo en el cliente
            Dim sendCount As Integer = 1
            Dim i As Integer
            '
            ' la opción que quiero enviar es 
            ' -s hostnaeme - p num_puerto
            udpSender = True ' para enviar

            Dim udpSocket As UdpClient = Nothing

            Dim sendBuffer(bufferSize) As Byte
            Dim receiveBuffer(bufferSize) As Byte

            Dim rc As Integer


            ' laika
            Dim auxs As String = ""
            Dim mensaje As String = Format(Now, "yyyyMMdd HH:MM ")
            mensaje &= Me.Proveedor.Replace(" ", "_") & " "
            auxs = System.Net.Dns.GetHostName()
            mensaje &= auxs & " " & System.Net.Dns.GetHostEntry(auxs).AddressList(0).ToString()
            mensaje &= " " & Me.NombrePat.Replace(" ", "_")
            Dim mensajeB() As Byte
            mensajeB = System.Text.Encoding.UTF8.GetBytes(mensaje)
            Dim posi As Integer = 0
            For posi = 0 To mensajeB.Length - 1
                If posi > bufferSize Then Exit For ' no relleno mas
                sendBuffer(posi) = mensajeB(posi)
            Next

            ' ya lo puedo enviar
            Dim sb As New StringBuilder("")
            Try

                ' Create an unconnected socket since if invoked as a receiver we don't necessarily
                '    want to associate the socket with a single endpoint. Also, for a sender socket
                '    specify local port of zero (to get a random local port) since we aren't receiving
                '    anything.


                sb.AppendLine("Creating a connectionless socket...")

                udpSocket = New UdpClient(New IPEndPoint(localAddress, 0))
                ' Join any multicast groups specified
                sb.AppendLine("Joining any multicast groups specified...")
                ' aquí diría que no entra nunca
                For i = 0 To multicastGroups.Count - 1
                    If (localAddress.AddressFamily = AddressFamily.InterNetwork) Then
                        ' For IPv4 multicasting only the group is specified and not the
                        '    local interface to join it on. This is bad as on a multihomed
                        '    machine, the application won't really know which interface
                        '    it is joined on (and there is no way to change it via the UdpClient).
                        udpSocket.JoinMulticastGroup(CType(multicastGroups(i), IPAddress))
                    ElseIf (localAddress.AddressFamily = AddressFamily.InterNetworkV6) Then
                        ' For some reason, the IPv6 multicast join allows the local interface index
                        '    to be specified such that the application can join multiple groups on
                        '    any interface which is great (but lacking in IPv4 multicasting with the
                        '    UdpClient). IPv6 multicast groups should be specified with the scope id
                        '    when passed on the command line (e.g. fe80::1%4).                       
                        udpSocket.JoinMulticastGroup( _
                            CType(multicastGroups(i), IPAddress).ScopeId, _
                            CType(multicastGroups(i), IPAddress))

                    End If

                Next

                ' If you want to send data with the UdpClient it must be connected -- either by
                '    specifying the destination in the UdpClient constructor or by calling the
                '    Connect method. You can call the Connect method multiple times to associate
                '    a different endpoint with the socket.
                sb.AppendLine("Connecting...")
                udpSocket.Connect(destAddress, portNumber)
                sb.AppendLine("Connect() is OK...")
                ' Send the requested number of packets to the destination
                sb.AppendLine("Sending the requested number of packets to the destination...")
                For i = 0 To sendCount - 1
                    rc = udpSocket.Send(sendBuffer, sendBuffer.Length)
                    sb.AppendLine(String.Format("Sent {0} bytes to {1}", rc, destAddress.ToString()))
                Next
                ' Send a few zero length datagrams to indicate to the receive to exit. Put a short
                '    sleep between them since UDP is unreliable and zero byte datagrams are really
                '    fast (and the local stack can actually drop datagrams before they even make it
                '    onto the wire).
                Console.WriteLine("Do some sleeping, Sleep(250)...")
                For i = 0 To 2
                    rc = udpSocket.Send(sendBuffer, 0)
                    System.Threading.Thread.Sleep(250)
                Next

            Catch err As SocketException
                sb.AppendLine(String.Format("Socket error occurred: {0}", err.Message))
                sb.AppendLine(String.Format("Stack: {0}", err.StackTrace))
                Me.TodoOK = False

            Finally
                If (Not IsNothing(udpSocket)) Then
                    ' Clean things up by dropping any multicast groups we joined
                    sb.AppendLine("Cleaning things up by dropping any multicast groups we joined...")
                    For i = 0 To multicastGroups.Count - 1
                        If (localAddress.AddressFamily = AddressFamily.InterNetwork) Then
                            udpSocket.DropMulticastGroup(CType(multicastGroups(i), IPAddress))
                        ElseIf (localAddress.AddressFamily = AddressFamily.InterNetworkV6) Then
                            ' IPv6 multicast groups should be specified with the scope id when passed
                            '    on the command line
                            udpSocket.DropMulticastGroup( _
                                CType(multicastGroups(i), IPAddress), _
                                (CType(multicastGroups(i), IPAddress)).ScopeId _
                                )
                        End If
                    Next
                    ' Free up the underlying network resources
                    sb.AppendLine("Freeing up the underlying network resources...")
                    udpSocket.Close()
                End If
                Me.info = sb.ToString
            End Try



        End Sub
    End Structure
    
    Sub Envia(NombrePat As String, Proveedor As String) ' lo intento y punto
        ' ojo hay que descomentar esto
       
    End Sub

End Module
