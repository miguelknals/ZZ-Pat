Imports System
Imports System.Net
Imports System.Net.Sockets
Imports System.Collections
Imports System.IO

Public Class LeeUDP

    Shared Sub main()
        '     ' versión simplificada del ejemplo  la parte de servidir 
        ' este ejemplo es de http://www.winsocketdotnetworkprogramming.com/clientserversocketnetworkcommunication8s.html
        '
        Dim multicastGroups As ArrayList = New ArrayList

        Dim localAddress As IPAddress = IPAddress.Any

        Dim destAddress As IPAddress = Nothing

        Dim portNumber As Long = 50505

        Dim udpSender As Boolean = False

        Dim bufferSize As Integer = 200 ' ojo que debe especificarse en la línea de mandatos o ser igual

        Dim sendCount As Integer = 1

        Dim i As Integer



        ' Parse the command line

        Dim args As String() = Environment.GetCommandLineArgs()



        usage()



        For i = 1 To args.GetUpperBound(0) - 1

            Try

                Dim CurArg() As Char = args(i).ToCharArray(0, args(i).Length)

                If (CurArg(0) = "-") Or (CurArg(0) = "/") Then

                    Select Case Char.ToLower(CurArg(1), System.Globalization.CultureInfo.CurrentCulture)

                        Case "m"        ' Multicast groups to join

                            i = i + 1

                            multicastGroups.Add(IPAddress.Parse(args(i)))

                        Case "l"        ' Local interface to bind to

                            i = i + 1

                            localAddress = IPAddress.Parse(args(i))

                        Case "n"        ' Number of times to send the response

                            i = i + 1

                            sendCount = System.Convert.ToInt32(args(i))

                        Case "p"        ' Port number (local for receiver, remote for sender)

                            i = i + 1

                            portNumber = System.Convert.ToInt32(args(i))

                        Case "r"        ' Indicates UDP receiver

                            udpSender = False

                        Case "s"        ' Indicates UDP sender as well as the destination address

                            udpSender = True

                            i = i + 1

                            destAddress = IPAddress.Parse(args(i))

                        Case "x"        ' Size of the send and receive buffers

                            i = i + 1

                            bufferSize = System.Convert.ToInt32(args(i))

                        Case Else

                            usage()

                            Exit Sub

                    End Select

                End If

            Catch e As Exception

                usage()

                Exit Sub

            End Try

        Next



        Dim udpSocket As UdpClient = Nothing

        Dim sendBuffer(bufferSize) As Byte
        Dim receiveBuffer(bufferSize) As Byte

        Dim rc As Integer



        Try

            ' Create an unconnected socket since if invoked as a receiver we don't necessarily
            '    want to associate the socket with a single endpoint. Also, for a sender socket
            '    specify local port of zero (to get a random local port) since we aren't receiving
            '    anything.

            Console.WriteLine("Creating a connectionless socket...")

            If (udpSender = False) Then
                udpSocket = New UdpClient(New IPEndPoint(localAddress, portNumber))

            Else

                udpSocket = New UdpClient(New IPEndPoint(localAddress, 0))

            End If

            ' Join any multicast groups specified

            Console.WriteLine("Joining any multicast groups specified...")

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

            Console.WriteLine("Connecting...")

            If (udpSender = True) Then

                udpSocket.Connect(destAddress, portNumber)

                Console.WriteLine("Connect() is OK...")

            End If

            If (udpSender = True) Then

                ' Send the requested number of packets to the destination

                Console.WriteLine("Sending the requested number of packets to the destination...")

                For i = 0 To sendCount - 1

                    rc = udpSocket.Send(sendBuffer, sendBuffer.Length)

                    Console.WriteLine("Sent {0} bytes to {1}", rc, destAddress.ToString())

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

            Else

                Dim senderEndPoint As IPEndPoint = New IPEndPoint(localAddress, 0)



                ' Receive datagrams in a loop until a zero byte datagram is received.
                '    Note the difference with the UdpClient in that the source address
                '    is simply an IPEndPoint that doesn't have to be cast to and EndPoint
                '    object as is the case with the Socket class.

                Console.WriteLine("Receiving datagrams in a loop until a zero byte datagram is received...")
                While (True)
                    receiveBuffer = udpSocket.Receive(senderEndPoint)
                    Console.WriteLine("Read {0} bytes from {1}", receiveBuffer.Length, senderEndPoint.ToString())
                    If (receiveBuffer.Length = 0) Then
                        GoTo AfterWhileLoop
                    End If
                    Dim salida As String = System.Text.Encoding.UTF8.GetString(receiveBuffer)
                    Dim ArchivoMF As String = "InfoUsoZZPat.txt"
                    Dim sw As System.IO.StreamWriter
                    sw = My.Computer.FileSystem.OpenTextFileWriter(ArchivoMF, True)
                    sw.WriteLine(salida)
                    sw.Close()
                    Console.WriteLine("Info recibida")
                    Console.WriteLine(salida)

                End While
AfterWhileLoop:
                ' lo impirmo

            End If

        Catch err As SocketException
            Console.WriteLine("Socket error occurred: {0}", err.Message)
            Console.WriteLine("Stack: {0}", err.StackTrace)
        Finally
            If (Not IsNothing(udpSocket)) Then
                ' Clean things up by dropping any multicast groups we joined
                Console.WriteLine("Cleaning things up by dropping any multicast groups we joined...")
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

                Console.WriteLine("Freeing up the underlying network resources...")

                udpSocket.Close()

            End If

        End Try

    End Sub
    Shared Sub usage()

        Console.WriteLine("Usage: UdpClientSampleVB [-m mcast] [-l local-ip] [-n count]")

        Console.WriteLine("Available options:")

        Console.WriteLine()

        Console.WriteLine("                           [-p port] [-r] [-s destination] [-x size]")

        Console.WriteLine("     -m mcast            Multicast group to join. May be specified multiple")

        Console.WriteLine("                            times. For IPv6 include interface to join group on")

        Console.WriteLine("                            as the scope id (e.g. ff12::1%4)")

        Console.WriteLine("     -l local-ip         Local IP address to bind to")

        Console.WriteLine("     -n count            Number of times to send a message")

        Console.WriteLine("     -p port             Port number: local port for receiver, remote port for sender")

        Console.WriteLine("     -r                  Receive UDP data")

        Console.WriteLine("     -s destination      Send UDP data to given destination")

        Console.WriteLine("     -x size             Size of send and receive buffers")

        Console.WriteLine()

    End Sub

End Class
