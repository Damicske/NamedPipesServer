Imports System.ComponentModel
Imports System.IO
Imports System.IO.Pipes
Imports System.Text
Imports System.Threading

Public Class Form1
    Private Shared NumThreads As Integer = 4 'if you only want 1 server you can adjust the code so you have less loops.
    Private ClosingForm As Boolean = False
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim server As New Thread(AddressOf ServerLoop)
        server.Start()
    End Sub

    Private Sub ServerLoop()
        Dim i As Integer
        Dim Servers(NumThreads) As Thread

        Console.WriteLine("*** Named pipe server stream with impersonation example ***")
        While Not ClosingForm
            Console.WriteLine("Waiting for client connect...")
            For i = 0 To NumThreads - 1
                servers(i) = New Thread(AddressOf ServerThread)
                servers(i).Start()
            Next i
            Thread.Sleep(250)
            While i > 0
                If ClosingForm Then
                    For j As Integer = 0 To NumThreads - 1
                        servers(j).Abort()
                        Console.WriteLine("Server thread[{0}] finished.", servers(j).ManagedThreadId)
                        servers(j) = Nothing
                    Next j
                    Exit While
                End If
                For j As Integer = 0 To NumThreads - 1
                    If Not (servers(j) Is Nothing) AndAlso servers(j).Join(250) Then
                        Console.WriteLine("Server thread[{0}] finished.", servers(j).ManagedThreadId)
                        servers(j) = Nothing
                        i -= 1    ' decrement the thread watch count
                    End If
                Next j
            End While
        End While
        Console.WriteLine("ServerLoop exiting.")
    End Sub

    Private Sub ServerThread(data As Object)
        Using pipeServer As New NamedPipeServerStream("testpipe", PipeDirection.InOut, NumThreads)

            Dim threadId As Integer = Thread.CurrentThread.ManagedThreadId

            ' Wait for a client to connect
            pipeServer.WaitForConnectionAsync()
            While Not pipeServer.IsConnected
                Thread.Sleep(100)
                If ClosingForm Then Exit Sub
            End While

            Console.WriteLine("Client connected on thread[{0}].", threadId)
            Try
                ' Read the request from the client. Once the client has
                ' written to the pipe its security token will be available.

                Dim ss As New StreamString(pipeServer)
                While pipeServer.IsConnected
                    ' Verify our identity to the connected client using a
                    ' string that the client anticipates.

                    ss.WriteString("I am the one true server!")
                    'Dim filename As String = ss.ReadString()
                    If ss.ReadString = "getdata" Then
                        ss.WriteString("Dit is wat data, get used to it :+")
                    End If
                End While
                ' Read in the contents of the file while impersonating the client.
                ' Dim fileReader As New ReadFileToStream(ss, filename)

                ' Display the name of the user we are impersonating.
                'Console.WriteLine("Reading file: {0} on thread[{1}] as user: {2}.", filename, threadId, pipeServer.GetImpersonationUserName())
                'pipeServer.RunAsClient(AddressOf fileReader.Start)
                ' Catch the IOException that is raised if the pipe is broken
                ' or disconnected.
            Catch e As IOException
                Console.WriteLine("ServerThread ERROR: {0}", e.Message)
            End Try
            pipeServer.Close()
        End Using
    End Sub

    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        ClosingForm = True
    End Sub

    Private Sub Form1_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        ' End
    End Sub
End Class


''' <summary>
''' Defines the data protocol for reading and writing strings on our stream
''' </summary>
Public Class StreamString
    Private ReadOnly IoStream As Stream
    Private ReadOnly StreamEncoding As UnicodeEncoding

    Public Sub New(ioStream As Stream)
        Me.IoStream = ioStream
        StreamEncoding = New UnicodeEncoding(False, False)
    End Sub

    ''' <summary>
    ''' Reads the stream for a string
    ''' </summary>
    Public Function ReadString() As String
        Try
            Dim Len As Integer = CType(IoStream.ReadByte(), Integer) * 256
            Len += CType(IoStream.ReadByte(), Integer)
            Dim inBuffer As Array = Array.CreateInstance(GetType(Byte), Len)
            IoStream.Read(inBuffer, 0, Len)
            Return StreamEncoding.GetString(CType(inBuffer, Byte()))
        Catch ex As Exception
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Writes the outString to the given stream
    ''' </summary>
    ''' <param name="outString">String to write to stream</param>
    ''' <returns>Buffer length</returns>
    Public Function WriteString(outString As String) As Integer
        Dim outBuffer As Byte() = StreamEncoding.GetBytes(outString)
        Dim len As Integer = outBuffer.Length
        If len > UShort.MaxValue Then len = CType(UShort.MaxValue, Integer)

        IoStream.WriteByte(CType(len \ 256, Byte))
        IoStream.WriteByte(CType(len And 255, Byte))
        IoStream.Write(outBuffer, 0, outBuffer.Length)
        IoStream.Flush()

        Return outBuffer.Length + 2
    End Function
End Class

''' <summary> 
''' Contains the method executed in the context of the impersonated user
''' </summary>
Public Class ReadFileToStream
    Private Fn As String
    Private Ss As StreamString

    Public Sub New(str As StreamString, filename As String)
        fn = filename
        ss = str
    End Sub

    Public Sub Start()
        Ss.WriteString(File.ReadAllText(Fn))
    End Sub
End Class

