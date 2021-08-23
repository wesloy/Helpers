Imports System.IO
Imports System.Net
Public Class VerificaConexaoInternet


    '1.VerificaConexao_TcpSocket - Verifica a porta 80, padrão para o tráfego http, de um site que esta sempre oline. A resposta é mais rápida que usar o  WebRequest;
    '2. VerificaConexao_WebClient -  Recebe dados de uma URL de um site que esta sempre online;
    '3. VerificaConexao_WebRequest - Envia uma requisição para um site que supõe-se estar sempre online e obtém a resposta;
    '4. VerificaConexao_Win32 - Usa a função  InternetConnectedState da API Win32;

    'As funções retornam TRUE indicando se existe uma conexão ou modem ativo, ou FALSE se não existir conexão ou se todas as conexão não estão ativas.
    Public Shared Function VerificaConexao_TcpSocket() As Boolean
        Try
            Dim cliente As New Sockets.TcpClient("www.google.com", 80)
            cliente.Close()
            Return True
        Catch ex As System.Exception
            CriaArquivoLogErro(ex, "Falha de internet")
            Return False
        End Try
    End Function
    Public Shared Function VerificaConexao_WebClient() As Boolean
        Try
            Using client = New WebClient()
                Using stream = client.OpenRead("http://www.google.com")
                    Return True
                End Using
            End Using
        Catch ex As Exception
            CriaArquivoLogErro(ex, "Falha de internet")
            Return False
        End Try
    End Function
    Public Shared Function VerificaConexao_WebRequest() As Boolean
        Dim objUrl As New System.Uri("http://www.google.com")
        Dim objWebReq As System.Net.WebRequest
        objWebReq = System.Net.WebRequest.Create(objUrl)
        Dim objresp As System.Net.WebResponse
        Try
            objresp = objWebReq.GetResponse
            objresp.Close()
            objresp = Nothing
            Return True
        Catch ex As Exception
            objresp = Nothing
            objWebReq = Nothing
            CriaArquivoLogErro(ex, "Falha de internet")
            Return False
        End Try
    End Function
    Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef _
        lpSFlags As Int32, ByVal dwReserved As Int32) As Boolean
    Private Enum InetConnState
        modem = &H1
        lan = &H2
        proxy = &H4
        ras = &H10
        offline = &H20
        configured = &H40
    End Enum
    Public Shared Function VerificaConexao_Win32() As Boolean
        Dim lngFlags As Long
        If InternetGetConnectedState(lngFlags, 0) Then
            If lngFlags And InetConnState.lan Then
            ElseIf lngFlags And InetConnState.modem Then
            ElseIf lngFlags And InetConnState.configured Then
            ElseIf lngFlags And InetConnState.proxy Then
            ElseIf lngFlags And InetConnState.ras Then
            ElseIf lngFlags And InetConnState.offline Then
            End If
            Return True
        Else
            CriaArquivoLogErro(, "Falha de internet")
            Return False
        End If
    End Function

    Public Shared Sub CriaArquivoLogErro(Optional ByRef e As Exception = Nothing, Optional infoOpcional As String = "")
        Try
            Dim trace = New Diagnostics.StackTrace(e, True)
            Dim line As String = Strings.Right(trace.ToString, 15).Trim
            Dim metodo As String = ""
            Dim texto As String = ""
            For Each sf As StackFrame In trace.GetFrames
                metodo = sf.GetMethod().Name.ToString & " "
            Next
            texto = metodo.ToString & vbNewLine & trace.ToString.Trim
            ' Get the current directory.
            Dim path As String = Directory.GetCurrentDirectory()
            Dim sw As New StreamWriter(path & "\LogErro.txt", True)
            With sw
                .WriteLine("Data: " & DateTime.Now.ToShortDateString())
                .WriteLine("Hora: " & DateTime.Now.ToShortTimeString())
                .WriteLine("Descrição do erro: " & e.Message)
                .WriteLine("Trace: " & texto)
                .WriteLine("Opcional: " & metodo & " [" & infoOpcional & "] ")
                .WriteLine("Computador: " & My.Computer.Name)
                .WriteLine("Usuário: " & My.User.Name)
                .WriteLine("---")
                .Flush()
                .Dispose()
            End With
        Catch
        End Try
    End Sub


End Class
