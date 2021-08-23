Imports System.Data.OleDb
Imports ADODB
Public Class Conexao
    Private conexao As New ADODB.Connection
    Private cmd As New ADODB.Command
    Public password_bd As String = ""
    Public diretorio_bd As String = ""
    Public banco_dados As String = ""
    Public servidor_bd As String = ""
    Public user_bd As String = ""
    Public SGBD As Integer
    Private sql As String = ""
    Public Const TITULO_ALERTA = "Alerta do Sistema"
    Public Enum FLAG_SGBD
        SQL = 1
        ACESS = 2
    End Enum

    Public Sub New(ByVal _SGBD As FLAG_SGBD, ByVal _password_bd As String, ByVal _banco_dados As String, ByVal _servidor_bd As String, ByVal _user_bd As String, ByVal _diretorio_bd As String)
        SGBD = _SGBD
        password_bd = _password_bd
        banco_dados = _banco_dados
        servidor_bd = _servidor_bd
        user_bd = _user_bd
        diretorio_bd = _diretorio_bd
    End Sub

    Public Function getStringConexao() As String
        Dim strconexao As String = ""

        If SGBD = 2 Then
            'String de conexao com banco de dados MSACCESS com senha
            strconexao = "Provider=Microsoft.Jet.OLEDB.4.0;"
            strconexao += "Data Source=" & diretorio_bd & banco_dados & ";"
            strconexao += "Jet OLEDB:Database Password= " & password_bd & ";"
            strconexao += "Persist Security Info=False"
        Else
            'SQL Server , usando SQL Server OLE DB Provider
            strconexao = "Provider=SQLOLEDB;"
            strconexao += "Data Source=" & servidor_bd & ";"
            strconexao += "Initial Catalog=" & banco_dados & ";"
            strconexao += "User Id=" & user_bd & ";"
            strconexao += "Password=" & password_bd & ";Connect Timeout=5000"
        End If

        Return strconexao
    End Function

    Public Function conectar() As Boolean
        Dim bln As Boolean = False
        Try
            If conexao.State = ConnectionState.Closed Then
                With conexao
                    .ConnectionString = getStringConexao()
                    .Mode = ConnectModeEnum.adModeReadWrite 'modo de conexao leitura e escrita
                    .Open()
                End With
            End If
            bln = True
        Catch ex As Exception
            desconectar()
            'MsgBox("Erro ao efetuar a conexão com a base de dados." & vbNewLine & ex.Message, vbCritical, TITULO_ALERTA)
            bln = False
        End Try
        Return bln
    End Function

    Public Sub desconectar()
        Try
            If Not conexao Is Nothing Then
                If Not conexao.State = ConnectionState.Closed Then
                    conexao.Close()
                End If
            End If
        Catch ex As Exception
            'MsgBox("Erro ao desconectar com a base de dados : " & ex.Message, vbCritical, TITULO_ALERTA)
        End Try
    End Sub

    Public Sub testaConexao()
        Try
            conectar()
            MsgBox("Conexão realizada com sucesso!!!")
        Catch ex As Exception
            'MsgBox("Não foi possível conectar ao banco de dados." & ex.Message, vbCritical, TITULO_ALERTA)
            conexao = Nothing
        End Try
        desconectar()
    End Sub

    'Executa um comando SQL e retorna um boleano
    Public Function executaQuery(ByVal strSql As String, Optional ByRef qtRegistroBlock As Long = 0) As Boolean
        Try
            'verifica se a conexao esta fechada
            If conexao.State = ConnectionState.Closed Then
                conectar()
            End If
            'executa a consulta
            With conexao
                .Execute(strSql, qtRegistroBlock)
            End With

            'retorna verdadeiro
            executaQuery = True
            desconectar()

        Catch ex As Exception
            'MsgBox("Erro de conexão(" & Err.Number & "): " & vbNewLine & "Não foi possível estabelecer uma conexão com o banco de dados. " & vbNewLine & strSql &
            '"Por favor, tente novamente.", vbCritical, TITULO_ALERTA)
            MsgBox("Erro de conexão(" & Err.Number & ")" & vbNewLine & "Não foi possível estabelecer uma conexão com o banco de dados. ", vbCritical, TITULO_ALERTA)
            Return False
            Exit Function
        End Try
    End Function

    Public Function retornaDataTable(ByVal strSQL As String) As DataTable
        Dim objDA As New OleDbDataAdapter
        Dim objDT As New DataTable
        Dim rsObjt As ADODB.Recordset
        Try
            If conexao.State = ConnectionState.Closed Then
                conectar()
            End If
            rsObjt = retornaRs(strSQL)
            objDT = recordSetToDataTable(rsObjt)
            desconectar()
        Catch ex As Exception
            desconectar()
            'MsgBox("Ocorreu um Erro: " & Err.Number & " " & ex.Message, vbCritical, TITULO_ALERTA)
        End Try
        retornaDataTable = objDT
    End Function

    'Retorna um recordset
    Public Function retornaRs(ByVal strSQL As String) As ADODB.Recordset
        Dim ADORecordset As New ADODB.Recordset
        Try
            If conexao.State = ConnectionState.Closed Then
                conectar()
            End If
            With ADORecordset
                .CursorLocation = CursorLocationEnum.adUseClient
            End With
            ADORecordset.Open(strSQL, conexao, CursorTypeEnum.adOpenKeyset, ADODB.LockTypeEnum.adLockOptimistic)
            retornaRs = ADORecordset
            ADORecordset = Nothing
        Catch ex As Exception
            'MsgBox("Ocorreu um Erro: " & Err.Number & " " & ex.Message, vbCritical, TITULO_ALERTA)
            Return Nothing
        End Try
    End Function

    Public Function recordSetToDataTable(ByVal objRS As ADODB.Recordset) As DataTable
        Dim objDA As New OleDbDataAdapter()
        Dim objDT As New DataTable()
        objDA.Fill(objDT, objRS)
        desconectar()
        Return objDT
    End Function

    Public Function logicoSql(ByVal argValor As Boolean, Optional ByVal sql As Boolean = True) As String
        'Função que troca os valores lógicos Verdadeiro/Falso
        'para True/False para utilização em consultas SQL
        'Se o valor for verdadeiro
        If argValor Then
            'Troca por True
            logicoSql = IIf(sql, "1", "True")
            'logicoSql = "True"
        Else
            'Senão troca por False
            logicoSql = IIf(sql, "0", "False")
            'logicoSql = "False"
        End If
    End Function

    Public Function pontoVirgula(ByVal varValor As Object) As String
        'Função que troca a vírgula de um valor decimal por
        'um ponto para utilização em consultas SQL

        Dim strValor As String
        Dim strInteiro As String
        Dim strDecimal As String
        Dim intPosicao As Integer

        'Converte o valor em string
        strValor = CStr(varValor)

        'Busca a posição da vírgula
        intPosicao = InStr(strValor, ",")

        'Se há uma vírgula em alguma posição
        If intPosicao > 0 Then
            'Retira a parte inteira
            strInteiro = Left(strValor, intPosicao - 1)
            'Retira a parte decimal
            strDecimal = Right(strValor, Len(strValor) - intPosicao)
            'Junta os dois novamente incluindo
            'agora o ponto no lugar da vírgula
            pontoVirgula = strInteiro & "." & strDecimal
        Else
            'Senão devolve o mesmo valor
            pontoVirgula = strValor
        End If

    End Function

    Public Function HoraSql(ByVal argData As DateTime, Optional ByVal sql As Boolean = True) As String
        'Função que formata uma data para o modo SQL
        'com a cerquilha: #YYYY-MM-DD HH:MM:SS#
        'sempre retorna uma string
        Dim strDataCompleta As String
        'Remonta no formato adequado (Padrão banco de dados)
        strDataCompleta = CDate(argData).ToString("HH:mm:ss")
        'HoraSql = "#" & strDataCompleta & "#"
        HoraSql = IIf(sql, "'" & strDataCompleta & "'", "#" & strDataCompleta & "#")
    End Function

    Public Function dataSql(ByVal argData As DateTime, Optional ByVal sql As Boolean = True) As String
        'Função que formata uma data para o modo SQL
        'com a cerquilha: #YYYY-MM-DD HH:MM:SS#
        'sempre retorna uma string
        Dim strDataCompleta As String
        'Remonta no formato adequado (Padrão banco de dados)
        strDataCompleta = CDate(argData).ToString("yyyy-MM-dd HH:mm:ss")
        'dataSql = "#" & strDataCompleta & "#"
        dataSql = IIf(sql, "'" & strDataCompleta & "'", "#" & strDataCompleta & "#")
    End Function

    Public Function dataSqlAbreviada(ByVal argData As DateTime, Optional ByVal sql As Boolean = True) As String
        'Função que formata uma data para o modo SQL
        'com a cerquilha: #YYYY-MM-DD HH:MM:SS#
        'sempre retorna uma string
        Dim strDataCompleta As String
        'Remonta no formato adequado (Padrão banco de dados)
        strDataCompleta = CDate(argData).ToString("yyyy-MM-dd")
        dataSqlAbreviada = IIf(sql, "'" & strDataCompleta & "'", "#" & strDataCompleta & "#")
    End Function

    Public Function valorSql(ByVal argValor As Object, Optional ByVal sql As Boolean = True) As String
        'Função que formata valores para utilização
        'em consultas SQL
        valorSql = Nothing

        If argValor = Nothing Then
            valorSql = "Null"
        End If
        'Seleciona o tipo de valor informado
        Select Case VarType(argValor)
            'Caso seja vazio ou nulo apenas
            'devolve a string Null
            Case vbEmpty, vbNull
                valorSql = "Null"
                'Caso seja inteiro ou longo apenas
                'converte em string
            Case vbInteger, vbLong
                valorSql = CStr(argValor)
                'Caso seja simples, duplo, decimal ou moeda
                'substitui a vírgula por ponto
            Case vbSingle, vbDouble, vbDecimal, vbCurrency
                valorSql = pontoVirgula(argValor)
                'Caso seja data chama a função dataSql()
            Case vbDate
                'verifica se esta vazio e retorna Null
                'Or argValor = "00:00:00" Or argValor = "12:00:00 AM"
                Dim dataVazia As DateTime = Nothing
                If CDate(argValor).ToString("yyyy-MM-dd HH:mm:ss") = CDate(dataVazia).ToString("yyyy-MM-dd HH:mm:ss") Then
                    valorSql = "Null"
                Else
                    valorSql = dataSql(argValor, sql)
                End If
                'Caso seja string acrescenta aspas simples
            Case vbString
                If String.IsNullOrEmpty(argValor) Or argValor = "" Then
                    'devolve a string Null
                    valorSql = "Null"
                Else
                    'acrescenta aspas simples para valores diferentes de vazio
                    valorSql = "'" & argValor & "'"
                End If
                'Caso seja lógico chama a função logicoSql()
            Case vbBoolean
                valorSql = logicoSql(argValor, sql)
        End Select
        Return valorSql
    End Function

    ''' <summary>
    '''Função para retornar um valor vazio ao invés de nulo.
    '''para utilização nas classes DTO
    '''Para setar campo data como null/nothing:
    '''campoDeData = objCon.retornaVazioParaValorNulo(drRow("data_inicial_viagem"), Nothing)
    ''' </summary>
    ''' <param name="valor"></param>
    ''' <param name="valorRetorno"></param>
    ''' <returns></returns>
    Public Function retornaVazioParaValorNulo(ByVal valor As Object, Optional ByVal valorRetorno As Object = "") As Object
        'verificamos se a variavel esta vazia ou nulla e retornamos vazio e/ou nothing nos casos de data vazia
        If String.IsNullOrEmpty(If(IsDBNull(valor), valorRetorno, valor)) Then
            Return valorRetorno
        ElseIf IsDBNull(valor) Then 'novo
            Return valorRetorno
        Else
            Return valor
        End If
    End Function

    ''' <summary>
    ''' Transformar vazio em nulo ou na opção de retorno desejada
    ''' campoDeData = objCon.retornaNuloParaVazio(drRow("data_inicial_viagem"))
    ''' campoDeData = objCon.retornaNuloParaVazio(drRow("data_inicial_viagem"), "O QUE QUISER")
    ''' </summary>
    ''' <param name="valor"></param>
    ''' <param name="valorRetorno"></param>
    ''' <returns></returns>
    Public Function retornaNuloParaVazio(ByVal valor As Object, Optional ByVal valorRetorno As Object = "") As Object
        If valor = "" Then
            If valorRetorno = "" Then
                Return vbNull
            Else
                Return valorRetorno
            End If
        Else
            Return valor
        End If
    End Function

    'Protected Overrides Sub Finalize()
    '    desconectar()
    'End Sub
End Class