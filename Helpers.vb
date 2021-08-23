Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.Globalization
Imports System.DirectoryServices
Imports System.DirectoryServices.ActiveDirectory
Imports System.IO
Imports excel = Microsoft.Office.Interop.Excel
Imports System.Net.NetworkInformation

Public Class Helpers
    Public Const TITULO_ALERTA = "Alerta do Sistema"
    Public Enum FLAG_REDE
        ALGAR = 1
        BRADESCO = 2
    End Enum
    Public Enum FLAG_ARQUIVO
        Copy = 1
        Move = 2
        Delete = 3
    End Enum

    Public Function Decrypt(str As String) As String
        Dim b As Byte() = Convert.FromBase64String(str)
        Dim decryp As String = System.Text.ASCIIEncoding.ASCII.GetString(b)
        Return decryp
    End Function

    Public Function Encrypt(str As String) As String
        Dim b As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes(str)
        Dim encryp As String = Convert.ToBase64String(b)
        Return encryp
    End Function

    Public Function RemoverSimbolos(Valor As String) As String
        Dim Remover As String, i As Byte, Temp As String, Simbolos As String

        'Removendo símbolos
        Simbolos = "*-+'@Ø'-!$%&(),./:;?[\]^`{|}~¿¢£¤¥€+<>««»»∆√√□§©®°µ¼½¾ÁÀÂÄÃÅÆČÇʣÉÈÊËĔĞĢÍÌÎÏʪÑºÓÒÔÖŌØŒŜŞß™ʦÚÙÛÜŪŸЉЊЫѬ"
        Temp = Valor
        Try
            For s = 1 To Len(Simbolos)
                Remover = Mid(Simbolos, s, 1)
                For i = 1 To Len(Valor)
                    Temp = Replace(Temp, Remover, "")
                Next
            Next
            Return Trim(Temp)
        Catch ex As Exception
            Return Trim(Valor)
        End Try

    End Function

    Public Function removeWhitespace(fullString As String) As String
        Try
            If Not String.IsNullOrEmpty(fullString) Then
                Return New String(fullString.Where(Function(x) Not Char.IsWhiteSpace(x)).ToArray())
            Else
                Return String.Empty
            End If
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function
    Public Function desacentua(ByVal argTexto As String) As String
        'Função que retira acentos de qualquer texto.
        Dim strAcento As String
        Dim strNormal As String
        Dim strLetra As String
        Dim strNovoTexto As String = ""
        Dim intPosicao As Integer
        Dim i As Integer

        'Informa as duas sequências de caracteres, com e sem acento
        strAcento = "ÃÁÀÂÄÉÈÊËÍÌÎÏÕÓÒÔÖÚÙÛÜÝÇÑãáàâäéèêëíìîïõóòôöúùûüýçñ'*_"
        strNormal = "AAAAAEEEEIIIIOOOOOUUUUYCNaaaaaeeeeiiiiooooouuuuycn "

        'Retira os espaços antes e após
        argTexto = Trim(argTexto)
        'Para i de 1 até o tamanho do texto
        For i = 1 To Len(argTexto)
            'Retira a letra da posição atual
            strLetra = Mid(argTexto, i, 1)
            'Busca a posição da letra na sequência com acento
            intPosicao = InStr(1, strAcento, strLetra)
            'Se a posição for maior que zero
            If intPosicao > 0 Then
                'Retira a letra na mesma posição na
                'sequência sem acentos.
                strLetra = Mid(strNormal, intPosicao, 1)
            End If
            'Remonta o novo texto, sem acento
            strNovoTexto = strNovoTexto & strLetra
        Next
        'Devolve o resultado
        desacentua = strNovoTexto
    End Function

    Public Function validaCPF(ByVal argCpf As String) As Boolean
        'Função que verifica a validade de um CPF.
        Dim wSomaDosProdutos
        Dim wResto
        Dim wDigitChk1
        Dim wDigitChk2
        Dim wI
        'Inicia o valor da Soma
        wSomaDosProdutos = 0
        'Para posição I de 1 até 9
        For wI = 1 To 9
            'Soma = Soma + (valor da posição dentro do CPF x (11 - posição))
            wSomaDosProdutos = wSomaDosProdutos + Val(Mid(argCpf, wI, 1)) * (11 - wI)
        Next wI
        'Resto = Soma - ((parte inteira da divisão da Soma por 11) x 11)
        wResto = wSomaDosProdutos - Int(wSomaDosProdutos / 11) * 11
        'Dígito verificador 1 = 0 (se Resto=0 ou 1 ) ou 11 - Resto (nos casos restantes)
        wDigitChk1 = IIf(wResto = 0 Or wResto = 1, 0, 11 - wResto)
        'Reinicia o valor da Soma
        wSomaDosProdutos = 0
        'Para posição I de 1 até 9
        For wI = 1 To 9
            'Soma = Soma + (valor da posição dentro do CPF x (12 - posição))
            wSomaDosProdutos = wSomaDosProdutos + (Val(Mid(argCpf, wI, 1)) * (12 - wI))
        Next wI
        'Soma = Soma (2 x dígito verificador 1)
        wSomaDosProdutos = wSomaDosProdutos + (2 * wDigitChk1)
        'Resto = Soma - ((parte inteira da divisão da Soma por 11) x 11)
        wResto = wSomaDosProdutos - Int(wSomaDosProdutos / 11) * 11
        'Dígito verificador 2 = 0 (se Resto=0 ou 1 ) ou 11 - Resto (nos casos restantes)
        wDigitChk2 = IIf(wResto = 0 Or wResto = 1, 0, 11 - wResto)
        'Se o dígito da posição 10 = Dígito verificador 1 E
        'dígito da posição 11 = Dígito verificador 2 Então
        If Mid(argCpf, 10, 1) = Mid(Trim(Str(wDigitChk1)), 1, 1) And Mid(argCpf, 11, 1) = Mid(Trim(Str(wDigitChk2)), 1, 1) Then
            'CPF válido
            validaCPF = True
        Else
            'CPF inválido
            validaCPF = False
        End If
    End Function

    'Usar no KeyPress do componente
    Public Function somenteNumero(ctrl As Control) As Boolean
        If Not IsNumeric(ctrl.Text.Trim) And Not String.IsNullOrEmpty(ctrl.Text) Then
            MsgBox("Número do " & ctrl.Tag & " inválido.", MsgBoxStyle.Information, TITULO_ALERTA)
            ctrl.Focus()
            ctrl.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
            Return False
            Exit Function
        Else
            ctrl.BackColor = System.Drawing.Color.White
            Return True
        End If
    End Function

    Public Function retornaSoNumeroDeString(texto As String) As String
        'Dim i As Integer, j As String = ""
        'Dim parteNumerica As String = ""
        'For i = 1 To Len(texto)
        '    If Asc(Mid(texto, i, 1)) < 48 Or
        '       Asc(Mid(texto, i, 1)) > 57 Then
        '    Else
        '        j = j & Mid(texto, i, 1)
        '    End If
        '    parteNumerica = j
        'Next
        'Return parteNumerica
        Try
            Dim re As Regex = New Regex("[0-9]")
            Dim s As String = String.Empty
            For Each m As Match In re.Matches(texto)
                s += m.Value
            Next
            Return s
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function
    Public Function retornaSoTextoDeString(texto As String) As String
        Try
            Dim reLetras As Regex = New Regex("[a-zA-Z]")
            Dim s As String = String.Empty
            For Each m As Match In reLetras.Matches(texto)
                s += m.Value
            Next
            Return s
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function

    Public Function validaEmail(ByVal eMail As String) As Boolean
        'Função de validação do formato de um e-mail.

        Dim posicaoA As Integer
        Dim posicaoP As Integer

        'Busca posição do caracter @
        posicaoA = InStr(eMail, "@")
        'Busca a posição do ponto a partir da posição
        'do @ ou então da primeira posição
        posicaoP = InStr(posicaoA Or 1, eMail, ".")

        'Se a posição do @ for menor que 2 OU
        'a posição do ponto for menor que a posição
        'do caracter @
        If posicaoA < 2 Or posicaoP < posicaoA Then
            'Formato de e-mail inválido
            validaEmail = False
        Else
            'Formato de e-mail válido
            validaEmail = True
        End If

    End Function

    Public Function nomeProprio(ByVal argNome As String) As String
        'Função recursiva para converter a primeira letra
        'dos nomes próprios para maiúscula, mantendo os
        'aditivos em caixa baixa.
        Dim sNome As String
        Dim lEspaco As Long
        Dim lTamanho As Long
        'Pega o tamanho do nome
        lTamanho = Len(argNome)
        'Passa tudo para caixa baixa
        argNome = LCase(argNome)
        'Se o nome passado é vazio
        'acaba a função ou a recursão
        'retornando string vazia
        If lTamanho = 0 Then
            nomeProprio = ""
        Else
            'Procura a posição do primeiro espaço
            lEspaco = InStr(argNome, " ")
            'Se não tiver pega a posição da última letra
            If lEspaco = 0 Then lEspaco = lTamanho
            'Pega o primeiro nome da string
            sNome = Left(argNome, lEspaco)
            'Se não for aditivo converte a primeira letra
            If Not InStr("e da das de do dos ", sNome) > 0 Then
                sNome = UCase(Left(sNome, 1)) & LCase(Right(sNome, Len(sNome) - 1))
            End If
            'Monta o nome convertendo o restante através da recursão
            nomeProprio = sNome & nomeProprio(LCase(Trim(Right(argNome, lTamanho - lEspaco))))
        End If
    End Function

    Public Function abreviaNome(ByVal argNome As String) As String
        'Função que abrevia o penúltimo sobrenome, levando
        'em consideração os aditivos de, da, do, dos, das, e.

        'Define variáveis para controle de posição e para as
        'partes do nome que serão separadas e depois unidas
        'novamente.
        Dim ultimoEspaco As Integer, penultimoEspaco As Integer
        Dim primeiraParte As String, ultimaParte As String
        Dim parteNome As String
        Dim tamanho As Integer, i As Integer

        'Tamanho do nome passado
        'no argumento
        tamanho = Len(argNome)

        'Loop que verifica a posição do último e do penúltimo
        'espaços, utilizando apenas um loop.
        For i = tamanho To 1 Step -1
            If Mid(argNome, i, 1) = " " And ultimoEspaco <> 0 Then
                penultimoEspaco = i
                Exit For
            End If
            If Mid(argNome, i, 1) = " " And penultimoEspaco = 0 Then
                ultimoEspaco = i
            End If
        Next i

        'Caso i chegue a zero não podemos
        'abreviar o nome
        If i = 0 Then
            abreviaNome = argNome
            Exit Function
        End If

        'Separação das partes do nome em três: primeira, meio e última
        primeiraParte = Left(argNome, penultimoEspaco - 1)
        parteNome = Mid(argNome, penultimoEspaco + 1, ultimoEspaco - penultimoEspaco - 1)
        ultimaParte = Right(argNome, tamanho - ultimoEspaco)

        'Para a montagem do nome já abreviado verificamos se a parte retirada
        'não é um dos nomes de ligação: de, da ou do. Caso seja usamos o método
        'recursivo para refazer os passos.
        'Caso seja necessário basta acrescentar outros nomes de ligação para serem
        'verificados.
        If parteNome = "da" Or parteNome = "de" Or parteNome = "do" Or parteNome = "dos" Or parteNome = "das" Or parteNome = "e" Then
            abreviaNome = abreviaNome(primeiraParte & " " & parteNome) & " " & ultimaParte
        Else
            abreviaNome = primeiraParte & " " & Left(parteNome, 1) & ". " & ultimaParte
        End If
    End Function

    'função para ajustar um valor em decimal
    Public Function transformarMoedaValidandoCampo(ctrl As Control) As String
        Dim valor As String = ctrl.Text
        Dim n As String = String.Empty
        Dim v As Double = 0

        valor = Replace(valor, "$", "")
        valor = Replace(valor, "R", "")

        Try
            'Verificando se o valor contem ',' ou '.' ou ausencia de pontuação
            If InStr(valor, ".") Or InStr(valor, ",") Then
                n = valor.Replace(",", "").Replace(".", "")
                If n.Equals("") Then n = "000"
                If n.Length > 3 And n.Substring(0, 1) = "0" Then n = n.Substring(1, n.Length - 1)
            Else 'Caso não haja pontuação apenas acrescenta 2 zeros para gerar o valor moeda
                n = valor.PadRight(valor.Length + 2, "0")
            End If
            v = Convert.ToDouble(n) / 100
            valor = String.Format("{0:C}", v)
            Return valor

        Catch ex As Exception
            MsgBox("Valor digitado não é um valor de moeda. Tente novamente!", vbInformation, TITULO_ALERTA)
            ctrl.Focus()
            Return "Erro"
            Exit Function
        End Try
    End Function
    Public Function transformarMoeda(ctrl As String) As Double
        Dim valor As String = ctrl
        Dim n As String = String.Empty
        Dim v As Double = 0
        Try
            'Formatando para duas casas decimais antes das validações
            'este procedimento corrigi um bug para numeros com apenas 1 digito nas casas decimais
            Dim getDuasCasasDecimais As String = Microsoft.VisualBasic.Right(ctrl, 2)
            Dim getVirgulaouPontoCasasDecimais As String = Microsoft.VisualBasic.Left(getDuasCasasDecimais, 1)
            If getVirgulaouPontoCasasDecimais = "." Or getVirgulaouPontoCasasDecimais = "," Then
                valor = ctrl.PadRight(ctrl.Length + 1, "0")
            End If

            'Verificando se o valor contem ',' ou '.' ou ausencia de pontuação
            If InStr(valor, ".") Or InStr(valor, ",") Then
                n = valor.Replace(",", "").Replace(".", "")
                If n.Equals("") Then n = "000"
                If n.Length > 3 And n.Substring(0, 1) = "0" Then n = n.Substring(1, n.Length - 1)

            Else 'Caso não haja pontuação apenas acrescenta 2 zeros para gerar o valor moeda
                n = valor.PadRight(valor.Length + 2, "0")
            End If
            v = Convert.ToDouble(n) / 100
            Return CDbl(v)
            'valor = String.Format("{0:C}", v) 'acess {0:N}
        Catch ex As Exception
            Return 0
            Exit Function
        End Try
    End Function

    Public Function removerCharEspecial(strIn As String) As String
        ' Replace invalid characters with empty strings.
        Dim retorno As String = ""
        Try
            retorno = Regex.Replace(strIn, "[^\w\.@-]", " ")
            retorno = Regex.Replace(retorno, "[^0-9a-zA-Z]+", " ").Trim
            Return retorno
            ' If we timeout when replacing invalid characters, 
            ' we should return String.Empty.
        Catch e As Exception
            Return String.Empty
        End Try
    End Function

    'Função para validar o preenchimento de campos obrigatórios
    'argForm = nome do formulario
    'strCamposObrigatorios = lista do com o nome dos campos separados por ";"
    'tituloCampos = lista dos titulos dos campos na mesma ordem e separados por ";"
    'validaCamposObrigatorios(Me, "nomeCampo1;nomeCampo2;etc", "TituloCampo1;TituloCampo2;etc")
    Public Function validaCamposObrigatorios(ByVal argForm As Control, ByVal strCamposObrigatorios As String, Optional ByVal tituloCampos As String = "") As Boolean
        Try

            Dim nomeCampos As Object
            Dim campos As Object
            Dim valor As Object
            Dim i As Long
            Dim inicio As Long
            Dim fim As Long
            Dim ctrl As String
            'Windows.Forms.Form
            'monta os arrays
            campos = Split(strCamposObrigatorios, ";")
            nomeCampos = Split(tituloCampos, ";")
            'captura o inicio e fim do array
            inicio = LBound(campos)
            fim = UBound(campos)
            i = inicio

            'inicia a validação uma a uma
            For i = inicio To fim
                'captura o nome do tipo de campo
                ctrl = argForm.Controls(campos(i)).GetType.Name
                Select Case ctrl
                'Caso seja ComboBox
                    Case "ComboBox"
                        valor = argForm.Controls(campos(i)).Text
                        If String.IsNullOrEmpty(valor) Then
                            MsgBox("Uma opção: " & argForm.Controls(campos(i)).Tag & ". Deve ser selecionada.", MsgBoxStyle.Information, TITULO_ALERTA)
                            'argForm(campos(i)).SetFocus() 'Coloca o cursor no campo
                            argForm.Controls(campos(i)).Focus()
                            'argForm.Controls(campos(i)).BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                            validaCamposObrigatorios = False
                            Exit Function
                        Else
                            'Altera a cor de fundo para branco
                            'argForm.Controls(campos(i)).BackColor = System.Drawing.Color.White
                        End If
                    'Caso seja TextBox
                    Case "TextBox"
                        valor = argForm.Controls(campos(i)).Text
                        If String.IsNullOrEmpty(valor) Then
                            MsgBox("O Campo: " & argForm.Controls(campos(i)).Tag & ". Deve ser preenchido.", MsgBoxStyle.Information, TITULO_ALERTA)
                            argForm.Controls(campos(i)).Focus()
                            'argForm.Controls(campos(i)).BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                            validaCamposObrigatorios = False
                            Exit Function
                        Else
                            'Altera a cor de fundo para branco
                            'argForm.Controls(campos(i)).BackColor = System.Drawing.Color.White
                        End If
                    'Caso seja MaskeCheckBox
                    Case "MaskedTextBox"
                        'tira a formatação
                        argForm.Controls(campos(i)).TextMaskFormat = MaskFormat.ExcludePromptAndLiterals
                        'captura o valor sem a mascara
                        valor = argForm.Controls(campos(i)).Text
                        'retorna a formatação
                        argForm.Controls(campos(i)).TextMaskFormat = MaskFormat.IncludePromptAndLiterals
                        If String.IsNullOrEmpty(valor) Then
                            MsgBox("O Campo: " & argForm.Controls(campos(i)).Tag & ". Deve ser preenchido.", MsgBoxStyle.Information, TITULO_ALERTA)
                            argForm.Controls(campos(i)).Focus()
                            'argForm.Controls(campos(i)).BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                            validaCamposObrigatorios = False
                            Exit Function
                        Else
                            'Altera a cor de fundo para branco
                            'argForm.Controls(campos(i)).BackColor = System.Drawing.Color.White
                        End If
                    'Caso seja CheckBox (Normalmente esta campo é opcional)
                    Case "CheckBox"
                        '    'Caso seja OptionButton
                        'Case "OptionButton"
                        '    valor = argForm.Controls(campos(i)).Text
                        '    If String.IsNullOrEmpty(valor) Then
                        '        MsgBox("Uma opção: " & argForm.Controls(campos(i)).Tag & ". Deve ser selecionada.", MsgBoxStyle.Information, TITULO_ALERTA)
                        '        argForm.Controls(campos(i)).Focus()
                        '        'argForm.Controls(campos(i)).BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                        '        validaCamposObrigatorios = False
                        '        Exit Function
                        '    End If

                        '    'Caso seja OptionGroup
                        'Case "OptionGroup"
                        '    valor = argForm.Controls(campos(i)).Text
                        '    If String.IsNullOrEmpty(valor) Then
                        '        MsgBox("Uma opção: " & argForm.Controls(campos(i)).Tag & ". Deve ser selecionada.", MsgBoxStyle.Information, TITULO_ALERTA)
                        '        argForm.Controls(campos(i)).Focus()
                        '        'argForm.Controls(campos(i)).BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                        '        validaCamposObrigatorios = False
                        '        Exit Function
                        '    End If

                End Select
            Next i
            validaCamposObrigatorios = True
        Catch ex As Exception
            MsgBox("Falha ao validar campos obrigatórios.", vbCritical, TITULO_ALERTA)
            Return False
        End Try
    End Function
    ''Limpa objetos de um determinado formulário
    Public Sub limparCampos(ByVal Tela As Control, Optional NomeCamposIgnorar As List(Of Object) = Nothing)
        Try
            'Declaramos uma variavel Campo do tipo Object
            '(Tipo Object porque iremos trabalhar com todos os campos do Form, podendo ser
            '       Label, Button, TextBox, ComboBox e outros)
            Dim Campo As Object

            'Usaremos For Each para passarmos por todos os controls do objeto atual
            For Each Campo In Tela.Controls
                'verificando se temos que ignorar alguns campos:
                Dim ignora As Boolean = False
                If NomeCamposIgnorar IsNot Nothing Then
                    For Each item In NomeCamposIgnorar
                        If item.ToString = Campo.Name.ToString Then
                            ignora = True
                            Exit For
                        End If
                    Next
                End If
                If Not ignora Then
                    'Verifica se o Campo é um GroupBox, TabPage ou Panel
                    'Se for então precisa limpar os campos que estão dentro dele também...
                    'Chamaremos novamente a função mas passando por referencia
                    '      O GroupBox, TabPage ou Panel atual
                    If TypeOf Campo Is System.Windows.Forms.GroupBox Or
                    TypeOf Campo Is System.Windows.Forms.TabPage Or
                    TypeOf Campo Is System.Windows.Forms.Panel Then
                        limparCampos(Campo)
                    ElseIf TypeOf Campo Is System.Windows.Forms.TextBox Then
                        Campo.Text = String.Empty 'Verificamos se o campo é uma TextBox se for então devemos limpar o campo
                    ElseIf TypeOf Campo Is System.Windows.Forms.ComboBox Then
                        'Verificamos se o campo é um ComboBox
                        If Campo.DropDownStyle = ComboBoxStyle.DropDownList Then
                            Campo.SelectedIndex = -1 'Se o tipo da ComboBox for DropDownList então devemos deixar sem seleção
                            'ElseIf Campo.DropDownStyle = ComboBoxStyle.DropDown Then
                            'Campo.Text = ""
                        Else
                            'Campo.Text = String.Empty
                            Campo.SelectedValue = 0
                        End If
                    ElseIf TypeOf Campo Is System.Windows.Forms.CheckBox Then
                        Campo.Checked = False
                    ElseIf TypeOf Campo Is System.Windows.Forms.DataGridView Then
                        Campo.DataSource = Nothing
                    ElseIf TypeOf Campo Is System.Windows.Forms.RadioButton Then
                        Campo.Checked = False
                    ElseIf TypeOf Campo Is System.Windows.Forms.MaskedTextBox Then
                        Campo.Text = String.Empty
                    End If
                End If
            Next
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical, TITULO_ALERTA)
        End Try
    End Sub

    'Função para limpar as infos de um combobox
    Public Sub limpaCombobox(cb As ComboBox)
        With cb
            .DataSource = Nothing
            .DisplayMember = Nothing
            .Items.Clear()
        End With
    End Sub



    'função para abrir um formulario
    Public Sub abrirForm(frm As Form, Optional janelaRestrita As Boolean = False, Optional maximiza As Boolean = False)
        Try

            If maximiza Then
                frm.WindowState = FormWindowState.Maximized
            Else
                frm.WindowState = FormWindowState.Normal
            End If

            If janelaRestrita Then
                frm.ShowDialog()
            Else
                frm.Show()
            End If
        Catch ex As Exception
            MsgBox("Falha ao abrir o formulário.", vbCritical, TITULO_ALERTA)
        End Try
    End Sub

    Public Sub fecharAplicativo()
        'Função para fechar o aplicativo
        If MsgBox("Deseja realmente fechar o aplicativo?", vbQuestion + vbYesNo, TITULO_ALERTA) = vbYes Then
            Application.Exit()
        End If
    End Sub

    'função para fechar um formulario
    Public Sub fecharForm(frm As Form)
        frm.Close()
    End Sub

    'função para capturar o id de rede
    Public Function capturaIdRede() As String
        capturaIdRede = Environ("USERNAME").ToString
    End Function
    'função para capturar o hostName
    Public Function capturaHostname() As String
        capturaHostname = System.Net.Dns.GetHostName
        'capturaHostname = System.Environment.MachineName.ToString()
    End Function

    'função para capturar o endereço IP
    Public Function capturaEnderecoIP() As String
        Dim myHost As String = System.Net.Dns.GetHostName
        Dim myIPs As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(myHost)
        Dim ip As String = ""
        For Each myIP As System.Net.IPAddress In myIPs.AddressList
            ip = myIP.ToString
        Next
        Return ip
    End Function
    'função para capturar o endereço MAC
    Public Function capturaMac() As String
        Dim nics() As NetworkInterface = NetworkInterface.GetAllNetworkInterfaces()
        Dim mac As String = nics(0).GetPhysicalAddress.ToString
        Return mac
    End Function

    'função para limitar uma quantidade minima e maxima de caracteres.
    'Utilizar no LostFocus
    'txtCartao_Leave(sender As Object, e As EventArgs) Handles txtCartao.Leave
    'hlp.validaTamanhoMinMax(txtCartao, 15, 15)
    Public Function validaTamanhoMinMax(ByVal ctl As Control, iMinLen As Integer, iMaxLen As Integer) As Boolean
        If Not String.IsNullOrEmpty(ctl.Text) Then
            Dim texto As String = Trim(Replace(ctl.Text, " ", ""))

            'se diferente de vazio
            If Not String.IsNullOrEmpty(Replace(texto.Trim, "_", "")) Then
                'Limite Maximo
                If Len(Replace(texto.Trim, "_", "")) > iMaxLen Then
                    MsgBox("Limite máximo de " & iMaxLen & " caracteres foi excedido." & vbNewLine, vbInformation, TITULO_ALERTA)
                    ctl.Text = Left(texto.Trim, iMaxLen)
                    ctl.Focus()
                    ctl.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                    Return False
                    Exit Function
                End If
                'Limite Minimo
                If Len(Replace(texto.Trim, "_", "")) < iMinLen Then
                    MsgBox("Limite mínimo de " & iMinLen & " caracteres." & vbNewLine, vbInformation, TITULO_ALERTA)
                    ctl.Text = Left(texto.Trim, iMinLen)
                    ctl.Focus()
                    ctl.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                    Return False
                    Exit Function
                End If
                ctl.BackColor = System.Drawing.Color.White
                ctl.Text = texto
            Else
                ctl.BackColor = System.Drawing.Color.White
                Return False
                Exit Function
            End If
        End If
        Return True
    End Function

    'verificar se uma data é valida.
    'Utilizar no LostFocus
    Public Function validaData(ByVal Controle As Control) As Boolean

        Try
            Dim idiomaPC As String
            Dim formato As String = ""
            'se diferente de vazio
            If Not String.IsNullOrEmpty(Replace(Replace(Controle.Text.Trim, "_", ""), "/", "").Trim) Then
                If Not IsDate(Controle.Text) Then
                    'captura o idioma da maquina
                    idiomaPC = CultureInfo.CurrentCulture.Name
                    If idiomaPC = "pt-BR" Then
                        formato = "dia/mês/ano"
                    Else
                        formato = "mês/dia/ano"
                    End If
                    'para o campo: " & Controle.Tag & "." & vbNewLine &
                    MsgBox("Data inválida! " & vbNewLine &
                           "Possíveis motivos: " & vbNewLine &
                           " > Data inexistente." & vbNewLine &
                           " > Utilize o formato: " & idiomaPC.ToUpper & " (" & formato.ToUpper & ").", MsgBoxStyle.Information, TITULO_ALERTA)
                    Controle.BackColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(128, Byte))
                    Controle.Focus()
                    Return False
                Else
                    Controle.BackColor = System.Drawing.Color.White
                    Return True
                End If
            Else
                Controle.BackColor = System.Drawing.Color.White
                Return False
            End If

        Catch ex As Exception
            Controle.BackColor = System.Drawing.Color.White
            Return False
        End Try
    End Function

    'Funções para formatação de data
    Public Function dataHoraAtual() As DateTime
        Return DateTime.Now
    End Function

    Public Function dataAbreviada() As Date
        Return CDate(DateTime.Now).ToString("yyyy-MM-dd")
    End Function

    Public Function formataHoraAbreviada(hr As DateTime) As Date
        Return CDate(hr).ToString("HH:mm:ss")
    End Function

    Public Function FormataDataAbreviada(dt As Object) As DateTime
        If Replace(Replace(dt, "/", ""), "_", "").Trim = Nothing Then
            Return Nothing
        End If
        Return CDate(dt).ToString("yyyy-MM-dd")
    End Function

    Public Function FormataDataHoraCompleta(hr As DateTime) As Date
        Return CDate(hr).ToString("yyyy-MM-dd HH:mm:ss")
    End Function

    Public Function convertDatetime(data As Object) As DateTime
        If IsDBNull(data) Then
            Return Nothing
        Else
            Return Convert.ToDateTime(data).ToString
        End If
    End Function

    Public Sub abrirArquivo(arquivo As String)
        System.Diagnostics.Process.Start(arquivo)
    End Sub

    Public Function isFileOpen(ByVal filename As String) As Boolean

        Dim filenum As Integer, errnum As Integer
        On Error Resume Next   ' Turn error checking off.
        filenum = FreeFile()   ' Get a free file number.
        FileOpen(filenum, filename, OpenMode.Random, OpenAccess.ReadWrite)
        FileClose(filenum)  'close the file.
        errnum = Err.Number 'Save the error number that occurred.
        On Error GoTo 0        'Turn error checking back on.
        ' Check to see which error occurred.
        Select Case errnum
            ' No error occurred.
            ' File is NOT already open by another user.
            Case 0
                Return False
                ' Error number for "Permission Denied."
                ' File is already opened by another user.
            Case 70, 55, 75
                Return True
                ' Another error occurred.
            Case Else
                Error errnum
        End Select

    End Function

    Public Function pegarExtensao(arquivo As String) As String
        Dim i As Integer
        Dim j As Integer
        i = InStrRev(arquivo, ".")
        j = InStrRev(arquivo, "\")
        If j = 0 Then j = InStrRev(arquivo, ":")
        'End If
        If j < i Or i > 0 Then
            pegarExtensao = Right(arquivo, (Len(arquivo) - i))
        Else
            Return ""
        End If
    End Function
    'para copiar para um determinado local
    Public Sub copiaArquivo(ByVal origem As String, ByVal destino As String, ByVal arquivo As String, Optional ByVal id As String = "")
        Dim novoNome As String = ""
        Try
            origem = Replace(origem, arquivo, "") & arquivo
            'atribui um novo nome unico para o arquivo
            novoNome = id & " " & capturaIdRede() & " " & Format(Now, "ddMMyyyy HHmmss") & "." & pegarExtensao(arquivo)
            destino = destino & novoNome

            'novoNome = capturaIdRede() & " " & Format(Now, "ddMMyyyy HHmmss") & "." & PegarExtensao(arquivo)
            'se o arquivo não existir
            If Len(Dir(destino)) = 0 Then
                'se não existir nenhum arquivo, pode copiar
                Microsoft.VisualBasic.FileCopy(origem, destino)
                'caso existir apaga e depois copia
            Else
                'verifica se o arquivo ja esta em uso
                If isFileOpen(destino) Then
                    MsgBox("Arquivo já em uso!", vbInformation, TITULO_ALERTA)
                    Exit Sub
                End If
                'apaga o arquivo antigo
                Microsoft.VisualBasic.Kill(destino)
                'copia o novo arquivo
                Microsoft.VisualBasic.FileCopy(origem, destino)
            End If
        Catch ex As Exception
            MsgBox(ex.Message & "Erro nº: " & Err.Number, vbCritical, TITULO_ALERTA)
        End Try
    End Sub

    Public Function copiaArquivoRetornaNome(ByVal origem As String, ByVal destino As String, ByVal arquivo As String, Optional ByVal id As String = "") As String
        Dim novoNome As String = ""
        Try
            origem = Replace(origem, arquivo, "") & arquivo
            'atribui um novo nome unico para o arquivo
            novoNome = id & " " & capturaIdRede() & " " & Format(Now, "ddMMyyyy HHmmss") & "." & pegarExtensao(arquivo)
            destino = destino & novoNome

            'novoNome = capturaIdRede() & " " & Format(Now, "ddMMyyyy HHmmss") & "." & PegarExtensao(arquivo)
            'se o arquivo não existir
            If Len(Dir(destino)) = 0 Then
                'se não existir nenhum arquivo, pode copiar
                Microsoft.VisualBasic.FileCopy(origem, destino)
                'caso existir apaga e depois copia
            Else
                'verifica se o arquivo ja esta em uso
                If isFileOpen(destino) Then
                    MsgBox("Este arquivo esta em uso.", vbCritical, TITULO_ALERTA)
                    Return novoNome
                    Exit Function
                End If
                'apaga o arquivo antigo
                Microsoft.VisualBasic.Kill(destino)
                'copia o novo arquivo
                Microsoft.VisualBasic.FileCopy(origem, destino)
            End If
            Return novoNome
        Catch ex As Exception
            MsgBox(ex.Message & "Erro nº: " & Err.Number, vbCritical, TITULO_ALERTA)
        End Try
        Return novoNome
    End Function

    'Metodo para desabilitar o botão "X" fechar
    'Disable the button on the current form:
    'RemoveXButton(Me.Handle())
    Public Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Integer, ByVal bRevert As Integer) As Integer
    Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Integer, ByVal nPosition As Integer, ByVal wFlags As Integer) As Integer
    Public Const SC_CLOSE = &HF060&
    Public Const MF_BYCOMMAND = &H0&
    Public Function removeXButton(ByVal iHWND As Integer) As Integer
        Dim iSysMenu As Integer
        iSysMenu = GetSystemMenu(iHWND, False)
        Return RemoveMenu(iSysMenu, SC_CLOSE, MF_BYCOMMAND)
    End Function

    Public Function retornaIdiomaPC() As String
        Return CultureInfo.CurrentCulture.Name.ToUpper.Trim
    End Function

    Public Function totalLinhasArquivoTxt(caminhoArquivo As String) As Long
        'Regular Expression e contar as quebras de linhas:
        Dim re As New System.Text.RegularExpressions.Regex("\r\n")
        Dim sr As New System.IO.StreamReader(caminhoArquivo)
        Dim txt As String = sr.ReadToEnd()
        Dim qtdLinhas As Long = re.Matches(txt).Count + 1
        sr.Close()
        'No final eu somo +1 para contar a última linha, que não tem a quebra de linha que é contada acima.
        Return qtdLinhas
    End Function

    'Procura uma determinada palavra em um texto e retorna verdadeiro caso encontre
    Public Function procurarPalavra(texto As String, texto_a_procurar As String) As Boolean
        Dim resultado As Long
        resultado = InStr(texto, texto_a_procurar)
        If resultado > 0 Then
            procurarPalavra = True
        Else : procurarPalavra = False
        End If
    End Function

    'Função para retornar vazio para campos textbox com data
    Public Function retornaDataTextBox(argValor As Object) As String
        Dim idiomaPC As String
        Dim formato As String = ""
        Dim dataVazia As DateTime = Nothing
        If argValor = Nothing Then
            argValor = dataVazia
        End If
        If CDate(argValor).ToString("yyyy-MM-dd HH:mm:ss") = CDate(dataVazia).ToString("yyyy-MM-dd HH:mm:ss") Then
            Return ""
        Else
            'captura o idioma da maquina
            idiomaPC = CultureInfo.CurrentCulture.Name
            If idiomaPC = "pt-BR" Then
                formato = "dd/MM/yyyy HH:mm:ss" 'dia/mês/ano
            Else
                formato = "MM/dd/yyyy HH:mm:ss" 'mês/dia/ano"
            End If
            Return CDate(argValor).ToString(formato)
        End If
    End Function

    'Chamada
    'hlp.CarregaComboBoxManualmente("FAB;INQ;Tudo", Me, Me.cFilas)
    'função
    'Carregamento de Combobox de forma manual
    Public Sub carregaComboBoxManualmente(ByVal strItens As String, ByVal frm As Form, ByVal cb As ComboBox)
        Dim itens As Object
        itens = Split(strItens, ";")
        'limpando o combobox para evitar duplo carregamento
        With frm
            With cb
                .DataSource = Nothing
                .Items.Clear()
            End With
        End With
        'Carregando itens
        For i = LBound(itens) To UBound(itens)
            With frm
                With cb
                    .Items.Add(itens(i))
                End With
            End With
        Next
    End Sub


    'Função para preencher um combobox 2 colunas
    Public Sub carregaComboBox(dt As DataTable, frm As Form,
                               cb As ComboBox,
                               Optional selecaoDefault As Boolean = False,
                               Optional DisplayMember As String = "",
                               Optional ValueMember As String = "",
                               Optional preencherQdoDtVazio As Boolean = False)
        With frm
            With cb
                .DataSource = dt

                'validação para verificar se a necessidade de carregar os dados pelo nome das colunas
                'por default carregamos sempre as duas primeiras
                If String.IsNullOrEmpty(DisplayMember) Or String.IsNullOrEmpty(DisplayMember) Then
                    .DisplayMember = dt.Columns(1).ToString
                    .ValueMember = dt.Columns(0).ToString
                Else
                    .DisplayMember = DisplayMember.ToString
                    .ValueMember = ValueMember.ToString
                End If
                .Text = Nothing

                'Se não carregar nenhuma informação, entra com uma informação GERAL e de forma manual
                If .Items.Count = 0 And preencherQdoDtVazio Then
                    carregaComboBoxManualmente("NÃO SE APLICA", frm, cb)
                End If

                'Se houver apenas um item no combobox este já fica selecionado
                If selecaoDefault Then
                    If .Items.Count = 1 Then
                        .SelectedIndex = 0
                    End If
                End If

                'Limpando caso não exista registros
                If .Items.Count = 0 Then
                    .DataSource = Nothing
                End If

            End With
        End With
    End Sub

    Public Sub desligarReiniciarWindows(ByVal executarFuncao_D_R As String)

        'No args                 Display this message (same as -?)
        '-i                      Display GUI interface, must be the first option
        '-l                      Log off (cannot be used with -m option)
        '-s                      Shutdown the computer
        '-r                      Shutdown And restart the computer
        '-a                      Abort a system shutdown
        '-m \\computername       Remote computer to shutdown/restart/abort
        '-t xx                   Set timeout for shutdown to xx seconds
        '-c "comment"            Shutdown comment (maximum of 127 characters)
        '-f                      Forces running applications to close without warning
        '-d [u][p]:xx : yy         The reason code for the shutdown
        '                        u Is the user code
        '                        p Is a planned shutdown code
        '                        xx Is the major reason code (positive integer less than 256)
        '                        yy Is the minor reason code (positive integer less than 65536)

        Try
            Select Case executarFuncao_D_R.ToUpper
                Case "D"
                    System.Diagnostics.Process.Start("shutdown", "-s -t 00 -f")
                Case "R"
                    System.Diagnostics.Process.Start("shutdown", "-r -t 00 -f")
            End Select
        Catch ex As Exception
            MsgBox("Ocorreu uma falha ao tentar Desligar/Reiniciar o windows.")
        End Try

    End Sub


    'matar processo do proprio aplicativo
    Public Sub killProcesso()
        'captura o processo do aplicativo
        Dim proc As Process = Process.GetCurrentProcess
        'captura o nome do processo deste aplicativo
        Dim processo As String = proc.ProcessName.ToString
        'percorrendo todos os processos abertos
        For Each prog As Process In Process.GetProcesses
            'fecha o processo deste aplicativo
            If prog.ProcessName = processo Then
                prog.Kill()
            End If
        Next
    End Sub

    'Função para limitar uma quantidade de caracteres por linha em um determinado Textbox
    Public Sub limiteCaracterPorLinha(ByVal limite As Long, ByVal ctrl As Control)
        Dim texto As String = ""
        Dim tamanho As Long = 0
        Dim nova_linha As String = ""
        Dim temp_linha As String = ""
        Dim delimitador As String = Replace(Space(limite), " ", "-")
        Dim nroBloco As Integer = 0

        texto = ctrl.Text.Trim
        'remove as quebras de linhas
        texto = texto.Replace(System.Environment.NewLine, String.Empty)
        tamanho = Len(texto)
        'se acima do limite
        If tamanho > limite Then
            'percorre toda a cadeia de caracteres uma a uma
            For i = 1 To tamanho
                temp_linha += Mid(texto, i, 1) 'recebe os caracteres
                If temp_linha.Length = limite Then 'verifica se alcançou o limite
                    nroBloco = nroBloco + 1
                    temp_linha += "\r\n" 'inserir quebra de linha na variavel temp
                    ''utilizar separador por blocos apenas se necessario
                    'If nroBloco = 4 Then
                    '    temp_linha += delimitador & "\r\n"
                    '    nroBloco = 0
                    'End If
                    nova_linha += temp_linha 'salva na variavel final
                    'formatando uma expressão regular para quebra de linha
                    nova_linha = System.Text.RegularExpressions.Regex.Unescape(String.Format(nova_linha))
                    temp_linha = "" 'limpa variavel temporaria para o proximo lote de caracteres
                ElseIf i = tamanho Then
                    nova_linha += temp_linha 'concatenar a ultima linha abaixo de 50 caracteres
                End If
            Next
            ctrl.Text = nova_linha 'retorna para o textbox
        End If
    End Sub

    Public Function formataLimiteCaracteres(nrCaracteres As Integer, valor As String) As String
        If String.IsNullOrEmpty(valor) Then
            Return Nothing
        Else
            Dim i As Long
            'Dim novovalor As String = valor.Trim
            Dim strRetorno As String = ""
            For i = 1 To nrCaracteres
                strRetorno = strRetorno & "0"
            Next
            Return Microsoft.VisualBasic.Right(strRetorno & valor.Trim, nrCaracteres)
        End If
    End Function

    Public Function getCurrentMethodName() As String
        Dim stack As New System.Diagnostics.StackFrame(1)
        Return stack.GetMethod().Name
    End Function

    'Função para converter segundos em hora / minuto / segundos
    Public Function converterSegundos(ByVal intSegundos As Long) As DateTime

        Dim emSegundos As Long, emMinutos As Long, emHoras As Long, emDias As Long
        Dim segundos As Long, miuntos As Long, horas As Long

        emSegundos = intSegundos
        segundos = emSegundos Mod (60)
        emMinutos = emSegundos \ (60)
        miuntos = emMinutos Mod (60)
        emHoras = emMinutos \ (60)
        horas = emHoras Mod (24)
        emDias = emHoras \ (24)

        Return Format(horas, "00") & ":" & Format(miuntos, "00") & ":" & Format(segundos, "00")

    End Function

    Public Function validarIdiomaPC(ByVal siglaIdioma As String) As Boolean

        'IDIOMAS MAIS USADOS:
        'PT-BR
        'EN-US

        'EXEMPLO:
        'If Not hlp.validarIdiomaPC("PT-BR") Then Exit Sub

        Dim idiomaPC As String = retornaIdiomaPC.ToLower

        If Not idiomaPC = siglaIdioma.ToLower Then
            MsgBox("O idioma para esta ação deve ser: " & siglaIdioma.ToUpper & ". " _
                    & vbNewLine & "Feche o aplicativo, troque o idioma e tente outra vez!" _
                    , MsgBoxStyle.Information, TITULO_ALERTA)
            Return False
        Else
            Return True
        End If

    End Function

    Public Function apenasNumeros(strOriginal As String) As String
        Dim retorno As String = ""
        If Not String.IsNullOrEmpty(strOriginal.ToString) Then
            retorno = String.Concat(
                        strOriginal.Where(
                            Function(c) "0123456789".Contains(c)))
        Else
            retorno = ""
        End If
        Return retorno
    End Function

    Public Function versaoSistema() As String
        Return Application.ProductVersion
    End Function

    Public Sub carregaDataGrid(frm As Form, dg As DataGridView, dt As DataTable)
        Try
            With frm
                With dg
                    .DataSource = dt
                End With
            End With
        Catch ex As Exception
        End Try
    End Sub

    ''' <summary>
    ''' Formata uma string de data no formato Date
    ''' Data criação/modificação: 17/04/2019
    ''' Formato de entrada: DD/MM/YYYY ou DDMMYYYY
    ''' </summary>  
    Public Function formataStringDataDDMMYYYY(strData As String) As Date
        Try

            Dim dtFormatada As Date
            'Dim milenio As Long = 2000
            strData = Replace(strData, "/", String.Empty)
            If strData.ToString <> "00000000" AndAlso strData.ToString <> "000" AndAlso Not String.IsNullOrEmpty(strData) Then
                dtFormatada = CDate(DateSerial(Microsoft.VisualBasic.Mid(strData, 5, 4), Microsoft.VisualBasic.Mid(strData, 3, 2), Microsoft.VisualBasic.Mid(strData, 1, 2)))
            Else : Return Nothing
            End If
            Return dtFormatada
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Formata uma string de data no formato Date
    ''' Data criação/modificação: 02/01/2018
    ''' Formato de entrada: 11/04/19
    ''' </summary>  
    Public Function formataStringDataDDMMYY(strData As String) As Date
        Try
            Dim dtFormatada As Date
            Dim milenio As Long = 2000
            Dim milenio_2 As Long = 1900
            strData = Replace(strData, "/", String.Empty)
            If strData.ToString <> "000000" AndAlso strData.ToString <> "000" AndAlso Not String.IsNullOrEmpty(strData) Then
                If Microsoft.VisualBasic.Mid(strData, 5, 2) >= 50 Then
                    dtFormatada = CDate(DateSerial(Microsoft.VisualBasic.Mid(strData, 5, 2) + milenio_2, Microsoft.VisualBasic.Mid(strData, 3, 2), Microsoft.VisualBasic.Mid(strData, 1, 2)))
                Else
                    dtFormatada = CDate(DateSerial(Microsoft.VisualBasic.Mid(strData, 5, 2) + milenio, Microsoft.VisualBasic.Mid(strData, 3, 2), Microsoft.VisualBasic.Mid(strData, 1, 2)))
                End If
            Else : Return Nothing
            End If
            Return dtFormatada
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Formata uma string de hora no formato Datetime
    ''' Data criação/modificação: 02/01/2018
    ''' </summary>  
    Public Function formataStringHoraHHMMSS(strHORA As String) As DateTime
        Try
            Dim milenio As Long = 2000
            Dim nmTime As New TimeSpan(Microsoft.VisualBasic.Mid(strHORA, 1, 2), Microsoft.VisualBasic.Mid(strHORA, 3, 2), Microsoft.VisualBasic.Mid(strHORA, 5, 2))
            Return Today.Add(nmTime).ToString("HH:mm:ss")
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Retorna um limite de caracteres especificos de uma string
    ''' Data criação/modificação: 02/01/2018
    ''' </summary>  
    Public Function retornaLimiteCaracterString(strValor As String, limite As Integer) As String
        Try
            Dim novoValor As String
            novoValor = Microsoft.VisualBasic.Mid(strValor, 1, limite)
            Return novoValor
        Catch ex As Exception
            Return strValor
        End Try
    End Function


    ''' <summary>
    ''' Autenticação via Active Directory
    ''' </summary>
    ''' <param name="User">usuario do dominio</param>
    ''' <param name="Senha">senha do dominio</param>
    ''' <param name="ValidaNovelRede">valida o usuário informado com usuário logado no computador, default = true</param>
    ''' <param name="IpServer">ip do dominio, opcional ou captura automatica da propria rede, default = auto</param>
    ''' <returns>Boolean</returns>
    Public Function autenticacaoActiveDirectory(ByVal User As String, ByVal Senha As String, Optional ByVal ValidaNovelRede As Boolean = True, Optional ByVal IpServer As String = "", Optional ByRef MsgErro As String = "", Optional ByVal exibeMsgErro As Boolean = True) As Boolean
        Dim resultado As String = ""
        Try
            '------------------------------------------------------------------------
            'verifica a necessidade de capturar automaticamente do ip dominio local
            Dim dominio As Domain
            Dim nomeDominio As String
            Dim directoryEntry As DirectoryEntry
            If String.IsNullOrEmpty(IpServer) Then
                dominio = Domain.GetCurrentDomain()
                nomeDominio = dominio.Name
            Else
                nomeDominio = IpServer.ToString
            End If
            '------------------------------------------------------------------------
            'autenticando...
            directoryEntry = New DirectoryEntry("LDAP://" + nomeDominio, User, Senha)
            resultado = directoryEntry.Name
            'validando autenticação
            If Not String.IsNullOrEmpty(resultado.ToString) Then
                If ValidaNovelRede Then
                    'Verifica se o usuário autenticado é o mesmo que esta logado no PC
                    If Environment.UserName.ToString.ToLower = User.ToString.ToLower Then
                        'MsgBox("Usuário Autenticado!", vbCritical)
                        Return True
                    Else
                        MsgErro = "Usuário logado no PC é diferente do usuário informado."
                        If exibeMsgErro Then
                            MsgBox("Erro na autenticação!" & vbNewLine & "O usuário logado neste computador" & vbNewLine & "é diferente do informado no login.", vbCritical, TITULO_ALERTA)
                        End If
                        Return False
                    End If
                Else
                    Return True 'autenticado
                End If
            Else
                MsgErro = "Usuário ou senha informado incorretamente!"
                If exibeMsgErro Then
                    MsgBox("Erro na autenticação!", vbCritical, TITULO_ALERTA)
                End If
                Return False
            End If
        Catch ex As Exception
            MsgErro = Err.Description & " - (" & Err.Number & ")"
            If exibeMsgErro Then
                MsgBox("Erro na autenticação!", vbCritical, TITULO_ALERTA)
            End If
            Return False
        End Try
    End Function

    'Public Sub abrirFormInPanel(frm As Form, painel As Panel, border As FormBorderStyle)
    '    Try
    '        If painel.Controls.Count > 0 Then
    '            painel.Controls.RemoveAt(0)
    '        End If
    '        frm.TopLevel = False
    '        frm.Dock = DockStyle.Fill
    '        frm.FormBorderStyle = border
    '        painel.Controls.Add(frm)
    '        painel.Tag = frm
    '        frm.Show()
    '    Catch ex As Exception
    '        MsgBox("Não foi possivel abrir o formulário!", vbCritical, TITULO_ALERTA)
    '    End Try
    'End Sub

    Public Sub fecharFormInPanel(sMDIPai As Form)
        Try
            'Fecha o formulario filho se tiver aberto
            If sMDIPai.ActiveMdiChild IsNot Nothing Then
                sMDIPai.ActiveMdiChild.Close()
            End If
        Catch ex As Exception
        End Try
    End Sub

    Public Sub abrirFormInPanelMDI(sMDIFilho As Form, sMDIPai As Form, sCtrlPainel As Control, Optional ByVal formBorder As FormBorderStyle = FormBorderStyle.FixedToolWindow)
        Try
            Dim frmFilho As New Form
            'Fecha o formulario filho se tiver aberto
            If sMDIPai.ActiveMdiChild IsNot Nothing Then
                If sMDIPai.ActiveMdiChild.Name <> sMDIFilho.Name Or sMDIPai.ActiveMdiChild.Name = sMDIFilho.Name Then
                    sMDIPai.ActiveMdiChild.Close()
                End If
            End If
            frmFilho = sMDIFilho
            'Configurações do formulario filho
            With frmFilho
                .MdiParent = sMDIPai
                .FormBorderStyle = formBorder
                .Dock = DockStyle.Fill
                .ShowInTaskbar = False
                .TopLevel = False
                .Show()
            End With

            'Limpa o controle
            'sCtrl.Controls.Clear()
            'Adiciona o formulario ao controle
            sCtrlPainel.Controls.Add(frmFilho)
            frmFilho.Activate()
            frmFilho.Focus()
        Catch ex As Exception
            MsgBox("Falha ao tentar abrir o formulário.", vbCritical, TITULO_ALERTA)
        End Try
    End Sub

    Public Function retornaFormularioPorString(nomeForm As String, assembly_name As String, root_namespace As String) As Form
        Try

            Dim minhaFrm As Form = Activator.CreateInstance(
            System.Reflection.Assembly.Load(assembly_name.ToString).GetType(root_namespace.ToString & "." & nomeForm))
            Return minhaFrm

            '    Dim returnObj As Object = Nothing
            '    Dim Type As Type = Assembly.GetExecutingAssembly().GetType("FRD." & objectName)
            '    If Type IsNot Nothing Then
            '        returnObj = Activator.CreateInstance(Type, args)
            '    End If
            '    Return returnObj


        Catch ex As Exception
            MsgBox("Falha ao tentar carregar o formulário!", vbCritical, TITULO_ALERTA)
            Return Nothing
        End Try
    End Function

    'Percorre todo o listview e marca os itens correspondentes
    Public Sub marcarCheckItensListview(lst As ListView, listaID As List(Of Object))
        Try
            If lst.Items.Count > 0 Then
                For Each item In lst.Items
                    'por padrão deixa todos desmarcados
                    item.ImageKey = 11
                    item.Checked = False
                    'Percorre toda a lista com os ID enviados para efetuar a marcação do check um a um
                    For Each lista In listaID
                        If item.SubItems.Item(0).Text.ToString = lista.ToString Then
                            item.ImageKey = 1
                            item.Checked = True
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            MsgBox("Falha ao tentar carregar a listagem!", vbCritical, TITULO_ALERTA)
        End Try
    End Sub

    Public Function capturaItensCheckListview(lst As ListView) As List(Of Object)
        Dim listaRetorno As New List(Of Object)
        Try
            If lst.Items.Count > 0 Then
                For Each item In lst.CheckedItems
                    listaRetorno.Add(item.SubItems.Item(0).Text.ToString)
                Next
            End If
            Return listaRetorno
        Catch ex As Exception
            Return listaRetorno
        End Try
    End Function

    'Função para abrir uma caixa de seleção de arquivos
    Public Function EnderecoArqCapturar() As String
        Dim open As New OpenFileDialog()
        Try
            If open.ShowDialog = DialogResult.OK Then
                Return open.FileName.ToString
            End If
        Catch ex As Exception
            MsgBox("Não foi possível identificar o endereço do arquivo. Motivo:" & ex.Message, MsgBoxStyle.SystemModal, TITULO_ALERTA)
        End Try
        Return String.Empty
    End Function


    Public Sub CursorPointer(bln As Boolean)
        If bln Then
            Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Else
            Cursor.Current = System.Windows.Forms.Cursors.Default
        End If
    End Sub

    'Public Function retornaDirPessoal(ByVal rede As FLAG_REDE) As String
    '    If rede = 1 Then
    '        retornaDirPessoal = "C:\Users\" & Environ("USERNAME") & "\Documents\"
    '    Else
    '        retornaDirPessoal = "\\" & Environ("COMPUTERNAME") & "\c$\Users\" & Environ("USERNAME") & "\"
    '    End If
    'End Function
    Public Function retornaDirPessoal(Optional algar_bradesco As String = "bradesco") As String

        retornaDirPessoal = "C:\Users\" & Environ("USERNAME") & "\Documents\"

        'Decomissionado em 20/10/19
        'If algar_bradesco.ToString.ToLower = "algar" Then
        '    retornaDirPessoal = "C:\Users\" & Environ("USERNAME") & "\Documents\"
        'Else
        '    retornaDirPessoal = "\\" & Environ("COMPUTERNAME") & "\c$\Users\" & Environ("USERNAME") & "\"
        'End If
    End Function

    Public Function properCase(ByVal texto As String, ByVal Optional nroLimiteCaracteres As Integer = 100) As String
        Dim str As String = ""
        str = nomeProprio(StrConv(texto.ToString.ToLower, vbProperCase))
        str = IIf(str.ToString.Length > nroLimiteCaracteres, Left(str.ToString, nroLimiteCaracteres) & "...", str.ToString)
        Return str
    End Function


    ''' <summary>
    ''' Método que formata para moeda o conteúdo de um TextBox
    ''' Para fazer uma chamada ao método TextBoxMoeda da classe Utils, 
    ''' adiciona a linha abaixo no evento TextChanged do controle TextBox.
    ''' </summary>
    ''' <param name="txt">Controle a ser formatado</param>
    Public Sub TextBoxMoeda(ByVal txt As TextBox)
        Dim n As String = String.Empty
        Dim v As Double = 0
        Try
            n = txt.Text.Replace(",", "").Replace(".", "")
            If n.Equals("") Then n = "000"
            n = n.PadLeft(3, "0")
            If n.Length > 3 And n.Substring(0, 1) = "0" Then n = n.Substring(1, n.Length - 1)
            v = Convert.ToDouble(n) / 100
            txt.Text = String.Format("{0:N}", v)
            txt.SelectionStart = txt.Text.Length
        Catch ex As Exception
            MessageBox.Show(ex.Message, "TextBoxMoeda")
        End Try
    End Sub

    'função para retornar caminho/nome do arquivo onde devemos "salvar Como"
    Public Function SalvarComo(Optional ByVal nomeArquivo As String = "") As String
        Dim saveFileDialog1 As New SaveFileDialog()
        With saveFileDialog1
            .Filter = "txt files (*.txt)|*.txt|csv files (*.csv)|*.csv|All files (*.*)|*.*"
            .Title = "Salvar arquivo em..."
            '.InitialDirectory = nomeArquivo
            .FileName = nomeArquivo
            '.ShowDialog()
            If saveFileDialog1.ShowDialog() = DialogResult.OK Then
                If .FileName <> "" Then
                    Return .FileName
                Else
                    Return ""
                End If
            Else
                Return ""
            End If
        End With
    End Function
    Public Sub Campos_Habilitar(ByVal Tela As Control, Optional bloquear As Boolean = False)
        'Caso ocorra erro, não mostrar o erro, ignorando e indo para á próxima linha
        On Error Resume Next
        'Declaramos uma variavel Campo do tipo Object
        '(Tipo Object porque iremos trabalhar com todos os campos do Form, podendo ser
        '       Label, Button, TextBox, ComboBox e outros)
        Dim Campo As Object
        'Usaremos For Each para passarmos por todos os controls do objeto atual
        For Each Campo In Tela.Controls

            If Not TypeOf Campo Is Label And
                Not TypeOf Campo Is MenuStrip And
                Not TypeOf Campo Is PictureBox And
                Not TypeOf Campo Is StatusStrip Then
                Campo.Enabled = bloquear
            End If

        Next

    End Sub
    Public Sub Campos_SomenteLeitura(ByVal Tela As Control, Optional Ativar As Boolean = False)
        'Caso ocorra erro, não mostrar o erro, ignorando e indo para á próxima linha
        On Error Resume Next
        'Declaramos uma variavel Campo do tipo Object
        '(Tipo Object porque iremos trabalhar com todos os campos do Form, podendo ser
        '       Label, Button, TextBox, ComboBox e outros)
        Dim Campo As Object
        'Usaremos For Each para passarmos por todos os controls do objeto atual
        For Each Campo In Tela.Controls

            If TypeOf Campo Is TextBox Or
                TypeOf Campo Is MaskedTextBox Then
                Campo.ReadOnly = Ativar
            End If

        Next

    End Sub

    Public Function exportarListViewParaExcel(ByVal ListView As ListView) As Boolean
        Try
            Dim objExcel As New excel.Application
            Dim bkWorkBook As excel.Workbook
            Dim shWorkSheet As excel.Worksheet
            Dim i As Integer
            Dim j As Integer
            objExcel = New excel.Application
            bkWorkBook = objExcel.Workbooks.Add
            shWorkSheet = CType(bkWorkBook.ActiveSheet, excel.Worksheet)
            shWorkSheet.Cells().NumberFormat = "@" 'Formatando as celulas como texto
            For i = 0 To ListView.Columns.Count - 1
                shWorkSheet.Cells(1, i + 1) = ListView.Columns(i).Text
            Next
            For i = 0 To ListView.Items.Count - 1
                For j = 0 To ListView.Items(i).SubItems.Count - 1
                    shWorkSheet.Cells(i + 2, j + 1) = ListView.Items(i).SubItems(j).Text.ToString
                Next
            Next
            objExcel.Visible = True
            Return True
        Catch ex As Exception
            MsgBox(ex.Message)
            Return False
        End Try
    End Function

    'AutoCloseMsgBox "Msgbox1 - Clique em OK ou aguarde 2 segundos", "Fechar MsgBox1 automaticamente", 2 
    '2 segundos
    Sub AutoCloseMsgBox(Mensagem As String, Titulo As String, Segundos As Integer)
        Dim oSHL As Object
        oSHL = CreateObject("WScript.Shell")
        oSHL.PopUp(Mensagem, Segundos, Titulo, vbOKOnly + vbInformation)
    End Sub

    Public Sub carregaBarraProgresso(ByVal frm As Form, ByVal nomeBarraProgresso As ProgressBar, Optional maximo As Integer = 0, Optional limpeza As Boolean = True, Optional saltoProgresso As Boolean = False)
        With frm
            With nomeBarraProgresso
                If Not saltoProgresso Then
                    .Maximum = maximo
                    .Minimum = 0
                    .Step = 1
                    .Value = 0
                    .Visible = IIf(limpeza, False, True)
                Else
                    .PerformStep()
                    Application.DoEvents()
                End If
            End With
        End With
    End Sub
    Public Function CriarCopiarMoverDeletarAquivo(ByVal caminhoArquivo As String, ByVal acao As String, Optional NovoCaminho As String = "") As Boolean

        Dim arq As New FileInfo(caminhoArquivo)
        Try
            Select Case acao
                Case "Criar"
                    arq.Create()
                Case "Copiar"
                    arq.CopyTo(NovoCaminho)
                Case "Mover"
                    arq.MoveTo(NovoCaminho)
                Case "Deletar"
                    arq.Delete()
            End Select
            Return True
        Catch ex As Exception
            Return False
            MsgBox("Ação: " & acao & " não realizada, tente novamente!", vbInformation, TITULO_ALERTA)
        End Try

    End Function


    Public Function exibeMensagemDoDia() As String
        'Imports System
        'Imports System.Globalization
        'Imports System.Threading
        'Dim dt As DateTime = DateTime.Now
        '' Sets the CurrentCulture property to U.S. English.
        'Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
        '' Displays dt, formatted using the ShortDatePattern
        '' and the CurrentThread.CurrentCulture.
        'Console.WriteLine(dt.ToString("d"))
        Dim dtcurrent As System.DateTime
        Dim ihour As Integer
        Dim dia As String = "Bom dia"
        Dim noite As String = "Boa noite"
        Dim tarde As String = "Boa tarde"
        dtcurrent = Date.Now()
        ihour = dtcurrent.Hour
        If (ihour < 12) Then
            Return dia
            Exit Function
        End If
        If (ihour >= 12) And (ihour < 18) Then
            Return tarde
            Exit Function
        End If
        If (ihour >= 18) Then
            Return noite
            Exit Function
        End If
        Return dia
    End Function

    Public Function exportarRS_excel(ByVal rst As ADODB.Recordset) As Boolean

        Try

            CursorPointer(True)
            If IsNothing(rst) Then
                MsgBox("Por favor, atualize o relatório antes de exportar para o Excel!", vbInformation, TITULO_ALERTA)
                Return False
            End If

            'Define variáveis para o Excel: 
            'Application(Excel) 
            'WorkBook(pasta de trabalho)
            'WorkSheet(Planilha)
            'Range(intervalo de células)
            Dim xlapp As New excel.Application
            Dim xlwbook As excel.Workbook = xlapp.Workbooks.Add(excel.XlWBATemplate.xlWBATWorksheet)
            Dim xlwsheet As excel.Worksheet = CType(xlwbook.Worksheets(1), excel.Worksheet)
            Dim xlrange As excel.Range = xlwsheet.Range("A1")
            'define modo de cálculo : 
            'xlCalculationAutomatic
            'xlCalculationManual
            'xlCalculationSemiautomatic
            Dim xlcalc As excel.XlCalculation
            'Outras variáveis
            Dim contadorCampos As Integer = Nothing

            'Desabilita temporariamente o calculo automatico passando para manual
            With xlapp
                xlcalc = .Calculation
                .Calculation = excel.XlCalculation.xlCalculationManual
            End With

            'Escreve os nomes dos campos na planilha
            For Each fld In rst.Fields
                xlrange.Offset(0, contadorCampos).Value = fld.Name
                xlrange.Offset(0, contadorCampos).Interior.Color = RGB(173, 216, 230) 'Background do cabeçalho
                xlrange.Offset(0, contadorCampos).Font.Bold = True 'Negrito
                contadorCampos = contadorCampos + 1
            Next

            xlrange.Offset(1, 0).CopyFromRecordset(rst) 'Copia o recordset para a planilha
            rst.Close() 'Fecha o recordset.

            With xlapp
                .Visible = True
                .UserControl = True
                .Calculation = xlcalc 'Restaura o calculo
            End With

            'libera os objetos
            rst = Nothing
            xlrange = Nothing
            xlwsheet = Nothing
            xlwbook = Nothing
            xlapp = Nothing
            Return True

        Catch ex As Exception
            MessageBox.Show("Ocorreu um erro : " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try


    End Function


    Public Sub killProcessoName(nameProcesso As String)
        'Mata todos os processos do nome especificado
        Dim Processos() As Process = System.Diagnostics.Process.GetProcessesByName(nameProcesso.ToString)
        For Each x As Process In Processos
            x.Kill()
        Next
    End Sub
    Public Sub killProcessoNameID(nameProcesso As String, ID As String)
        'Mata o processos de um nome e HashCode especifico 
        Dim Processos() As Process = System.Diagnostics.Process.GetProcessesByName(nameProcesso.ToString)
        For Each x As Process In Processos
            If x.Id.ToString = ID.ToString Then
                x.Kill()
            End If
        Next
    End Sub
    Public Function IDProcessByName(nameProcesso As String, Optional IDExcecao As String = "") As String
        'Retorna o HashCode do processo em especifico
        Dim Processos() As Process = System.Diagnostics.Process.GetProcessesByName(nameProcesso.ToString)
        Dim hashCode As String = String.Empty
        For Each x As Process In Processos
            If IDExcecao <> x.Id.ToString Then
                hashCode = x.Id.ToString
            End If
        Next
        Return hashCode.ToString
    End Function

    'Gera um GUID
    Public Function GenerateGUID() As String
        Try
            Dim sGUID As String
            sGUID = System.Guid.NewGuid.ToString()
            Return sGUID
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function

    Public Function CriaMascaraCartao(strCartao As String) As String
        Try
            Dim cartaoNew As String = String.Empty
            If Microsoft.VisualBasic.Len(strCartao.Trim) = 15 Then
                cartaoNew = strCartao.Replace(Mid(strCartao, 5, 7), "*******")
            Else
                cartaoNew = strCartao.Replace(Mid(strCartao, 5, 8), "********")
            End If
            Return cartaoNew
        Catch ex As Exception
            Return strCartao
        End Try
    End Function


    Public Function RealocarAquivo(tipo As FLAG_ARQUIVO, arquivo As String, PathOrigem As String, PathDestino As String) As Boolean
        Dim thisPath As String = String.Empty
        Dim thisPathDestino As String = String.Empty
        Try
            thisPath = LCase(PathOrigem & arquivo)
            thisPathDestino = LCase(PathDestino & arquivo)
            If Not String.IsNullOrEmpty(thisPath) And Not String.IsNullOrEmpty(thisPathDestino) Then
                Select Case tipo
                    Case FLAG_ARQUIVO.Copy
                        File.Copy(thisPath, thisPathDestino)
                        Return True
                    Case FLAG_ARQUIVO.Move
                        File.Move(thisPath, thisPathDestino)
                        Return True
                    Case FLAG_ARQUIVO.Delete
                        File.Delete(thisPath)
                        Return True
                End Select
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' Converter DataJson em Datetime
    Public Function UnixToTime(ByVal strUnixTime As String) As DateTime
        Try
            Dim result As String
            Dim dt As New DateTime(1970, 1, 1)
            dt = dt.AddMilliseconds(Long.Parse(strUnixTime))
            dt = dt.ToLocalTime()
            result = dt.ToString("yyyy-MM-dd HH:mm:ss")
            Return result
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function FormatoDataUniversal() As String
        Return "yyyy-MM-dd"
    End Function
    Public Function FormatoDataHoraUniversal() As String
        Return "yyyy-MM-dd HH:mm:ss"
    End Function

    Public Function PreencheListViewColunasDinamicas(ByVal lst As ListView,
                                                        ByVal dt As DataTable) As ListView
        Try

            Dim str(dt.Columns.Count) As String
            lst.Clear()
            With lst
                .View = View.Details
                .LabelEdit = False
                .CheckBoxes = False
                '.SmallImageList = imglist()
                .GridLines = True
                .FullRowSelect = True
                .HideSelection = False
                .MultiSelect = False

                'colunas dinamicas da tabela
                For Each column As DataColumn In dt.Columns
                    .Columns.Add(column.ColumnName, 125, HorizontalAlignment.Left)
                Next

            End With
            'POPULANDO
            If dt.Rows.Count > 0 Then
                For Each drRow As DataRow In dt.Rows
                    'nome das colunas dinamicamente
                    For col As Integer = 0 To dt.Columns.Count - 1
                        str(col) = drRow(col).ToString()
                    Next
                    Dim ii As New ListViewItem(str)
                    lst.Items.Add(ii)
                Next drRow
            End If
            Return lst
        Catch ex As Exception
            Return lst
        End Try
    End Function

    Public Function capturaFinalCartao(cartao As String) As String
        Try
            If cartao.Length = 15 Then
                Return Microsoft.VisualBasic.Right(cartao, 9)
            ElseIf cartao.Length = 16 Then
                Return Microsoft.VisualBasic.Right(cartao, 10)
            Else
                Return String.Empty
            End If
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function
    Public Function formataCartao19Posicoes(cartao As String) As String
        Try
            Dim novoCard As String
            If cartao.Length = 15 Then 'amex
                novoCard = Microsoft.VisualBasic.Right("0000" & cartao, 19)
                Return novoCard.ToString
            ElseIf cartao.Length = 16 Then  'visa,master,elo
                novoCard = Microsoft.VisualBasic.Right("000" & cartao, 19)
                Return novoCard.ToString
            Else
                Return String.Empty
            End If
        Catch ex As Exception
            Return cartao
        End Try
    End Function
    Public Function validaDataMembersice(strData As String) As Date
        'sempre supondo que a data com ano de quatro dígitos deva estar entre os últimos 70 anos.
        Try
            Dim dtAbertura As Date
            Dim dia As String = ""
            Dim mes As String = ""
            Dim ano As String = ""
            Dim ano_ As String = ""
            Dim data() As String
            If Not String.IsNullOrEmpty(strData.ToString) Then
                '31/07/14
                data = Split(strData, "/")
                dia = data(0).ToString
                mes = data(1).ToString
                ano = data(2).ToString
                ano_ = Microsoft.VisualBasic.Right(Year(Now.AddYears(-70)).ToString, 2)
                If Int(ano) >= Int(ano_) Then
                    dtAbertura = CDate(DateSerial("19" & ano.ToString, mes.ToString, dia.ToString)).Date
                Else
                    dtAbertura = CDate(DateSerial("20" & ano.ToString, mes.ToString, dia.ToString)).Date
                End If
            Else
                dtAbertura = Nothing
            End If
            Return dtAbertura
        Catch ex As Exception
            Return Nothing
        End Try
    End Function


End Class