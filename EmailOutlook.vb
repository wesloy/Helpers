Imports Microsoft.Office.Interop
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports System.IO

Public Class EmailOutlook

#Region "Propriedades"

    Dim _guidEmail As String
    Dim _data As DateTime
    Dim _remetente As String
    Dim _destinatarios As String
    Dim _cCopiaPara As String
    Dim _assunto As String
    Dim _corpoEmail As String
    Dim _anexo As New List(Of String)
    Dim _arquivoDeEmail As String
    Dim _EmailConta As String
    Dim _EmailPastaEntrada As String
    Dim _emailPastaBackup As String
    Dim _pathAnexo As String
    Dim _salvarAnexosDoEmail As Boolean = False
    Dim _salvarEmailComoAnexo As Boolean = False
    Dim _prefixoArquivo As String = String.Empty
    Dim _bodyHTML As Boolean = False

    Public Property GuidEmail As String
        Get
            Return _guidEmail
        End Get
        Set(value As String)
            _guidEmail = value
        End Set
    End Property

    Public Property Data As Date
        Get
            Return _data
        End Get
        Set(value As Date)
            _data = value
        End Set
    End Property

    Public Property Remetente As String
        Get
            Return _remetente
        End Get
        Set(value As String)
            _remetente = value
        End Set
    End Property

    Public Property Destinatarios As String
        Get
            Return _destinatarios
        End Get
        Set(value As String)
            _destinatarios = value
        End Set
    End Property

    Public Property CCopiaPara As String
        Get
            Return _cCopiaPara
        End Get
        Set(value As String)
            _cCopiaPara = value
        End Set
    End Property

    Public Property Assunto As String
        Get
            Return _assunto
        End Get
        Set(value As String)
            _assunto = value
        End Set
    End Property

    Public Property CorpoEmail As String
        Get
            Return _corpoEmail
        End Get
        Set(value As String)
            _corpoEmail = value
        End Set
    End Property

    Public Property Anexo As List(Of String)
        Get
            Return _anexo
        End Get
        Set(value As List(Of String))
            _anexo = value
        End Set
    End Property

    Public Property ArquivoDeEmail As String
        Get
            Return _arquivoDeEmail
        End Get
        Set(value As String)
            _arquivoDeEmail = value
        End Set
    End Property

    Public Property EmailConta As String
        Get
            Return _EmailConta
        End Get
        Set(value As String)
            _EmailConta = value
        End Set
    End Property

    Public Property EmailPastaEntrada As String
        Get
            Return _EmailPastaEntrada
        End Get
        Set(value As String)
            _EmailPastaEntrada = value
        End Set
    End Property

    Public Property EmailPastaBackup As String
        Get
            Return _emailPastaBackup
        End Get
        Set(value As String)
            _emailPastaBackup = value
        End Set
    End Property

    Public Property PathAnexo As String
        Get
            Return _pathAnexo
        End Get
        Set(value As String)
            _pathAnexo = value
        End Set
    End Property

    Public Property SalvarAnexosDoEmail As Boolean
        Get
            Return _salvarAnexosDoEmail
        End Get
        Set(value As Boolean)
            _salvarAnexosDoEmail = value
        End Set
    End Property

    Public Property SalvarEmailComoAnexo As Boolean
        Get
            Return _salvarEmailComoAnexo
        End Get
        Set(value As Boolean)
            _salvarEmailComoAnexo = value
        End Set
    End Property

    Public Property BodyHTML As Boolean
        Get
            Return _bodyHTML
        End Get
        Set(value As Boolean)
            _bodyHTML = value
        End Set
    End Property

    Public Property PrefixoArquivo As String
        Get
            Return _prefixoArquivo
        End Get
        Set(value As String)
            _prefixoArquivo = value
        End Set
    End Property



#End Region

    'MODELO DE HTML P/ BODY
    'HTMLBody = "<html>"
    'HTMLBody += "<head><title>Testing</title></head>"
    'HTMLBody += "<style type=""text/css"">"
    'HTMLBody += "table, th, td {border: 1px solid #ccc;border-collapse: collapse;} "
    'HTMLBody += "body {font-size: 12px;} "
    'HTMLBody += "</style>"
    'HTMLBody += "<body>"
    'HTMLBody += "<table>"
    'HTMLBody += "<tr>"
    'HTMLBody += "<td><Label><strong>Test Text:</strong></Label></td>"
    'HTMLBody += "<td width=""600""><Label id=""protocolo"" >xxx</Label></td>"
    'HTMLBody += "</tr>"
    'HTMLBody += "</table>"
    'HTMLBody += "</body>"
    'HTMLBody += "</html>"
#Region "Metodos"

    Public Function Enviar(ByVal objEmail As EmailOutlook) As Boolean

        Try
            IniciaProcessoOutlook()
            Dim oApp As New Outlook.Application
            Dim oMail As Outlook.MailItem
            Dim hlp As New Helpers
            Dim extensaoEmail As String = ".msg"
            Dim nomeArquivo As String = ""

            oMail = oApp.CreateItem(Outlook.OlItemType.olMailItem)
            oMail.UnRead = True
            oMail.SentOnBehalfOfName = objEmail.Remetente
            oMail.To = objEmail.Destinatarios
            oMail.CC = objEmail.CCopiaPara
            oMail.Subject = objEmail.Assunto
            'anexando arquivos ao email
            If objEmail.Anexo IsNot Nothing And objEmail.Anexo.Count > 0 Then
                For i = 0 To objEmail.Anexo.Count
                    oMail.Attachments.Add(objEmail.Anexo(i)) 'p/anexos
                Next
            End If

            If objEmail.BodyHTML Then
                ''HTML BODY
                oMail.BodyFormat = Outlook.OlBodyFormat.olFormatHTML
                oMail.HTMLBody = objEmail.CorpoEmail
            Else
                oMail.Body = objEmail.CorpoEmail
            End If

            'salva o arquivo de email como backup
            If objEmail.SalvarEmailComoAnexo _
                And Not String.IsNullOrEmpty(objEmail.GuidEmail.ToString) _
                And Not String.IsNullOrEmpty(objEmail.PathAnexo.ToString) Then
                nomeArquivo = objEmail.PathAnexo.ToString & objEmail.GuidEmail.ToString & objEmail.PrefixoArquivo.ToString & extensaoEmail.ToString
                objEmail.ArquivoDeEmail = nomeArquivo
                oMail.SaveAs(nomeArquivo, Microsoft.Office.Interop.Outlook.OlSaveAsType.olMSG)
            End If

            'para enviar o email
            oMail.Send()
            'para visualizar o email antes de enviar
            'oMail.Display()
            ' Clean up.
            'oApp.Quit()
            oApp = Nothing
            oMail = Nothing
            Return True
            Exit Function
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function Ler(conta_email As String, pasta_entrada As String, pasta_backup As String, capturaBodyHTML As Boolean, salvarAnexos_Email As Boolean, salvarArquivoDeEmailMsg As Boolean, strPathAnexo As String, Optional prefixoNomeArquivo As String = "") As List(Of EmailOutlook)
        Dim listaDeEmails As New List(Of EmailOutlook)
        Dim objEmail As EmailOutlook

        Try
            IniciaProcessoOutlook()
            Dim oApp As New Outlook.Application
            Dim oNS As Outlook.NameSpace
            Dim oInbox As Outlook.MAPIFolder
            Dim oItems As Outlook.Items
            Dim newMail As Microsoft.Office.Interop.Outlook.MailItem

            Dim cont As Long = 0
            Dim totalEmail As Long = 0
            Dim nomeAnexo As String = ""
            Dim GuidEmail As String = ""
            Dim extensaoEmail As String = ".msg"
            Dim hlp As New Helpers

            oApp = CreateObject("Outlook.Application")
            'Cria a instancia Namespace do outlook
            oNS = oApp.GetNamespace("mapi")
            'Captura as mensagem de uma conta/pasta especifica dentro do outlook
            'Dim oInbox As Outlook.MAPIFolder = tempApp.Folders(conta).Folders(pasta_entrada).Folders(subPastas)
            oInbox = oNS.Folders(conta_email).Folders(pasta_entrada)
            oItems = oInbox.Items
            ' Filtra somente os e-mails não lidos
            oItems = oItems.Restrict("[Unread] = true")
            cont = oItems.Count
            totalEmail = oItems.Count

            'For Each newMail In oItems
            Do Until cont = 0
                newMail = oItems(cont)
                objEmail = New EmailOutlook
                objEmail.EmailConta = conta_email
                objEmail.EmailPastaEntrada = pasta_entrada
                objEmail.EmailPastaBackup = pasta_backup
                objEmail.BodyHTML = capturaBodyHTML
                objEmail.SalvarAnexosDoEmail = salvarAnexos_Email
                objEmail.SalvarEmailComoAnexo = salvarArquivoDeEmailMsg
                objEmail.PathAnexo = strPathAnexo
                objEmail.PrefixoArquivo = prefixoNomeArquivo

                'Captura as informações do email
                If newMail.UnRead Then

                    '00 - Setar não confirmação de leitura
                    newMail.ReadReceiptRequested = False

                    '01 - Captura o email do remetente  
                    If newMail.SenderEmailType = "EX" Then
                        Dim sender As Outlook.AddressEntry
                        sender = newMail.Sender 'oItems.Sender
                        If Not sender Is Nothing Then
                            'Now we have an AddressEntry representing the Sender
                            If sender.AddressEntryUserType = Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry Or sender.AddressEntryUserType = Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry Then
                                'Use the ExchangeUser object PrimarySMTPAddress
                                Dim exchUser As Outlook.ExchangeUser
                                exchUser = sender.GetExchangeUser()
                                If Not exchUser Is Nothing Then
                                    objEmail.Remetente = exchUser.PrimarySmtpAddress
                                Else
                                    objEmail.Remetente = vbNullString
                                End If
                            Else
                                objEmail.Remetente = vbNullString 'sender.PropertyAccessor.GetProperty("PR_SMTP_ADDRESS")
                            End If
                        Else
                            objEmail.Remetente = vbNullString
                        End If
                    Else
                        objEmail.Remetente = newMail.SenderEmailAddress
                    End If


                    '02 - Captura o assunto
                    If IsNothing(newMail.Subject) = False Then
                        objEmail.Assunto = newMail.Subject.Trim
                        'retira marca de responda e encaminhamento do assunto do email
                        objEmail.Assunto = objEmail.Assunto.Replace("'", "")
                        objEmail.Assunto = objEmail.Assunto.Replace("RE:", "")
                        objEmail.Assunto = objEmail.Assunto.Replace("RES:", "")
                        objEmail.Assunto = objEmail.Assunto.Replace("ENC:", "")
                        objEmail.Assunto = objEmail.Assunto.Replace("FW:", "")
                        objEmail.Assunto = objEmail.Assunto.Trim
                    Else
                        objEmail.Assunto = String.Empty
                    End If


                    '03 - Outras informações
                    objEmail.Data = newMail.ReceivedTime
                    If objEmail.BodyHTML Then
                        'newMail.BodyFormat = Outlook.OlBodyFormat.olFormatHTML
                        objEmail.CorpoEmail = newMail.HTMLBody
                    Else
                        'newMail.BodyFormat = Outlook.OlBodyFormat.olFormatRichText
                        objEmail.CorpoEmail = newMail.Body
                    End If


                    '04 - Salvo os anexos do email
                    If objEmail.SalvarAnexosDoEmail Then
                        If (newMail.Attachments.Count > 0) Then
                            For Each newAttchments In newMail.Attachments
                                nomeAnexo = hlp.GenerateGUID() & " " & newAttchments.FileName.ToString()
                                newAttchments.SaveAsFile(objEmail.PathAnexo & nomeAnexo)
                                objEmail.Anexo.Add(objEmail.PathAnexo & nomeAnexo)
                            Next
                        End If
                    End If

                    '05 - Salvar o arquivo de email completo com a extensão .msg no patch de anexo
                    'cria um guid unico
                    If objEmail.SalvarEmailComoAnexo _
                        And Not String.IsNullOrEmpty(objEmail.PathAnexo.ToString) Then

                        'Localizar o guid do email que esta sendo lido
                        'a busca é feita pelo Assunto e pelo body do email
                        GuidEmail = LocalizaGuidEmString(objEmail.Assunto & objEmail.CorpoEmail)
                        objEmail.GuidEmail = GuidEmail.ToString

                        'cria um guid somente se nao encontrar no subject
                        If String.IsNullOrEmpty(GuidEmail.ToString) Then
                            GuidEmail = hlp.GenerateGUID
                            objEmail.GuidEmail = GuidEmail.ToString
                        End If

                        nomeAnexo = objEmail.PathAnexo.ToString & GuidEmail.ToString & objEmail.PrefixoArquivo.ToString & extensaoEmail.ToString
                        objEmail.ArquivoDeEmail = nomeAnexo
                        newMail.SaveAs(nomeAnexo, Microsoft.Office.Interop.Outlook.OlSaveAsType.olMSG)
                    End If

                    '06 - Marcar como lido
                    newMail.UnRead = False
                    newMail.Save()

                    '07 - Mover email para outra pasta
                    Dim destFolder As Outlook.MAPIFolder = oNS.Folders(objEmail.EmailConta).Folders(objEmail.EmailPastaEntrada).Folders(objEmail.EmailPastaBackup)
                    newMail.Move(destFolder)
                    'newMail = Nothing

                    'adiciona na lista
                    listaDeEmails.Add(objEmail)
                End If

                cont = cont - 1 'pula para o proximo email
                'newMail.Close(Outlook.OlInspectorClose.olDiscard)
                newMail = Nothing
            Loop

            'oApp.Quit()
            oApp = Nothing
            oNS = Nothing
            oItems = Nothing
            oInbox = Nothing
            newMail = Nothing
            Return listaDeEmails
        Catch ex As Exception
            Return listaDeEmails
        End Try
    End Function
    Public Function ResponderAquivoEmail(guid As String, ArquivoEmail As String, path As String, mensagemResposta As String, EmailRemetente As String, Optional prefixoNomeArquivo As String = "") As Boolean

        Try
            IniciaProcessoOutlook()
            Dim Reply As Microsoft.Office.Interop.Outlook.MailItem
            Dim newMail As Microsoft.Office.Interop.Outlook.MailItem
            Dim oApp As New Outlook.Application

            Dim cont As Long = 0
            Dim totalEmail As Long = 0
            Dim nomeAnexo As String = ""
            Dim GuidEmail As String = ""
            Dim extensaoEmail As String = ".msg"
            Dim hlp As New Helpers
            Dim replyOk As Boolean = False
            Dim thisFile As String = String.Empty
            Dim thisPath As String = String.Empty
            Dim objEmail As EmailOutlook

            thisFile = LCase(Dir(path & ArquivoEmail))
            thisPath = LCase(path & ArquivoEmail)

            If Not String.IsNullOrEmpty(thisFile) Then
                'Abre o arquivo de e-mail
                'Reply = oApp.Session.OpenSharedItem(thisPath)
                oApp = CreateObject("Outlook.Application")
                'newMail = oApp.GetNamespace("mapi").OpenSharedItem(thisPath)
                newMail = oApp.Session.OpenSharedItem(thisPath)

                objEmail = New EmailOutlook
                objEmail.SalvarEmailComoAnexo = True
                objEmail.PathAnexo = path
                objEmail.Data = newMail.ReceivedTime
                objEmail.PrefixoArquivo = prefixoNomeArquivo
                objEmail.GuidEmail = guid

                '01 - Captura o email do remetente  
                If newMail.SenderEmailType = "EX" Then
                    Dim sender As Outlook.AddressEntry
                    sender = newMail.Sender 'oItems.Sender
                    If Not sender Is Nothing Then
                        'Now we have an AddressEntry representing the Sender
                        If sender.AddressEntryUserType = Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry Or sender.AddressEntryUserType = Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry Then
                            'Use the ExchangeUser object PrimarySMTPAddress
                            Dim exchUser As Outlook.ExchangeUser
                            exchUser = sender.GetExchangeUser()
                            If Not exchUser Is Nothing Then
                                objEmail.Remetente = exchUser.PrimarySmtpAddress
                            Else
                                objEmail.Remetente = vbNullString
                            End If
                        Else
                            objEmail.Remetente = vbNullString 'sender.PropertyAccessor.GetProperty("PR_SMTP_ADDRESS")
                        End If
                    Else
                        objEmail.Remetente = vbNullString
                    End If
                Else
                    objEmail.Remetente = newMail.SenderEmailAddress
                End If


                '01 - Cria e-mail de resposta
                Reply = newMail.Reply
                Reply.SentOnBehalfOfName = EmailRemetente
                Reply.BodyFormat = Outlook.OlBodyFormat.olFormatHTML
                Reply.HTMLBody = "<p>" & mensagemResposta.ToString & "</p>"

                'Formata rementente anterior
                Reply.HTMLBody += "<hr></hr>"
                Reply.HTMLBody += "<strong>De: </strong>" & objEmail.Remetente & "<br></br>"
                Reply.HTMLBody += "<strong>Enviada em: </strong>" & objEmail.Data & "<br></br>"


                'Append mensagem anterior
                Reply.HTMLBody += newMail.HTMLBody

                '02 - Salvar o arquivo de email completo com a extensão .msg no patch de anexo
                'com um prefixo de "_reply"
                'cria um guid unico
                If objEmail.SalvarEmailComoAnexo _
                    And Not String.IsNullOrEmpty(objEmail.PathAnexo.ToString) _
                    And Not String.IsNullOrEmpty(objEmail.GuidEmail.ToString) Then
                    nomeAnexo = objEmail.PathAnexo.ToString & objEmail.GuidEmail.ToString & objEmail.PrefixoArquivo.ToString & extensaoEmail.ToString
                    objEmail.ArquivoDeEmail = nomeAnexo
                    Reply.SaveAs(nomeAnexo, Microsoft.Office.Interop.Outlook.OlSaveAsType.olMSG)
                End If

                '03 - Envia a resposta
                Reply.Send()
                replyOk = True
                thisFile = Dir()



                'força o fechamento caso tenha algum anexo
                Dim attachs As Outlook.Attachments = newMail.Attachments
                Dim z As Integer = attachs.Count
                System.Runtime.InteropServices.Marshal.ReleaseComObject(attachs)
                attachs = Nothing
                'fecha o email original
                newMail.Close(Microsoft.Office.Interop.Outlook.OlInspectorClose.olDiscard)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(newMail)
                'fecha o email criado 
                'Reply.Close(Microsoft.Office.Interop.Outlook.OlInspectorClose.olDiscard)
            End If

            'oApp.Quit()
            oApp = Nothing
            Reply = Nothing
            newMail = Nothing

            Return replyOk
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Sub IniciaProcessoOutlook()
        Try
            If Process.GetProcessesByName("outlook").Count = 0 Then
                Dim myProcess As New Process()
                myProcess = System.Diagnostics.Process.Start("outlook")
            End If
        Catch ex As Exception
            'MsgBox(hlp.getCurrentMethodName & " " & ex.Message, vbCritical, TITULO_ALERTA)
        End Try
    End Sub
    Public Function GetHtmlDoc(html As String) As System.Windows.Forms.HtmlDocument
        Try
            Dim browser As New WebBrowser
            browser.ScriptErrorsSuppressed = True
            browser.DocumentText = html
            browser.Document.OpenNew(True)
            browser.Document.Write(html)
            browser.Refresh()
            Return browser.Document
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function GetElementHtmlByID(idTagHTML As String, strHTML As String) As String
        Try
            Dim doc As HtmlDocument = Nothing
            doc = GetHtmlDoc(strHTML)
            If doc IsNot Nothing Then
                Return doc.GetElementById(idTagHTML.ToString).InnerText.ToString
            Else
                Return String.Empty
            End If
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function
    Public Function removeTags(ByVal HTML As String) As String
        Try
            Dim retorno As String
            'Remove as tags do HTML
            retorno = Regex.Replace(HTML, "<[^>]*>", String.Empty)
            retorno = Regex.Replace(retorno, "<(.|\n)*?>", String.Empty)
            retorno = Regex.Replace(retorno, "<[^<>]+>", String.Empty)
            retorno = Regex.Replace(retorno, "<.*?>", String.Empty)
            retorno = Regex.Replace(retorno, "&#8211;", " ")
            retorno = Regex.Replace(retorno, "\&nbsp;", " ")
            retorno = retorno.Replace(vbCrLf, String.Empty).Replace(vbTab, String.Empty)
            'Return Regex.Replace(HTML, "<.*?>", String.Empty)
            Return retorno.ToString
        Catch ex As Exception
            Return String.Empty
            'MessageBox.Show(" Erro : " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function
    Public Function validaGuid(strGuid As String) As Boolean
        Try
            Dim guid As New Guid
            Dim isValid As Boolean = False
            isValid = Guid.TryParse(strGuid.Trim.ToString, guid)
            Return isValid
        Catch ex As Exception
            Return False
        End Try
    End Function
    Public Function LocalizaGuidEmString(str As String) As String
        Try
            Dim guid As New Guid
            Dim isValid As Boolean = False
            Dim strArray As Array
            Dim separadorString() As String = {"|", ";", "{", "}", ",", ":", "[", "]", vbTab, vbCrLf}
            Dim StrRetorno As String = String.Empty
            If Not String.IsNullOrEmpty(str) Then
                strArray = str.Split(separadorString, System.StringSplitOptions.RemoveEmptyEntries)
                For i = 0 To UBound(strArray)
                    If validaGuid(Trim(strArray(i))) Then
                        StrRetorno = Trim(strArray(i).ToString)
                        Return StrRetorno
                        Exit Function
                    End If
                Next
            End If
            Return String.Empty
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function

#End Region
End Class



