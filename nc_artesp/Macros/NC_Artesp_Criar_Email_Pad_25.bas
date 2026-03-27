Attribute VB_Name = "NC_Artesp_Criar_Email_Pad_25"
Sub NC_Artesp_Criar_Email_Padrao_Rotina_Artesp_2025()

'Criar Email Padr�o para respostas dos Apontamentos da Artesp de Rotina


    Dim myText As String
    Dim Assunto As String
    Dim nume As String
    
    pula = Chr$(10)
    
    Dim newBook As Workbook
    Dim sheet As Worksheet
    Dim i As Byte
    Dim pastas As Workbooks
    Dim pasta As Workbook
    Dim Wb As Workbook, sFile As String, Spath As String

'On Error Resume Next
    Dim m As MailItem 'object/mail item iterator
    Dim recip As Recipient 'object to represent recipient(s)
    Dim reply As MailItem 'object which will represent the reply email


Dim Cod_fiscaliza��o(1000), Data_fiscaliza��o(1000), horario(1000), Rodovia(1000), concession�ria(1000), Km_Inicial(1000), m_Inicial(1000) As String
Dim km_final(1000), m_Final(1000), Sentido(1000), Data_Retorno(1000), Status_Retorno(1000), Tipo_Atividade(1000), Grupo_Atividade(1000), Atividade(1000) As String
Dim N�mero_notifica��o(1000), Data_Envio(1000), Data_Reparo(1000), Repons�vel(1000), Foto(1000) As String



  DisplayAlerts = False

    Spath = "L:\ENGENHARIA\CONSERVA\07 - Controles Artesp\_Relat�rio EAF - NC\Exportar\"
    sFile = Dir(Spath & "*.xls")
    
  Do While sFile <> ""
i = 5
x = 5

Dim exapp As Excel.Application
         Dim ExWbk As Workbook
         Dim ExWbk2 As Workbook
         Set exapp = New Excel.Application
         Set ExWbk2 = exapp.Workbooks.Open(Spath & sFile, UpdateLinks:=0)
         exapp.Visible = True

'Artesp = ActiveWorkbook.Name

ultimalinha = ExWbk2.Sheets("Sheet0").Cells(65536, 3).End(xlUp).Row

For i = 5 To ultimalinha

Cod_fiscaliza��o(x) = ExWbk2.Sheets("Sheet0").Range("C" & i).Value
Data_fiscaliza��o(x) = ExWbk2.Sheets("Sheet0").Range("D" & i).Value
horario(x) = ExWbk2.Sheets("Sheet0").Range("E" & i).Value
Rodovia(x) = ExWbk2.Sheets("Sheet0").Range("F" & i).Value
    If Left(Rodovia(x), 6) = "SP 075" Then Rodovia(x) = "SP 075"
    If Left(Rodovia(x), 6) = "SP 127" Then Rodovia(x) = "SP 127"
    If Left(Rodovia(x), 6) = "SP 280" Then Rodovia(x) = "SP 280"
    If Left(Rodovia(x), 6) = "SP 300" Then Rodovia(x) = "SP 300"
    If Left(Rodovia(x), 6) = "SPI 10" Then Rodovia(x) = "SPI 102/300"


concession�ria(x) = ExWbk2.Sheets("Sheet0").Range("G" & i).Value
Km_Inicial(x) = ExWbk2.Sheets("Sheet0").Range("H" & i).Value
m_Inicial(x) = ExWbk2.Sheets("Sheet0").Range("I" & i).Value
km_final(x) = ExWbk2.Sheets("Sheet0").Range("J" & i).Value
m_Final(x) = ExWbk2.Sheets("Sheet0").Range("K" & i).Value
Sentido(x) = ExWbk2.Sheets("Sheet0").Range("L" & i).Value
Data_Retorno(x) = ExWbk2.Sheets("Sheet0").Range("M" & i).Value
Status_Retorno(x) = ExWbk2.Sheets("Sheet0").Range("N" & i).Value
Tipo_Atividade(x) = ExWbk2.Sheets("Sheet0").Range("O" & i).Value
Grupo_Atividade(x) = ExWbk2.Sheets("Sheet0").Range("P" & i).Value
Atividade(x) = ExWbk2.Sheets("Sheet0").Range("Q" & i).Value
N�mero_notifica��o(x) = ExWbk2.Sheets("Sheet0").Range("R" & i).Value
Data_Envio(x) = ExWbk2.Sheets("Sheet0").Range("S" & i).Value
Data_Reparo(x) = ExWbk2.Sheets("Sheet0").Range("T" & i).Value
Repons�vel(x) = ExWbk2.Sheets("Sheet0").Range("U" & i).Value
Foto(x) = ExWbk2.Sheets("Sheet0").Range("V" & i).Value
x = x + 1
Next


Dim aOutlook As Object
Dim aEmail As Object
Dim obj As Object
Dim olInsp As Object
Dim myDoc As Object
Dim oRng As Object

Const PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
Set oApp = CreateObject("Outlook.Application")
Set oEmail = oApp.CreateItem(olMailItem)
Dim ToCc As Range, strBody, strSig As String
Dim fColorBlue, fColorGreen, fColorRed, fDukeBlue1, fDukeBlue2, fAggieMaroon, fAggieGray As String
Dim Greeting, emailContent As String
Dim emailOpen, emailSig As String
Const olFormatHTML As Long = 2


For Each m In Application.ActiveExplorer.Selection
If m.Class = olMail Then
Set reply = m.ReplyAll

            Assunto1 = reply.Subject
            Assunto = Replace(Assunto1, " [Email Externo] ", "")
            Assunto = Assunto & " - " & Rodovia(5) & " (" & Atividade(5) & ") - " & "Const: " & Data_fiscaliza��o(5) & " - Prazo: " & Data_Reparo(5)
            
myText = ""
 
          myText = "Prezados," & "<BR><BR>" & _
            "Seguem registros fotogr�ficos das supera��es de n�o conformidade, dentro do prazo regulamentado." & "<BR><BR>"
            


mytext2 = ""
Dim BaseArqFoto As String
Dim fname_nc As String, fname_nc2 As String
BaseArqFoto = "L:\ENGENHARIA\CONSERVA\06 - Abertura Externa Evento Kria\Arquivos\Arquivo Foto - Conserva\"
For l = 5 To ultimalinha

fname = BaseArqFoto & "Imagens Provis�rias - PDF\pdf (" & Foto(l) & ").jpg"
fname_nc = BaseArqFoto & "Imagens Provis�rias\nc (" & Foto(l) & ").jpg"
fname_nc2 = BaseArqFoto & "Imagens Provis�rias\nc (" & Cod_fiscaliza��o(l) & ")_1.jpg"

Set colAttach = reply.Attachments
Set oAttach = colAttach.Add(fname)
Set olkPA = oAttach.PropertyAccessor
olkPA.SetProperty PR_ATTACH_CONTENT_ID, "pdf%20(" & Foto(l) & ").jpg"

' Vistoria de campo / contra-foto executado (Art_022 / M02): nc (N).jpg e nc (cod)_1.jpg
If Dir(fname_nc) <> "" Then
    Set oAttach = colAttach.Add(fname_nc)
    Set olkPA = oAttach.PropertyAccessor
    olkPA.SetProperty PR_ATTACH_CONTENT_ID, "nc%20(" & Foto(l) & ").jpg"
End If
If Dir(fname_nc2) <> "" Then
    Set oAttach = colAttach.Add(fname_nc2)
    Set olkPA = oAttach.PropertyAccessor
    olkPA.SetProperty PR_ATTACH_CONTENT_ID, "nc%20(" & Cod_fiscaliza��o(l) & ")_1.jpg"
End If

mytext2 = "<b><u>" & mytext2 & Rodovia(l) & " - km " & Km_Inicial(l) & "," & m_Inicial(l) & " " & Sentido(l) & " - Const: " & Data_fiscaliza��o(l) & " - Prazo: " & Data_Reparo(l) & " - " & Atividade(l) & " - Cod. Fisc.: " & Cod_fiscaliza��o(l) & "</u></b><BR><BR>" & _
"<img src=""cid:pdf%20(" & Foto(l) & ").jpg"" height=295 width=711><BR><BR>"
If (Dir(fname_nc) <> "" And Dir(fname_nc2) <> "") Then
    mytext2 = mytext2 & "<table role=""presentation"" width=""711"" cellspacing=""6"" cellpadding=""0"" border=""0""><tr>"
    mytext2 = mytext2 & "<td width=""50%"" align=""center""><img src=""cid:nc%20(" & Foto(l) & ").jpg"" width=355 height=147></td>"
    mytext2 = mytext2 & "<td width=""50%"" align=""center""><img src=""cid:nc%20(" & Cod_fiscaliza��o(l) & ")_1.jpg"" width=355 height=147></td>"
    mytext2 = mytext2 & "</tr></table><BR><BR>"
ElseIf Dir(fname_nc) <> "" Then
    mytext2 = mytext2 & "<img src=""cid:nc%20(" & Foto(l) & ").jpg"" width=355 height=147><BR><BR>"
ElseIf Dir(fname_nc2) <> "" Then
    mytext2 = mytext2 & "<img src=""cid:nc%20(" & Cod_fiscaliza��o(l) & ")_1.jpg"" width=355 height=147><BR><BR>"
End If
mytext2 = mytext2 & "<BR><BR><BR><BR>"


Next


                Select Case reply.BodyFormat
                    
                    Case olFormatPlain, olFormatRichText, olFormatUnspecified
                        reply.To = ""
                        reply.Body = myText & mytext2 & reply.Body
                        reply.Subject = Assunto
                        reply.CC = "otavio.santos@viaappia.com.br;robert.rossi@viaappia.com.br; henrique.souza@viaappia.com.br"
                    
                    
                    
                    Case olFormatHTML
                        reply.To = ""
                        reply.HTMLBody = "<p>" & myText & "</p>" & mytext2 & reply.HTMLBody
                        reply.Subject = Assunto
                        reply.CC = "otavio.santos@viaappia.com.br ; robert.rossi@viaappia.com.br; henrique.souza@viaappia.com.br"

                End Select
                
reply.Save 'saves a draft copy to your SENT folder


End If
Next
 

ExWbk2.Close SaveChanges:=False

sFile = Dir()
Loop
MsgBox "Arquivos Lan�ados"
Exit Sub

MsgBox "Arquivos Lan�ados"

   

End Sub










