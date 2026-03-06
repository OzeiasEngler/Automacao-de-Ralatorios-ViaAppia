Attribute VB_Name = "NC_Artesp_Criar_Email_Pad_25"
Sub NC_Artesp_Criar_Email_Padrao_Rotina_Artesp_2025()

'Criar Email Padrão para respostas dos Apontamentos da Artesp de Rotina


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


Dim Cod_fiscalização(1000), Data_fiscalização(1000), horario(1000), Rodovia(1000), concessionária(1000), Km_Inicial(1000), m_Inicial(1000) As String
Dim km_final(1000), m_Final(1000), Sentido(1000), Data_Retorno(1000), Status_Retorno(1000), Tipo_Atividade(1000), Grupo_Atividade(1000), Atividade(1000) As String
Dim Número_notificação(1000), Data_Envio(1000), Data_Reparo(1000), Reponsável(1000), Foto(1000) As String



  DisplayAlerts = False

    Spath = "L:\ENGENHARIA\CONSERVA\07 - Controles Artesp\_Relatório EAF - NC\Exportar\"
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

Cod_fiscalização(x) = ExWbk2.Sheets("Sheet0").Range("C" & i).Value
Data_fiscalização(x) = ExWbk2.Sheets("Sheet0").Range("D" & i).Value
horario(x) = ExWbk2.Sheets("Sheet0").Range("E" & i).Value
Rodovia(x) = ExWbk2.Sheets("Sheet0").Range("F" & i).Value
    If Left(Rodovia(x), 6) = "SP 075" Then Rodovia(x) = "SP 075"
    If Left(Rodovia(x), 6) = "SP 127" Then Rodovia(x) = "SP 127"
    If Left(Rodovia(x), 6) = "SP 280" Then Rodovia(x) = "SP 280"
    If Left(Rodovia(x), 6) = "SP 300" Then Rodovia(x) = "SP 300"
    If Left(Rodovia(x), 6) = "SPI 10" Then Rodovia(x) = "SPI 102/300"


concessionária(x) = ExWbk2.Sheets("Sheet0").Range("G" & i).Value
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
Número_notificação(x) = ExWbk2.Sheets("Sheet0").Range("R" & i).Value
Data_Envio(x) = ExWbk2.Sheets("Sheet0").Range("S" & i).Value
Data_Reparo(x) = ExWbk2.Sheets("Sheet0").Range("T" & i).Value
Reponsável(x) = ExWbk2.Sheets("Sheet0").Range("U" & i).Value
Foto(x) = ExWbk2.Sheets("Sheet0").Range("V" & i).Value
x = x + 1
Next


'_________________
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
'____________________


For Each m In Application.ActiveExplorer.Selection
If m.Class = olMail Then
Set reply = m.ReplyAll

            Assunto1 = reply.Subject
            Assunto = Replace(Assunto1, " [Email Externo] ", "")
            Assunto = Assunto & " - " & Rodovia(5) & " (" & Atividade(5) & ") - " & "Const: " & Data_fiscalização(5) & " - Prazo: " & Data_Reparo(5)
            
myText = ""
 
          myText = "Prezados," & "<BR><BR>" & _
            "Seguem registros fotográficos das superações de não conformidade, dentro do prazo regulamentado." & "<BR><BR>"
            


mytext2 = ""
For l = 5 To ultimalinha

fname = "L:\ENGENHARIA\CONSERVA\06 - Abertura Externa Evento Kria\Arquivos\Arquivo Foto - Conserva\Imagens Provisórias - PDF\pdf (" & Foto(l) & ").jpg"

'____________________
Set colAttach = reply.Attachments
Set oAttach = colAttach.Add(fname)
Set olkPA = oAttach.PropertyAccessor
olkPA.SetProperty PR_ATTACH_CONTENT_ID, "pdf%20(" & Foto(l) & ").jpg"


mytext2 = "<b><u>" & mytext2 & Rodovia(l) & " - km " & Km_Inicial(l) & "," & m_Inicial(l) & " " & Sentido(l) & " - Const: " & Data_fiscalização(l) & " - Prazo: " & Data_Reparo(l) & " - " & Atividade(l) & " - Cod. Fisc.: " & Cod_fiscalização(l) & "</u></b><BR><BR>" & _
"<img src=""cid:pdf%20(" & Foto(l) & ").jpg""height=295 width=711>" & "<BR><BR><BR><BR>"


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
MsgBox "Arquivos Lançados"
Exit Sub

MsgBox "Arquivos Lançados"

   

End Sub










