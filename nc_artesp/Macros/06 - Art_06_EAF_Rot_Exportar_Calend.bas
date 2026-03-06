Attribute VB_Name = "Art_06_EAF_Rot_Exportar_Calend"
Sub Artesp_06_Exportar_Calendar()

'exporta os eventos apontados pela EAF rotina para o calendário
    
    Dim objOutlook As Object
    Dim ObjAppt As Object
    Dim objNamespace As Object
    Dim objFolder As Object
    Dim OpenMAPIFolder As Object
    Dim objCalendar As Object
    Dim oWs As Worksheet, r As Long, i As Long, sStart As String


    Set objOutlook = CreateObject("Outlook.Application")
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objFolder = objNamespace.GetDefaultFolder(9).folders("Exportar")
 
Const wdPASTERTF As Long = 1
 
 Set oWs = ActiveSheet
 r = oWs.Range("A1").CurrentRegion.Rows.Count
 


    For i = 2 To r
    
    Set ObjAppt = objFolder.Items.Add 'create task item
    
Tipo_NC = Range("E" & i).Value

    With ObjAppt
        
               sStart = Right(oWs.Cells(i, 21), 10)
               Numero_Kria = Right(oWs.Cells(i, 21), 10)
               'sStart = Left(sStart, Len(sStart) - 4) & Year(Date)

               .Subject = Tipo_NC & " - " & oWs.Cells(i, 6) & " " & oWs.Cells(i, 7) & " " & oWs.Cells(i, 9) & " - Kria: " & oWs.Cells(i, 25)                                    'Assunto
               
               .Body = oWs.Cells(i, 21) & vbCrLf & vbCr & " - Data Constatação: " & Left(oWs.Cells(i, 13), 10) & vbCrLf & vbCr & oWs.Cells(i, 20)

                .Start = CDate(sStart)
                '.End = CDate(sStart)
                
                .AllDayEvent = True
                '.ReminderSet = True
                '.ReminderMinutesBeforeStart = 4320
                
                e = InStr(oWs.Cells(i, 23), ";")
                
                .Attachments.Add oWs.Cells(i, 22) & "\" & Left(oWs.Cells(i, 23), e - 1)
                
                .Attachments.Add oWs.Cells(i, 22) & "\" & Right(oWs.Cells(i, 23), Len(oWs.Cells(i, 23)) - e)
                 
                .Save
                
                
    End With
Next i
Set ObjAppt = Nothing
Set objFolder = Nothing
Set objNamespace = Nothing
Set objOutlook = Nothing

Call salvar_imagem_NC_Arteso_Pasta_Separada

MsgBox "Itens Inseridos no Calendário Outlook", vbInformation, "Mensagem de Conclusão"


End Sub

