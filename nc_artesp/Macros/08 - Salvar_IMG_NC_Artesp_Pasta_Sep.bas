Attribute VB_Name = "Salvar_IMG_NC_Artesp_Pasta_Sep"

Sub salvar_imagem_NC_Arteso_Pasta_Separada()

Dim km, sfile, Prazo As String


ultimalinha = Cells(65536, 1).End(xlUp).Row
y = 2
Do
Do While y < ultimalinha + 1

e = InStr(Range("W" & y).Value, ";")
Origem = "L:\ENGENHARIA\CONSERVA\06 - Abertura Externa Evento Kria\Imagens\Conservação\" & Left(Range("W" & y).Value, e - 1)

rodovia = Range("F" & y).Value
If rodovia = "SPI102/300" Then
rodovia = "SP102"
Else
End If


km = Range("G" & y).Value


For b = 1 To Len(km)
    pedacokm = Mid(km, b, 1)
    If pedacokm = "+" Then
    pedacokm = ","
    Else
    End If
    wssent = wssent + pedacokm

    Next
    km = wssent


Sentido = Range("I" & y).Value

Tipo = Range("E" & y).Value

If Tipo = "Galhos/Árvores - Remoção" Then
 Tipo = "Galhos_Árvores - Remoção"
 Else
 End If
 
 If Tipo = "Galhos/Árvores - Poda" Then
 Tipo = "Galhos_Árvores - Poda"
 Else
 End If
 
  If Tipo = "Carga - Limpeza/Pista" Then
 Tipo = "Carga - Limpeza_Pista"
 Else
 End If
 
   If Tipo = "Carga - Limpeza/Faixa Domínio" Then
 Tipo = "Carga - Limpeza_Faixa Domínio"
 Else
 End If
 
    If Tipo = "Louças/ Metais" Then
 Tipo = "Predio e Patio"
 Else
 End If

 


Data_Sol_d = Left(Range("M" & y).Value, 2)

Data_Sol_m = Right(Left(Range("M" & y).Value, 5), 2)

Data_Sol_a = Right(Left(Range("M" & y).Value, 10), 4)

Data_Sol = Data_Sol_a & Data_Sol_m & Data_Sol_d



Prazo_d = Left(Range("P" & y).Value, 2)

Prazo_m = Right(Left(Range("P" & y).Value, 5), 2)

Prazo_a = Right(Left(Range("P" & y).Value, 10), 4)

Prazo = Prazo_d & Prazo_m & Prazo_a




evento = Range("T" & y).Value

numero_Ev = InStr(1, evento, "NC:", vbTextCompare)



ndesc = Len(evento)
evento = Right(evento, ndesc - numero_Ev - 3)

evento = evento & " - " & Range("Y" & y).Value


'65

If Tipo = "Pav. - Depressão no pavimento" Or Tipo = "Pav. - Pano de Rolamento" Then
Destino1 = "D:\Apontamentos NC Artesp - Imagens Classificadas\_Exportar" & "\" & rodovia & " - " & Sentido & " - " & km & " - " & Data_Sol & " - " & Prazo & " - " & evento & ".jpg"
FileCopy Origem, Destino1
End If

Destino = "D:\Apontamentos NC Artesp - Imagens Classificadas\" & Tipo & "\" & rodovia & " - " & Sentido & " - " & km & " - " & Data_Sol & " - " & Prazo & " - " & evento & ".jpg"

FileCopy Origem, Destino

       y = y + 1
       i = i + 1
       wssent = ""
Loop
Loop Until y = ultimalinha + 1


MsgBox "Processo Concluído - Arquivos de Fotos Gerados", vbInformation, "Gerar Arquivo de Foto"
End Sub

