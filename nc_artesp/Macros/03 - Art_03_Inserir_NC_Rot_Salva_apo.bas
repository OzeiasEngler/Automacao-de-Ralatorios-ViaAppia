Attribute VB_Name = "Art_03_Inserir_NC_Rot_Salva_apo"
Sub xx_Inserir_NaoConformidade_Rotina_Salvar_Imagem_Rev1_Kria()
        '
 
    Dim newBook As Workbook
    Dim sheet As Worksheet
    Dim i As Integer
    Dim pastas As Workbooks
    Dim pasta As Workbook
    
    Dim Wb As Workbook, sfile As String, spath As String, entrada As String
    Dim constata(1000), arquivo(1000), kmfinal_t(1000), serv(1000), classifica(1000), executor(1000), kminicial_t(1000), texto(1000), rodovia(1000), cco, numero, kminicial, rodoviat(1000), Sentido(1000), relatorio(1000), codigo(1000), embasamento(1000), complemento(1000) As String
    Dim Prazo(1000), y, foto(1000) As Integer
    Dim data(1000), ano, mes, dia, n As Integer
    
    Dim tmpSheet As Worksheet
    Dim tmpChart As Chart
    Dim tmpImg As Object
    Dim fGIF As String
    Dim margem As Integer
    
    'Desativa os avisos e atualiação da tela
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    'entrada = InputBox("Digite o nome do arquivo")
    cco = ActiveWorkbook.Name

Diretorio = "L:\ENGENHARIA\CONSERVA\06 - Abertura Externa Evento Kria\Imagens\Conservação"
                                       
y = 9
a = 0
b = 0
i = 1

ultimalinha = Cells(65536, 3).End(xlUp).Row


Do
Do While y < ultimalinha + 1
      
                                    relatorio(i) = Range("H" & y + 1).Value
                                    codigo(i) = Range("H" & y).Value
                                    complemento(i) = Range("L" & y).Value
                                    embasamento(i) = Range("G" & y - 2).Value

                                    rodovia(i) = Range("D" & y - 1).Value
                                    texto(i) = Range("G" & y - 1).Value
                                    foto(i) = Range("L" & y).Value
                                    
                                    
Origem = "L:\ENGENHARIA\CONSERVA\06 - Abertura Externa Evento Kria\Arquivos\Arquivo Foto - Conserva\Imagens Provisórias - PDF\pdf (" & foto(i) & ").jpg"
Destino = "L:\ENGENHARIA\CONSERVA\06 - Abertura Externa Evento Kria\Imagens\Conservação\pdf (" & foto(i) & ").jpg"
FileCopy Origem, Destino
                                    


If rodovia(i) = "SP-075" Then
n = 1
rodovia(i) = "SP075"
rodoviat(i) = "SP075"
ElseIf rodovia(i) = "SP-127" Then
n = 2
rodovia(i) = "SP127"
rodoviat(i) = "SP127"
ElseIf rodovia(i) = "SP-280" Then
n = 3
rodovia(i) = "SP280"
rodoviat(i) = "SP280"
ElseIf rodovia(i) = "SP-300" Then
n = 4
rodovia(i) = "SP300"
rodoviat(i) = "SP300"
ElseIf rodovia(i) = "SPI-102/300" Then
n = 5
rodovia(i) = "SPI102/300"
rodoviat(i) = "SPI102_300"
ElseIf rodovia(i) = "CP-127_147" Then
n = 6
rodovia(i) = "FORA"
rodoviat(i) = "CP-127_147"
ElseIf rodovia(i) = "CP-127_308" Then
n = 7
rodovia(i) = "FORA"
rodoviat(i) = "CP-127_308"

Else
End If

                                    Prazo(i) = Range("L" & y + 1).Value
                                    data(i) = Range("F" & y + 1).Value
dia = Left(data(i), 2)
mes = Right(Left(data(i), 5), 2)
ano = Right(data(i), 4)
                                    
                                    
                                   
                                    kminicial_t(i) = Range("D" & y).Value
                                    kmfinal_t(i) = Range("F" & y).Value

kminicial_0 = Format(Range("D" & y).Value, "0.000")
 For f = 1 To Len(kminicial_0)
 pedaco = Mid(kminicial_0, f, 1)
 If pedaco = "+" Then pedaco = ","
 kminicial_1 = kminicial_1 + pedaco
 Next
 kminicial = kminicial_1
            
                                    Sentido(i) = Range("F" & y - 1).Value
            Range("B" & y).Select
            Selection.NumberFormat = "000000"
            
            numero = Range("B" & y - 3).Value
            If numero < 10 Then
            numero = "00000" & numero
            ElseIf numero > 9 And numero < 100 Then
            numero = "0000" & numero
            Else
            numero = "000" & numero
            End If
            descricao = Range("G" & y - 1).Value
            

serv(i) = descricao


If serv(i) = "Pichação ao longo da rodovia" Then
            serv(i) = "Pichação"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"
            
            ElseIf serv(i) = "Pichações e vandalismo" Then
            serv(i) = "Pichação"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"

ElseIf serv(i) = "Substituição de pano rol. Medianamente comprometido" Then
            serv(i) = "Pav. - Pano de Rolamento"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"
            
    ElseIf serv(i) = "Reparo definitivo com recorte" Then
    serv(i) = "Pav. - Pano de Rolamento"
    classifica(i) = "Conservação Rotina"
    executor(i) = "Soluciona - Conserva"
    
    
    ElseIf serv(i) = "Substituição de pano rol. Medianamente comprometido" Then
            serv(i) = "Pav. - Pano de Rolamento"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"
            
         

ElseIf serv(i) = "Reparo de elemento de drenagem - manutenção" Then
            serv(i) = "Drenagem - Danificada"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"

ElseIf serv(i) = "Despraguejamento" Then
            serv(i) = "Controle fitossanitário"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"

ElseIf serv(i) = "Aceiros" Then
            serv(i) = "Vegetação - Outras Anomalias"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"
            
ElseIf serv(i) = "Selagem de trincas" Then
            serv(i) = "Pav. - Trinca"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"
            
ElseIf serv(i) = "Limpeza e varredura de áreas pavimentadas" Then
            serv(i) = "Pav. - Limpeza"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"

ElseIf serv(i) = "Remoção de lixo e entulho da faixa de domínio" Then
            serv(i) = "Lixo"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"
            
            ElseIf serv(i) = "Remoção de lixo doméstico das instalações" Then
            serv(i) = "Lixo"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"

ElseIf serv(i) = "Defensa metálica (manutenção ou substituição)" Then
            serv(i) = "Defensa Metálica - Danificada"
            classifica(i) = "Sinalização"
            executor(i) = "Soluciona - Sinalização"

ElseIf serv(i) = "Depressão ou recalque de pequena extensão" Then
            serv(i) = "Pav. - Depressão no pavimento"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"
            
ElseIf serv(i) = "Panela ou buraco na faixa rolamento" Then
            serv(i) = "Pav. - Buraco"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"

ElseIf serv(i) = "Reparo e reposição de cerca" Then
            serv(i) = "Cerca - Reparo"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"

ElseIf serv(i) = "Manutenção árvores e arbustos" Then
            serv(i) = "Galhos/Árvores - Poda"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"

ElseIf serv(i) = "Drenagem fora de  plataforma limpeza geral" Then
            serv(i) = "Drenagem - Limpeza"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"

ElseIf serv(i) = "Remoção de árvores ou galhos que não tem risco" Then
            serv(i) = "Galhos/Árvores - Remoção"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"

ElseIf serv(i) = "Drenagem plataforma limpeza geral" Then
            serv(i) = "Drenagem - Limpeza"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"

ElseIf serv(i) = "Recomposição de erosão em corte / aterro" Then
            serv(i) = "Erosão - Faixa Domínio"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"
                
ElseIf serv(i) = "Substituição de junta de dilatação" Then
            serv(i) = "OAE - Junta de Dilatação"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"
            
ElseIf serv(i) = "Juntas e trincas: Limpeza e Resselagem" Then
            serv(i) = "OAE - Junta de Dilatação"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"
            

ElseIf serv(i) = "Depressão em encontro de obra de arte" Then
            serv(i) = "Pav. - Depressão no encontro OAE"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"
            
ElseIf serv(i) = "Recuperação do revestimento vegetal" Then
            serv(i) = "Vegetação - Recomposição do Revestimento"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"
            
ElseIf serv(i) = "Remoção de massa verde" Then
            serv(i) = "Vegetação - Outras Anomalias"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"
            
            
ElseIf serv(i) = "Drenagem profunda limpeza geral" Then
            serv(i) = "Drenagem - Limpeza"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"
            
            
ElseIf serv(i) = "Pavimentação/ Passeio/ Alambrado" Then
            serv(i) = "Pav. - Outras anomalias"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"
            
            
            
ElseIf serv(i) = "Poda manual ou mecanizada" Then
            serv(i) = "Vegetação - Poda do revestimento"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"
            
ElseIf serv(i) = "Bueiros limpeza geral" Then
            serv(i) = "Drenagem - Limpeza"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"
            
ElseIf serv(i) = "Bordos e lajes quebrados reparo definitivo com recorte" Then
            serv(i) = "Pav. - Rígido danificado"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"
            
                        
ElseIf serv(i) = "Correção de degrau entre pista e acostam. não pavimentado" Then
            serv(i) = "Pav. - Bordo danificado"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"
                   
ElseIf serv(i) = "Correção de degrau entre a pista e acostamento" Then
            serv(i) = "Pav. - Bordo danificado"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"
            
ElseIf serv(i) = "Desobstrução de elemento de drenagem" Then
            serv(i) = "Drenagem - Limpeza"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"
            


ElseIf serv(i) = "Conformação lateral" Then
            serv(i) = "Conformação lateral"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"
            


ElseIf serv(i) = "Hidráulica/ Esgoto/ Drenagem" Then
            serv(i) = "OAE - Estrutura - Danos"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"
            
            ElseIf serv(i) = "Barreira rígida manutenção e ou reparo" Then
            serv(i) = "Barreira Rígida - Danificada"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"
                 
   
    ElseIf serv(i) = "Reconformação de vias secundárias" Then
            serv(i) = "Conformação lateral"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"
            
            
            ElseIf serv(i) = "Louças/ Metais" Then
            serv(i) = "Pav. - Outras anomalias"
            classifica(i) = "Conservação Rotina"
            executor(i) = "Soluciona - Conserva"

Else

serv(i) = ""

End If

            
        
        ActiveSheet.Range("c" & y - 3 & ":" & "f" & y + 1).Select
        

      Selection.CopyPicture Appearance:=xlScreen, Format:=xlBMP
      Application.Wait (Now + TimeValue("0:00:02"))
    'impede que se veja a acção acelerando o procedimento de cópia
    'e exportação
    Application.ScreenUpdating = False
    'uma folha para colocarmos o grafico sem atrapalhar o resto
    Set tmpSheet = Worksheets.Add
    'colocar um grafico nesta nova folha
    Charts.Add
      Application.Wait (Now + TimeValue("0:00:02"))
    'definições essenciais ao grafico, para que fique numa worksheet
    'e não numa folha grafico
    ActiveChart.Location Where:=xlLocationAsObject, Name:=tmpSheet.Name
    'Colar a  zona copiada para dentro da area do grafico
     Set tmpChart = ActiveChart
        With ActiveChart.Parent
         .Height = 540 ' resize
         .Width = 675 ' resize
             End With
         
         With tmpChart

         .Paste

           Application.Wait (Now + TimeValue("0:00:01"))
            End With
    
    With Selection
                .Height = 540 ' resize
                .Width = 675 ' resize
    End With
        

    
                        arquivo(i) = ano & mes & dia & " - " & Format(Now, "hhmmss") & " - " & n & "_" & "Roti-" & numero & "-" & rodoviat(i) & " " & kminicial & " " & Sentido(i) & ".jpg"

    fGIF = "L:\ENGENHARIA\CONSERVA\06 - Abertura Externa Evento Kria\Imagens\Conservação\" & arquivo(i)
    
    
  
    'ano & mes & dia & "_Artes" & n & "_" & "Roti-" & numero & "-" & rodoviat(i) & " " & kminicial & " " & sentido(i) & ".jpg"
       
    tmpChart.Export fileName:=fGIF, filtername:="gif"
    Application.Wait (Now + TimeValue("0:00:02"))
    'eliminar a folha temporaria sem avisos
    Application.DisplayAlerts = False
    tmpSheet.Delete
    Application.DisplayAlerts = True
    'repor o estado normal
    Application.ScreenUpdating = True

    GoTo fim
erro:
    MsgBox "Erro: " & Err.Description, _
            vbCritical, _
           "Erro: " & Err.Number
fim:
    Set tmpSheet = Nothing
    Set tmpChart = Nothing
    Set tmpImg = Nothing

   
h:
       y = y + 5
       i = i + 1
kminicial = ""
kminicial_0 = ""
kminicial_1 = ""
 Loop
Loop Until y = ultimalinha + 4

ActiveWorkbook.Close False


Workbooks.Open ("L:\ENGENHARIA\CONSERVA\06 - Abertura Externa Evento Kria\Modelo\_Planilha Modelo Kcor-Kria.XLSX")
Application.Wait (Now + TimeValue("0:00:10"))
Artesp = ActiveWorkbook.Name
j = 2
x = 1

Do
Do While x < i
Range("A" & j).Select
ActiveCell.FormulaR1C1 = x

ActiveCell.Offset(0, 1).Select
ActiveCell.FormulaR1C1 = "Artesp"

ActiveCell.Offset(0, 1).Select
ActiveCell.FormulaR1C1 = "2"

ActiveCell.Offset(0, 1).Select
ActiveCell.FormulaR1C1 = classifica(x)

ActiveCell.Offset(0, 1).Select
ActiveCell.FormulaR1C1 = serv(x)


ActiveCell.Offset(0, 1).Select
ActiveCell.FormulaR1C1 = rodovia(x)


ActiveCell.Offset(0, 1).Select
ActiveCell.FormulaR1C1 = kminicial_t(x)


ActiveCell.Offset(0, 1).Select
ActiveCell.FormulaR1C1 = kmfinal_t(x)


ActiveCell.Offset(0, 1).Select
ActiveCell.FormulaR1C1 = Sentido(x)

ActiveCell.Offset(0, 2).Select
ActiveCell.FormulaR1C1 = "Conservação"

ActiveCell.Offset(0, 1).Select
ActiveCell.FormulaR1C1 = executor(x)


ActiveCell.Offset(0, 1).Select
ActiveCell.FormulaR1C1 = Format(data(x), "mm/dd/yyyy")

ActiveCell.Offset(0, 2).Select
ActiveCell.FormulaR1C1 = Format(data(x), "mm/dd/yyyy")

ActiveCell.Offset(0, 1).Select
ActiveCell.FormulaR1C1 = Format(embasamento(x), "mm/dd/yyyy")


ActiveCell.Offset(0, 3).Select
ActiveCell.FormulaR1C1 = Prazo(x)

ActiveCell.Offset(0, 1).Select
ActiveCell.FormulaR1C1 = "--> Relatório EAF Conservação Rotina nº: " & relatorio(x) & vbCrLf & vbCr & "--> Código NC: " & codigo(x)

ActiveCell.Offset(0, 1).Select
If complemento(x) = "" Then
ActiveCell.FormulaR1C1 = texto(x) & vbCrLf & vbCr & vbCrLf & vbCr & "- Data Superação Artesp ----> " & embasamento(x)
Else
ActiveCell.FormulaR1C1 = texto(x) & vbCrLf & vbCr & vbCrLf & vbCr & "- Complemento ----> " & complemento(x) & vbCrLf & vbCr & vbCrLf & vbCr & "- Embasamento ----> " & embasamento(x)
End If

ActiveCell.Offset(0, 1).Select
ActiveCell.FormulaR1C1 = Diretorio
ActiveCell.Offset(0, 1).Select
ActiveCell.FormulaR1C1 = arquivo(x) & ";" & "pdf (" & foto(x) & ").jpg"

x = x + 1
j = j + 1
Loop
Loop Until x = i

Namefolder = "L:\ENGENHARIA\CONSERVA\06 - Abertura Externa Evento Kria\Arquivos\Conservação\"
NameFile = Format(Now, "yyyymmdd-hhmm") & " - " & cco
ActiveWorkbook.SaveAs (Namefolder & NameFile)
ActiveWorkbook.Close
    

Arq_Antigo = "L:\ENGENHARIA\CONSERVA\06 - Abertura Externa Evento Kria\Arquivos\Arquivo Foto - Conserva\" & cco
Arq_Novo = "L:\ENGENHARIA\CONSERVA\06 - Abertura Externa Evento Kria\Arquivos\Arquivo Foto - Conserva\" & "_Processado - " & cco
Name Arq_Antigo As Arq_Novo

End Sub



