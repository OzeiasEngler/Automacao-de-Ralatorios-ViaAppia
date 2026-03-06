Attribute VB_Name = "Art_022_EAF_Gerar_Mod_Ft_Exc_NC"
Sub Artesp_022_EAF_Gerar_Mod_Foto_Excel_NC_Rev2()

    Dim newBook As Workbook
    Dim sheet As Worksheet
    Dim i As Integer
    Dim pastas As Workbooks
    Dim pasta As Workbook
    Dim Wb As Workbook, sfile As String, spath As String
 
    'Desativa os avisos e atualiação da tela
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    

    spath = "L:\ENGENHARIA\CONSERVA\07 - Controles Artesp\_Relatório EAF - NC\Exportar\"
    sfile = Dir(spath & "*.xls")
    
  Do While sfile <> ""
  
  Workbooks.Open (spath & sfile)
  DisplayAlerts = False
  
    Dim nc(1000), hora(1000), arquivo(1000), km_i(1000), km_f(1000), km_i_Virg(1000), km_f_Virg(1000), texto(1000), rodovia(1000), cco, numero, kminicial, rodoviat(1000), Sentido(1000), relatorio, codigo(1000), embasamento(1000), complemento(1000) As String
    Dim Prazo(1000), y, foto(1000) As Integer
    Dim ano, mes, dia, n As Integer
    Dim Data_Reparo(1000), data(1000) As Date
    
    
    Dim tmpSheet As Worksheet
    Dim tmpChart As Chart
    Dim tmpImg As Object
    Dim fGIF As String
    Dim margem As Integer
    
    'Desativa os avisos e atualiação da tela
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
Artesp1 = ActiveWorkbook.Name
num = Len(Artesp1) - 4
Artesp = Left(Artesp1, num)
relatorio = Left(Artesp1, 8)

Diretorio = "L:\ENGENHARIA\CONSERVA\06 - Abertura Externa Evento Kria\Arquivos\Arquivo Foto - Conserva"
                                       
y = 5
a = 0
b = 0
i = 1

ultimalinha = Cells(65536, 4).End(xlUp).Row


Do
Do While y < ultimalinha + 1
      
                                    
                                    codigo(i) = Range("C" & y).Value
                                    data(i) = Range("D" & y).Value
                                    hora(i) = Range("E" & y).Value
                                    
                                    rodovia(i) = Left(Range("F" & y).Value, 6)
                                    If rodovia(i) = "SPI 10" Then
                                    rodovia(i) = "SPI 102/300"
                                    Else
                                    End If
                                    
                                    km_i(i) = Range("H" & y).Value & "+" & Range("I" & y).Value
                                    km_i_Virg(i) = Range("H" & y).Value & "," & Range("I" & y).Value
                                    km_f(i) = Range("j" & y).Value & "+" & Range("k" & y).Value
                                    km_f_Virg(i) = Range("j" & y).Value & "," & Range("k" & y).Value
                                    Sentido(i) = Range("L" & y).Value
                                    nc(i) = Range("Q" & y).Value
                                    
                                    foto(i) = Range("v" & y).Value
                                    Prazo(i) = DateDiff("d", Range("D" & y).Value, Range("T" & y).Value)
                                    'prazo(i) = DateDiff("d", data(i), data_reparo(i))
                                    Data_Reparo(i) = data(i) + Prazo(i)
                                    
If rodovia(i) = "SP 075" Then
n = 1
rodovia(i) = "SP-075"
rodoviat(i) = "SP075"
ElseIf rodovia(i) = "SP 127" Then
n = 2
rodovia(i) = "SP-127"
rodoviat(i) = "SP127"
ElseIf rodovia(i) = "SP 280" Then
n = 3
rodovia(i) = "SP-280"
rodoviat(i) = "SP280"
ElseIf rodovia(i) = "SP 300" Then
n = 4
rodovia(i) = "SP-300"
rodoviat(i) = "SP300"
ElseIf rodovia(i) = "SPI 102/300" Then
n = 5
rodovia(i) = "SPI-102/300"
rodoviat(i) = "SPI102_300"
ElseIf rodovia(i) = "CP 127_147" Then
n = 6
rodovia(i) = "FORA"
rodoviat(i) = "FORA"
ElseIf rodovia(i) = "CP 127_308" Then
n = 7
rodovia(i) = "FORA"
rodoviat(i) = "FORA"

Else
End If

            
   
       y = y + 1
       i = i + 1

 Loop
Loop Until y = ultimalinha + 1

ActiveWorkbook.Close False


Workbooks.Open ("L:\ENGENHARIA\CONSERVA\06 - Abertura Externa Evento Kria\Modelo\Modelo Abertura Evento Kria Conserva Rotina.xlsx")
Application.Wait (Now + TimeValue("0:00:10"))
Relatorio_Fotografico = ActiveWorkbook.Name

Namefolder = "L:\ENGENHARIA\CONSERVA\06 - Abertura Externa Evento Kria\Arquivos\Arquivo Foto - Conserva\"
NameFile = Format(Now, "yyyymmdd-hhmm") & " - " & Artesp & ".xlsx"
ActiveWorkbook.SaveAs (Namefolder & NameFile)

cc = 6
tt = 0
Do
Do While tt < i - 2

Rows(cc & ":" & cc + 4).Select
Selection.Copy

Rows(cc + 5 & ":" & cc + 5).Select
ActiveSheet.Paste


tt = tt + 1
cc = cc + 5
Loop
Loop Until tt = i - 2

x = 1
j = 8
Do
Do While x < i

Range("B" & j - 2).Select
ActiveCell.FormulaR1C1 = x

Range("D" & j).Select
ActiveCell.FormulaR1C1 = rodovia(x)

Range("D" & j + 1).Select
ActiveCell.FormulaR1C1 = km_i(x)

Range("F" & j).Select
ActiveCell.FormulaR1C1 = Sentido(x)

Range("F" & j + 1).Select
ActiveCell.FormulaR1C1 = km_f(x)

Range("F" & j + 2).Select
ActiveCell.FormulaR1C1 = data(x)

Range("G" & j - 1).Select
ActiveCell.FormulaR1C1 = Format(Data_Reparo(x), "mm/dd/yyyy")

Range("D" & j + 2).Select
ActiveCell.FormulaR1C1 = Format(Data_Reparo(x), "mm/dd/yyyy")

Range("C" & j - 2).Select
ActiveCell.FormulaR1C1 = nc(x)

Range("C" & j + 2).Select
ActiveCell.FormulaR1C1 = "Vencimento"

Range("H" & j + 1).Select
ActiveCell.FormulaR1C1 = codigo(x)

Range("G" & j).Select
ActiveCell.FormulaR1C1 = nc(x)

Range("H" & j + 2).Select
ActiveCell.FormulaR1C1 = relatorio

Range("L" & j + 2).Select
ActiveCell.FormulaR1C1 = Prazo(x)

Range("L" & j + 1).Select
ActiveCell.FormulaR1C1 = foto(x)

s = "L:\ENGENHARIA\CONSERVA\06 - Abertura Externa Evento Kria\Arquivos\Arquivo Foto - Conserva\Imagens Provisórias\" & "nc (" & foto(x) & ").jpg"

Range("C" & j - 1).Select
 With Range("C" & j - 1)
            Set imgIcon = ActiveSheet.Shapes.AddPicture( _
            fileName:=s, LinkToFile:=False, _
            SaveWithDocument:=True, Left:=.Left, Top:=.Top, Width:=275, Height:=210)
            

End With


x = x + 1
j = j + 5
Loop
Loop Until x = i

ActiveWorkbook.Save
ActiveWorkbook.Close


ano = Right(Data_Reparo(1), 4)
mes = Left(Right(Data_Reparo(1), 7), 2)
dia = Left(Data_Reparo(1), 2)

ano1 = Right(data(1), 4)
mes1 = Left(Right(data(1), 7), 2)
dia1 = Left(data(1), 2)


If nc(1) = "Recomposição de erosão em corte / aterro" Then
nc(1) = "Recomposição de erosão em corte_aterro"

ElseIf nc(1) = "Pavimentação/ Passeio/ Alambrado" Then
nc(1) = "Prédio e Pátio"


ElseIf nc(1) = "Hidráulica/ Esgoto/ Drenagem" Then
nc(1) = "Hidr_Esg_Dren"


ElseIf nc(1) = "Louças/ Metais" Then
nc(1) = "Predio - Louças_Metais"
End If


Origem = "L:\ENGENHARIA\CONSERVA\06 - Abertura Externa Evento Kria\Modelo\Modelo.xlsx"
Destino = "L:\ENGENHARIA\CONSERVA\07 - Controles Artesp\_Respostas - _Relatório EAF - NC\Pendentes\" & ano & mes & dia & " - " & Format(Now, "hhmmss") & " - " & rodoviat(1) & " - " & dia1 & "-" & mes1 & "-" & ano1 & " - " & nc(1) & ".xlsx"
FileCopy Origem, Destino

Workbooks.Open (Destino)
Application.Wait (Now + TimeValue("0:00:03"))
Relatorio_Fotografico = ActiveWorkbook.Name
h = 1
Range("B" & 1).Value = rodoviat(h) & " - km " & km_i_Virg(h) & " " & Sentido(h) & " - Const: " & dia1 & "/" & mes1 & "/" & ano1 & " - Prazo: " & dia & "/" & mes & "/" & ano & " - " & nc(h) & " - Cod. Fisc.: " & codigo(h)
Range("B" & 2).Value = dia1 & "-" & mes1 & "-" & ano1 & " - " & rodoviat(h) & " - " & km_i(h) & " - " & Sentido(h) & " - " & dia & "-" & mes & "-" & ano & " - " & nc(h) & " - " & codigo(h)

linha = 29




For t = 2 To i - 1

Rows(1 & ":" & 28).Select
h = h + 1
Selection.Copy
Rows(linha & ":" & linha).Select
ActiveSheet.Paste

Range("B" & linha).Value = rodoviat(h) & " - km " & km_i_Virg(h) & " " & Sentido(h) & " - Const: " & dia1 & "/" & mes1 & "/" & ano1 & " - Prazo: " & dia & "/" & mes & "/" & ano & " - " & nc(h) & " - Cod. Fisc.: " & codigo(h)
Range("B" & linha + 1).Value = dia1 & "-" & mes1 & "-" & ano1 & " - " & rodovia(h) & " - " & km_i(h) & " - " & Sentido(h) & " - " & dia & "-" & mes & "-" & ano & " - " & nc(h) & " - " & codigo(h)



v = "L:\ENGENHARIA\CONSERVA\06 - Abertura Externa Evento Kria\Arquivos\Arquivo Foto - Conserva\Imagens Provisórias - PDF\" & "pdf (" & foto(h) & ").jpg"

Range("B" & linha + 1).Select
 With Range("B" & linha + 1)
            Set imgIcon = ActiveSheet.Shapes.AddPicture( _
            fileName:=v, LinkToFile:=False, _
            SaveWithDocument:=True, Left:=.Left, Top:=.Top, Width:=480, Height:=202)
            imgIcon.Select
            Selection.ShapeRange.IncrementLeft 0.75
            Selection.ShapeRange.IncrementTop 0.75
 End With
 
linha = linha + 29

Next

v = "L:\ENGENHARIA\CONSERVA\06 - Abertura Externa Evento Kria\Arquivos\Arquivo Foto - Conserva\Imagens Provisórias - PDF\" & "pdf (" & foto(1) & ").jpg"

Range("B" & 2).Select
 With Range("B" & 2)
            Set imgIcon = ActiveSheet.Shapes.AddPicture( _
            fileName:=v, LinkToFile:=False, _
            SaveWithDocument:=True, Left:=.Left, Top:=.Top, Width:=480, Height:=202)
            imgIcon.Select
            Selection.ShapeRange.IncrementLeft 0.75
            Selection.ShapeRange.IncrementTop 0.75
            
End With

ActiveWorkbook.Save
ActiveWorkbook.Close

'ActiveWorkbook.Close savechanges:=True
  
sfile = Dir()
Loop

MsgBox "Processo Concluído - Arquivos de Fotos Gerados", vbInformation, "Gerar Arquivo de Foto"
Exit Sub



End Sub

