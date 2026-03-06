Attribute VB_Name = "Art_04_EAF_Rot_Juntar_Arq_Sepa"
Sub Artesp_04_EAF_Rot_Juntar_Arquivo_Exportar_Kria()

'Junta os arquivos separados de não conformidade para poder exportar para o kria

    Dim newBook As Workbook
    Dim sheet As Worksheet
    Dim i As Integer
    Dim pastas As Workbooks
    Dim pasta As Workbook
    Dim Wb As Workbook, sfile As String, spath As String
    
   Dim NumItem(5000), Origem(5000), Motivo(5000), Classificacao(5000), Tipo(5000), rodovia(5000), KMi(5000), KMf(5000), Sentido(5000), Local_(5000), Gestor(5000), Executores(5000), Data_Solicitação(5000), Data_Suspensão(5000), DtInicio_Prog(5000), DtFim_Prog(5000), DtInicio_Exec(5000), DtFim_Exec(5000), Prazo(5000), ObservaçãoGestor(5000), Observações(5000), Diretorio(5000), Arquivos(5000), Indicador(5000), Unidade(5000) As String
    

    
  'Desativa os avisos e atualiação da tela
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    spath = "L:\ENGENHARIA\CONSERVA\06 - Abertura Externa Evento Kria\Arquivos\Conservação\"
    sfile = Dir(spath & "*.xlsx")
t = 1
  Do While sfile <> ""
  
    Workbooks.Open (spath & sfile)
    Application.Wait (Now + TimeValue("0:00:04"))
    DisplayAlerts = False
  
  
    ultimalinhaprov = Cells(65536, 1).End(xlUp).Row
    
    
    For p = 2 To ultimalinhaprov
    Range("A" & p).Select
    NumItem(t) = Range("A" & p).Value
    
    ActiveCell.Offset(0, 1).Select
    Origem(t) = ActiveCell.Value
    
    ActiveCell.Offset(0, 1).Select
    Motivo(t) = ActiveCell.Value
    
    ActiveCell.Offset(0, 1).Select
    
    Classificacao(t) = ActiveCell.Value
    
    ActiveCell.Offset(0, 1).Select
    Tipo(t) = ActiveCell.Value
    
    ActiveCell.Offset(0, 1).Select
    rodovia(t) = ActiveCell.Value
    
    ActiveCell.Offset(0, 1).Select
    KMi(t) = ActiveCell.Value
    
    ActiveCell.Offset(0, 1).Select
    KMf(t) = ActiveCell.Value
    
    ActiveCell.Offset(0, 1).Select
    Sentido(t) = ActiveCell.Value
    
    ActiveCell.Offset(0, 1).Select
    Local_(t) = ActiveCell.Value
    
    ActiveCell.Offset(0, 1).Select
    Gestor(t) = ActiveCell.Value
    
    ActiveCell.Offset(0, 1).Select
    Executores(t) = ActiveCell.Value
    
    ActiveCell.Offset(0, 1).Select
    Data_Solicitação(t) = ActiveCell.Value
    
    ActiveCell.Offset(0, 1).Select
    Data_Suspensão(t) = ActiveCell.Value
    
    ActiveCell.Offset(0, 1).Select
    DtInicio_Prog(t) = ActiveCell.Value
    
    ActiveCell.Offset(0, 1).Select
    DtFim_Prog(t) = ActiveCell.Value
    
    ActiveCell.Offset(0, 1).Select
    DtInicio_Exec(t) = ActiveCell.Value
    
    ActiveCell.Offset(0, 1).Select
    DtFim_Exec(t) = ActiveCell.Value
    
    ActiveCell.Offset(0, 1).Select
    Prazo(t) = ActiveCell.Value
    
    ActiveCell.Offset(0, 1).Select
    ObservaçãoGestor(t) = ActiveCell.Value
    
    ActiveCell.Offset(0, 1).Select
    Observações(t) = ActiveCell.Value
    
    ActiveCell.Offset(0, 1).Select
    Diretorio(t) = ActiveCell.Value
    
    ActiveCell.Offset(0, 1).Select
    Arquivos(t) = ActiveCell.Value
    
    ActiveCell.Offset(0, 1).Select
    Indicador(t) = ActiveCell.Value
    
    ActiveCell.Offset(0, 1).Select
    Unidade(t) = ActiveCell.Value
   
t = t + 1
Next
   
ActiveWorkbook.Close False
    Application.Wait (Now + TimeValue("0:00:02"))
  
sfile = Dir()
Loop
    
    
    
    
Workbooks.Open ("L:\ENGENHARIA\CONSERVA\06 - Abertura Externa Evento Kria\Arquivos\Conservação\Acumulado\Padrão\Eventos Acumulado Artesp para Exportar Kria.xlsx")
Application.Wait (Now + TimeValue("0:00:04"))

ultimalinha = Cells(65536, 1).End(xlUp).Row
Range("A" & ultimalinha + 1).Select

g = 1

For r = 2 To t

     Range("A" & r).Select
    Range("A" & r).Value = NumItem(g)
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Origem(g)
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Motivo(g)
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Classificacao(g)
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Tipo(g)
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = rodovia(g)
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = KMi(g)
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = KMf(g)
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Sentido(g)
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Local_(g)
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Gestor(g)
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Executores(g)
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Data_Solicitação(g)
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Data_Suspensão(g)
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = DtInicio_Prog(g)
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = DtFim_Prog(g)
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = DtInicio_Exec(g)
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = DtFim_Exec(g)
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Prazo(g)
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = ObservaçãoGestor(g)
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Observações(g)
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Diretorio(g)
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Arquivos(g)
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Indicador(g)
    
    ActiveCell.Offset(0, 1).Select
    ActiveCell.Value = Unidade(g)

g = g + 1
Next
' ActiveSheet.Paste
'ultimalinha_baixo = Cells(65536, 3).End(xlUp).Row

  'Desativa os avisos e atualiação da tela
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
'Range("B" & ultimalinha + 1 & ":" & "B" & ultimalinha_baixo).Select
'Range("B" & ultimalinha + 1 & ":" & "B" & ultimalinha_baixo).Value = Artesp


'ActiveWorkbook.Close True
  '  Application.Wait (Now + TimeValue("0:00:04"))


Namefolder = "L:\ENGENHARIA\CONSERVA\06 - Abertura Externa Evento Kria\Arquivos\Conservação\Acumulado\"

            dia = Left(Data_Solicitação(g - 1), 2)
            mes = Right(Left(Data_Solicitação(g - 1), 5), 2)
            ano = Right(Left(Data_Solicitação(g - 1), 10), 4)

NameFile = ano & mes & dia & " - " & Format(Now, "hhmmss") & " - Eventos Acumulado Artesp para Exportar Kria.xlsx"
ActiveWorkbook.SaveAs (Namefolder & NameFile)
ActiveWorkbook.Close


MsgBox "Processo Concluído - Arquivos unidos", vbInformation, "Juntar Arquivos"


End Sub




