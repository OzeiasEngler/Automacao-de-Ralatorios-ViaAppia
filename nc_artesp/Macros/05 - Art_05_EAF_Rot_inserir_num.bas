Attribute VB_Name = "Art_05_EAF_Rot_inserir_num"
Sub Artesp_05_EAF_Rot_Ins_Num_Kria_Plan_Padrao()

'Inserir Numero do Evento do Kria na Planilha Padrão para exportar ao Calendário
Dim y As Long
Dim entrada As Long


p = 0
ultimalinha = 0
s = 0
nome = ActiveSheet.Name
ultimalinha = Worksheets(nome).Cells(65536, 1).End(xlUp).Row
y = 2
b = 0
entrada = InputBox("Digite o evento inicial")

            For y = 2 To ultimalinha
    
            Range("Y" & y).Select
            
            ActiveCell.FormulaR1C1 = entrada & "24"
            
            entrada = entrada + 1
          
    
            Next

End Sub



