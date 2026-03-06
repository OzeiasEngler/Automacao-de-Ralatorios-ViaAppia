Attribute VB_Name = "Art_031_EAF_Rotina_Gerar_Kria"
Sub Artesp_031_EAF_Gerar_Mod_Foto_Excel_NC()

    Dim newBook As Workbook
    Dim sheet As Worksheet
    Dim i As Integer
    Dim pastas As Workbooks
    Dim pasta As Workbook
    Dim Wb As Workbook, sfile As String, spath As String
 
    'Desativa os avisos e atualiação da tela
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    

    spath = "L:\ENGENHARIA\CONSERVA\06 - Abertura Externa Evento Kria\Arquivos\Arquivo Foto - Conserva\"
    sfile = Dir(spath & "*.xlsx")
    
  Do While sfile <> ""
  
  Workbooks.Open (spath & sfile)
  DisplayAlerts = False
  

Call xx_Inserir_NaoConformidade_Rotina_Salvar_Imagem_Rev1_Kria

'ActiveWorkbook.Close savechanges:=True
  
sfile = Dir()
Loop

MsgBox "Processo Concluído - Arquivos de Fotos Gerados", vbInformation, "Gerar Arquivo de Foto"

Exit Sub



End Sub
