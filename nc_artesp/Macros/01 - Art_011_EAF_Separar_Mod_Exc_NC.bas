Attribute VB_Name = "Art_011_EAF_Separar_Mod_Exc_NC"
Public Type SHFILEOPSTRUCT
 hwnd As Long
 wFunc As Long
 pFrom As String
 pTo As String
 fFlags As Integer
 fAborted As Boolean
 hNameMaps As Long
 sProgress As String
 End Type
 Public Const FO_MOVE = &H1
 Public Const FO_RENAME = &H4
 Public Const FOF_SILENT = &H4
 Public Const FOF_NOCONFIRMATION = &H10
 Public Const FOF_FILESONLY = &H80
 Public Const FOF_SIMPLEPROGRESS = &H100
 Public Const FOF_NOCONFIRMMKDIR = &H200
 Public Const SHARD_PATH = &H2&
 Public Declare Function SHFileOperation _
 Lib "shell32.dll" Alias "SHFileOperationA" _
 (lpFileOp As SHFILEOPSTRUCT) As Long

Sub Artesp_011_EAF_Separar_Mod_Excel_NC_Rev2()

'A partir do modelo do arquivo excel da EAF Rotina, das Nao conformidades, ele separa e gera os arquivos de Nc separados, e numerado"


    Dim newBook As Workbook
    Dim sheet As Worksheet
    Dim i As Integer
    Dim pastas As Workbooks
    Dim pasta As Workbook
    Dim Wb As Workbook, sfile As String, spath As String, sfile_2 As String, entrada As String
    Dim cco As String
    
 
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

ultimalinha = Cells(65536, 3).End(xlUp).Row

y = 5
Q = 1

For u = 5 To ultimalinha
Range("i" & u).Select
Selection.NumberFormat = "@"
j = Len(Range("i" & u).Value)
If j = 0 Then Range("i" & u).Value = "000"
If j = 1 Then Range("i" & u).Value = "00" & Range("i" & u).Value
If j = 2 Then Range("i" & u).Value = "0" & Range("i" & u).Value

Range("k" & u).Select
Selection.NumberFormat = "@"
h = Len(Range("k" & u).Value)
If h = 0 Then Range("k" & u).Value = "000"
If h = 1 Then Range("k" & u).Value = "00" & Range("k" & u).Value
If h = 2 Then Range("k" & u).Value = "0" & Range("k" & u).Value

Range("V" & u).Value = Q
Q = Q + 1
Next

ActiveWorkbook.Save


For i = 5 To ultimalinha

Data1 = Range("T" & y).Value
Rodov = Range("F" & y).Value
tipo1 = Range("Q" & y).Value

dia1 = Left(Range("T" & y).Value, 2)
mes1 = Right(Left(Range("T" & y).Value, 5), 2)
ano1 = Right(Range("T" & y).Value, 4)

num = ano1 & mes1 & dia1


rod = Left(Range("F" & y).Value, 6)
If rod = "SPI 10" Then
rod = "SPI 102-300"
Else
End If


serv = Range("Q" & y).Value



If serv = "Pichação ao longo da rodovia" Then
            serv = "PICHAÇÃO"

ElseIf serv = "Substituição de pano rol. Medianamente comprometido" Then
            serv = "PAVIMENTO"
            
ElseIf serv = "Substituição de pano rol. Medianamente comprometido" Then
            serv = "PAVIMENTO"
            
ElseIf serv = "Reparo definitivo com recorte" Then
            serv = "REPARO RECORTE"
            
ElseIf serv = "Remoção de lixo doméstico das instalações" Then
            serv = "LIXO INST"

ElseIf serv = "Reparo de elemento de drenagem - manutenção" Then
            serv = "REPARO DE DRENAGEM"

ElseIf serv = "Despraguejamento" Then
            serv = "DESPRAGUEJAMENTO"

ElseIf serv = "Aceiros" Then
            serv = "ACEIRO"
            
ElseIf serv = "Selagem de trincas" Then
            serv = "SELAGEM TRINCA"
            
ElseIf serv = "Limpeza e varredura de áreas pavimentadas" Then
            serv = "LIMPEZA DE PAVIMENTO"

ElseIf serv = "Remoção de lixo e entulho da faixa de domínio" Then
            serv = "REMOÇÃO LIXO_ENTULHO"

ElseIf serv = "Defensa metálica (manutenção ou substituição)" Then
            serv = "REPARO DE DEFENSA"

ElseIf serv = "Depressão ou recalque de pequena extensão" Then
            serv = "PAVIMENTO - DEPRESSÃO"
            
ElseIf serv = "Panela ou buraco na faixa rolamento" Then
            serv = "PANELA"

ElseIf serv = "Reparo e reposição de cerca" Then
               serv = "REPARO CERCA"

ElseIf serv = "Manutenção árvores e arbustos" Then
                serv = "MANUTENÇÃO ÁRVORES"

ElseIf serv = "Drenagem fora de  plataforma limpeza geral" Then
                serv = "LIMP DRENAGEM FORA PLAT"

ElseIf serv = "Remoção de árvores ou galhos que não tem risco" Then
                serv = "REMOÇÃO DE GALHOS"

ElseIf serv = "Drenagem plataforma limpeza geral" Then
                serv = "LIMP DRENAGEM PLAT"

ElseIf serv = "Recomposição de erosão em corte / aterro" Then
                serv = "EROSÃO"
                
ElseIf serv = "Substituição de junta de dilatação" Then
                serv = "JUNTA DILATAÇÃO"
                
ElseIf serv = "Juntas e trincas: Limpeza e Resselagem" Then
                serv = "JUNTA DILATAÇÃO - LIMPEZA"

ElseIf serv = "Depressão em encontro de obra de arte" Then
                serv = "DEPRESSÃO OAE"
                
ElseIf serv = "Recuperação do revestimento vegetal" Then
                serv = "PLANTIO DE GRAMA"
                
ElseIf serv = "Remoção de massa verde" Then
                serv = "MASSA VERDE"
                
ElseIf serv = "Drenagem profunda limpeza geral" Then
                serv = "LIMP DE DRENAGEM PROF"
                
ElseIf serv = "Pavimentação/ Passeio/ Alambrado" Then
                serv = "PRÉDIO E PÁTIO - OUTROS"
                
ElseIf serv = "Poda manual ou mecanizada" Then
            serv = "PODA DO REVESTIMENTO"
            
ElseIf serv = "Bueiros limpeza geral" Then
            serv = "BUEIROS - LIMPEZA"
            
ElseIf serv = "Bordos e lajes quebrados reparo definitivo com recorte" Then
            serv = "PAVIMENTO RIGIDO"
            
ElseIf serv = "Correção de degrau entre pista e acostam. não pavimentado" Then
            serv = "DEGRAU PISTA_ACOSTAMENTO"
           
ElseIf serv = "Correção de degrau entre a pista e acostamento" Then
            serv = "DEGRAU PISTA_ACOSTAMENTO"
            
ElseIf serv = "Desobstrução de elemento de drenagem" Then
            serv = "DESOBSTRUÇÃO DE DRENAGEM"
                       
ElseIf serv = "Conformação lateral" Then
            serv = "CONFORMAÇÃO LATERAL"
            
        ElseIf serv = "Pichações e vandalismo" Then
        serv = "PICHAÇÃO"
        
        ElseIf serv = "Hidráulica/ Esgoto/ Drenagem" Then
        serv = "HIDR_ESG_DREN"
        
        ElseIf serv = "Barreira rígida manutenção e ou reparo" Then
        serv = "BARREIRA RIGIDA"
        
        ElseIf serv = "Reconformação de vias secundárias" Then
        serv = "CONFORM. LATERAL"
        
        ElseIf serv = "Louças/ Metais" Then
        serv = "PREDIO - LOUÇAS_METAIS"
        
     
  
Else
serv = serv
End If





dia = Left(Range("D" & y).Value, 2)
mes = Right(Left(Range("D" & y).Value, 5), 2)
ano = Right(Range("D" & y).Value, 4)
    
spath = "L:\ENGENHARIA\CONSERVA\07 - Controles Artesp\_Relatório EAF - NC\"
sfile = ActiveWorkbook.Name
    
'If num < 2000 Then
'sfile_2 = "L:\ENGENHARIA\CONSERVA\07 - Controles Artesp\_Relatório EAF - NC\Exportar\" & num & " - CONSTATAÇÕES NC LOTE 13 (" & rod & " - " & serv & ") - DATA - " & dia & "-" & mes & "-" & ano & " - P" & ".xls"
'Else
'sfile_2 = "L:\ENGENHARIA\CONSERVA\07 - Controles Artesp\_Relatório EAF - NC\Exportar\" & num & " - CONSTATAÇÕES NC LOTE 13 (" & rod & " - " & serv & ") - DATA - " & dia & "-" & mes & "-" & ano & ".xls"
'End If
sfile_2 = "L:\ENGENHARIA\CONSERVA\07 - Controles Artesp\_Relatório EAF - NC\Exportar\" & ano & mes & dia & " - CONSTATAÇÕES NC LOTE 13 (" & rod & " - " & serv & ") - Prazo - " & dia1 & "-" & mes1 & "-" & ano1 & ".xls"


strPath = sfile_2
If Dir(strPath) = vbNullString Then
Else
GoTo f
End If




ShellCopyFiles spath & sfile, sfile_2


Workbooks.Open (sfile_2)
ultimalinha = Cells(65536, 3).End(xlUp).Row
  
x = 5

            For k = 5 To ultimalinha

If Data1 = Range("T" & x).Value And Left(Rodov, 6) = Left(Range("F" & x).Value, 6) And tipo1 = Range("Q" & x).Value Then
x = x + 1
Else
Rows(x & ":" & x).Select
Selection.Delete
End If

Next
ActiveWorkbook.Close savechanges:=True
f:
Workbooks(sfile).Activate
y = y + 1
            Next
            
'Call Acumulado_Apontamento_Artesp_Rotina_REvisão_2
            
MsgBox "Processo Concluído - Arquivos Separados", vbInformation, "Separar Relatórios"

End Sub

Function ShellCopyFiles(sFile1 As String, sDestination As String)
 Dim r As Long
 Dim i As Integer
 Dim sFiles As String
 Dim SHFileOp As SHFILEOPSTRUCT
 sFile1 = sFile1 & Chr$(0) & Chr$(0)
 With SHFileOp
 .wFunc = FO_COPY
 .pFrom = sFile1
 .pTo = sDestination
 .fFlags = FOF_NOCONFIRMATION
 End With
 r = SHFileOperation(SHFileOp)
 End Function




