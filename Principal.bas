Attribute VB_Name = "Principal"
Option Explicit

Public UltimaLinha As Integer
Public UltimaColuna As Integer
Sub PrepararDados()
    
    Application.ScreenUpdating = True
    
    ContarRegistros 'Sub-Rotina para identificar colunas que não possuem registros
    removerLinhas   'Sub-Rotina para remover linhas que não estão na analise
    removerColunas 'Sub-Rotina parar remover  colunas que não possuem registros
    
    AjustarDados  'Sub-Rotina para trazer dados das variaiveis originais para as novas variávies
    
End Sub

Private Sub CriarBaseAnaiseteste()
    Sheets("tccOPO").Select
    ContarRegistros
    Selection.CurrentRegion.Select
    Selection.Copy
    Sheets("tccTeste").Select
    Range("A1").Select
    ActiveSheet.Paste

End Sub


Private Sub removerColunas()

UltimaColuna = 1

Columns(2).Select
Selection.Delete Shift:=xlToLeft

While Cells(1, UltimaColuna) <> ""
DoEvents

    If Cells(UltimaLinha + 1, UltimaColuna) = 0 Or _
            Cells(1, UltimaColuna) = "Amount" Or _
            Cells(1, UltimaColuna) = "ForecastCategoryName" Or _
            Cells(1, UltimaColuna) = "LastViewedDate" Or _
            Cells(1, UltimaColuna) = "LastReferencedDate" Then
        'excluir a coluna se ela não possuir valore
        'Debug.Print col, Cells(18263, col)
        Columns(UltimaColuna).Select
        Selection.Delete Shift:=xlToLeft
    Else
        UltimaColuna = UltimaColuna + 1
    End If
   

Wend

ContarRegistros

End Sub

Sub ContarRegistros()
    'rotina para identificar colunas que não possuem registros
Dim i As Integer

    'Dim strFormula As String
    UltimaColuna = 1
    
    While Cells(1, UltimaColuna) <> ""
    DoEvents
        UltimaColuna = UltimaColuna + 1
    Wend
    
    Range("A1").Select
    Selection.End(xlDown).Select
    UltimaLinha = ActiveCell.Row
    If IsNumeric(Cells(UltimaLinha, 1)) Then
        Rows(UltimaLinha).Select
        Selection.Delete Shift:=xlToLeft
        Selection.Delete Shift:=xlToLeft
        Selection.Delete Shift:=xlToLeft
        UltimaLinha = UltimaLinha - 1
    End If
    
'    If IsNumeric(Cells(UltimaLinha + 1, 1)) Then
'        Rows(UltimaLinha + 1).Select
'        Selection.Delete Shift:=xlToLeft
'        Selection.Delete Shift:=xlToLeft
'        Selection.Delete Shift:=xlToLeft
'    End If
    
    
    Range("A" & UltimaLinha + 1).Select
    ActiveCell.FormulaR1C1 = "=COUNTA(R[-" & UltimaLinha - 1 & "]C:R[-1]C)"
    Range("A" & UltimaLinha + 2).Select
    ActiveCell.FormulaR1C1 = "=SMALL(R[-" & UltimaLinha & "]C:R[-2]C,1)"
    Range("A" & UltimaLinha + 3).Select
    ActiveCell.FormulaR1C1 = "=LARGE(R[-" & UltimaLinha + 1 & "]C:R[-3]C,1)"
    
    If Cells(1, 136) = "XXX" Then
        For i = 1 To 3
            Range("A" & UltimaLinha + i & ":A" & UltimaLinha + i).Select
            Selection.Copy
            Range("B" & UltimaLinha + i & ":EF" & UltimaLinha + i & "").Select
            ActiveSheet.Paste
            Range("G" & UltimaLinha + i & "").Select
        Next
    Else
        For i = 1 To 3
            Range("A" & UltimaLinha + i & ":A" & UltimaLinha + i).Select
            Selection.Copy
            Range("B" & UltimaLinha + i & ":EA" & UltimaLinha + i & "").Select
            ActiveSheet.Paste
            Range("G" & UltimaLinha + i & "").Select
        Next
    End If
    
    UltimaColuna = 1



End Sub

Private Sub LimparDados()

    Selection.CurrentRegion.Select
    Selection.Copy
    Selection.PasteSpecial xlPasteValuesAndNumberFormats
    
    Selection.CurrentRegion.Select
    Cells.Replace What:=";", Replacement:="..", LookAt:=xlPart, SearchOrder _
        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False _
        , FormulaVersion:=xlReplaceFormula2
        
End Sub


Sub TesteRemoverLinhas()

Dim lin As Integer
Dim col As Integer
Dim colAno As Integer
Dim colStage As Integer
Dim colSetor As Integer
Dim colClosed As Integer

lin = 2
col = 1

Do
DoEvents
    If Cells(1, col) = "FiscalYear" Then
        colAno = col
    ElseIf Cells(1, col) = "StageName" Then
        colStage = col
    ElseIf Cells(1, col) = "Setor" Then
        colSetor = col
    ElseIf Cells(1, col) = "IsClosed" Then
        colClosed = col
    End If
        
    col = col + 1
Loop Until Cells(1, col) = ""

'Debug.Print UltimaLinha
''''UltimaLinha = 18309

Application.ScreenUpdating = False
While lin <= UltimaLinha
DoEvents
     
    If lin = 1445 Then
        Debug.Print
    End If
     
    'Remover linhas que não fazem parte da analise
    If (Cells(lin, colAno) < 2020 Or Cells(lin, colAno) > 2021) _
        Or (Cells(lin, colStage) = "Migrada") _
        Or (Cells(lin, colClosed) = "Falso") Then
        'Debug.Print col, Cells(18263, col)
        Rows(lin).Select
        Selection.Delete Shift:=xlToLeft
        UltimaLinha = UltimaLinha - 1
    Else
        lin = lin + 1
    End If
    If lin / 100 = Int(lin / 100) Or _
        UltimaLinha / 100 = Int(UltimaLinha / 100) Then
        'Application.ScreenUpdating = True
        Cells(lin, 1).Select
        Rows(lin).Select
        'Application.ScreenUpdating = False
    End If
    

Wend
Cells(1, 1).Select

'UltimaLinha = Lin - 1
Debug.Print UltimaLinha

End Sub

Sub removerLinhas()

Dim lin As Integer
Dim col As Integer
Dim colAno As Integer
Dim colStage As Integer
Dim colSetor As Integer
Dim colClosed As Integer

lin = 2
col = 1

Do
DoEvents
    If Cells(1, col) = "FiscalYear" Then
        colAno = col
    ElseIf Cells(1, col) = "StageName" Then
        colStage = col
    ElseIf Cells(1, col) = "Setor" Then
        colSetor = col
    ElseIf Cells(1, col) = "IsClosed" Then
        colClosed = col
    End If
        
    col = col + 1
Loop Until Cells(1, col) = ""

'Debug.Print UltimaLinha
''''UltimaLinha = 18309

Application.ScreenUpdating = False
While lin <= UltimaLinha
DoEvents
     
    If lin = 1445 Then
        Debug.Print
    End If
     
    'Remover linhas que não fazem parte da analise
    If (Cells(lin, colAno) < 2020 Or Cells(lin, colAno) > 2021) _
        Or (Cells(lin, colStage) = "Migrada") _
        Or (Cells(lin, colClosed) = "Falso") Then
        'Debug.Print col, Cells(18263, col)
        Rows(lin).Select
        Selection.Delete Shift:=xlToLeft
        UltimaLinha = UltimaLinha - 1
    Else
        lin = lin + 1
    End If
    If lin / 100 = Int(lin / 100) Or _
        UltimaLinha / 100 = Int(UltimaLinha / 100) Then
        'Application.ScreenUpdating = True
        Cells(lin, 1).Select
        Rows(lin).Select
        'Application.ScreenUpdating = False
    End If
    

Wend
Cells(1, 1).Select

'UltimaLinha = Lin - 1
Debug.Print UltimaLinha

End Sub


Sub CriarBaseAnaise()
    Sheets("tccOPO").Select
    ContarRegistros
    Selection.CurrentRegion.Select
    Selection.Copy
    Sheets("tcc").Select
    Range("A1").Select
    ActiveSheet.Paste

End Sub



Sub testeCriarBase()
    Sheets("tccOPO").Select
    ContarRegistros
    Selection.CurrentRegion.Select
    Selection.Copy
    Sheets("tccTeste").Select
    Range("A1").Select
    ActiveSheet.Paste

End Sub






Sub AjustarDados()
Debug.Print UltimaLinha, UltimaColuna

''''UltimaLinha = 18171
Dim vlr As Long
Dim lin As Integer

Dim colAno As Integer
Dim colStage As Integer
Dim colSetor As Integer
Dim colPonto As Integer
Dim colBudget As Integer
Dim colConcorente As Integer


If UltimaLinha = 0 Or UltimaColuna = 0 Then
    ContarRegistros
End If


lin = 2
UltimaColuna = 1

Do
DoEvents
    If Cells(1, UltimaColuna) = "FiscalYear" Then
        colAno = UltimaColuna
    ElseIf Cells(1, UltimaColuna) = "StageName" Then
        colStage = UltimaColuna
    ElseIf Cells(1, UltimaColuna) = "Setor" Then
        colSetor = UltimaColuna
    ElseIf Cells(1, UltimaColuna) = "Pontuacao_Media_de_Fechamento__c" Then
        colPonto = UltimaColuna
    ElseIf Cells(1, UltimaColuna) = "Ha_budget__c" Then
        colBudget = UltimaColuna
    ElseIf Cells(1, UltimaColuna) = "Modelo_concorrente__c" Then
        colConcorente = UltimaColuna
    End If
        
    UltimaColuna = UltimaColuna + 1
Loop Until Cells(1, UltimaColuna) = ""
UltimaColuna = UltimaColuna - 1

If Cells(1, UltimaColuna) = "_PontoQ" Then
    Cells(1, UltimaColuna + 1) = "_Ponto"
    Cells(1, UltimaColuna + 1) = "_PontoQ"
    'UltimaColuna = UltimaColuna + 2
End If

While lin <= UltimaLinha
DoEvents
    
    
    
    If Cells(lin, colSetor) = "0" Then
        Cells(lin, colSetor) = "N/A"
    ElseIf Cells(lin, colSetor) = "Tecnologia" Then
        Cells(lin, colSetor) = "TI e Serviços"
    ElseIf Cells(lin, colSetor) = "Tecnologia da Informação e Serviços" Then
        Cells(lin, colSetor) = "TI e Serviços"
    ElseIf Cells(lin, colBudget) = "Sim e não informou" Then
        Cells(lin, colBudget) = "Sim"
    ElseIf Cells(lin, colBudget) = "" Then
        Cells(lin, colBudget) = "n/a"
    ElseIf Cells(lin, colConcorente) = 0 Then
        Cells(lin, colConcorente) = ""
    ElseIf Cells(lin, 2) = "FALSO" Then
        Cells(lin, 2) = ""
    ElseIf Cells(lin, 2) = "VERDADEIRO" Then
        Cells(lin, 2) = ""
       
    End If

    'Padronizar a pontuação e dividir em quartil
    vlr = (Cells(lin, colPonto) / (Cells(UltimaLinha + 3, colPonto) - Cells(UltimaLinha + 2, colPonto))) - Cells(UltimaLinha + 2, colPonto)
    Cells(lin, UltimaColuna - 1) = vlr
    Cells(lin, UltimaColuna) = IIf(vlr <= 0.25, "Q1", IIf(vlr <= 0.5, "Q2", IIf(vlr <= 0.75, "Q3", "Q4")))
    
    If Cells(lin, colSetor) = 0 Then
            Cells(lin, colSetor) = ""
    ElseIf Cells(lin, colStage) = "Cancelada" Then
        Cells(lin, 4) = "Perdida"
    End If
    lin = lin + 1
    If lin / 100 = Int(lin / 100) Then
        Cells(lin, 1).Select
    End If
        

Wend

'UltimaLinha = lin - 1
Debug.Print UltimaLinha
End Sub

