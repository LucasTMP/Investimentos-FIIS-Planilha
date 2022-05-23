Attribute VB_Name = "Módulo1"
Declare PtrSafe Sub Sleep Lib "Kernel32" (ByVal DwMilliSeconds As Long)
Dim IE As InternetExplorer

Sub ColetarValorPapel()

Dim papel As String
Dim url As String

Set IE = New InternetExplorer
IE.Visible = False

For Each c In Worksheets("Investimentos").Range("custodia").Rows
 
    papel = c.Cells(1, 1).Value
 
    If papel = "" Then
        GoTo DoNothing
    End If
 
    url = "https://www.infomoney.com.br/cotacoes/b3/fii/fundos-imobiliarios-" & papel & "/"
    IE.Navigate url
 
    Do Until IE.ReadyState = READYSTATE_COMPLETE
        DoEvents: Sleep 100
    Loop
    
    c.Cells(1, 3).Value = CCur(IE.Document.getElementsByClassName("typography__display--2-noscale typography--numeric spacing--mr1")(0).innerText)
    c.Cells(1, 4).Value = CCur(IE.Document.getElementsByClassName("typography__body--2 typography--wmedium")(0).innerText)
 
DoNothing:
    
Next
    
    
Worksheets("Investimentos").Range("J3:J45").Select
Selection.Style = "Currency"
Worksheets("Investimentos").Range("A1").Select
    
End Sub
