Attribute VB_Name = "Módulo1"
Declare PtrSafe Sub Sleep Lib "Kernel32" (ByVal DwMilliSeconds As Long)
Dim ie As InternetExplorer

Sub ColetarValorPapel()

Dim papel As String
Dim url As String
Dim ie As Object
Dim html As Object
Dim table As Object
Dim row As Object
Dim cell As Object
Dim rowIndex As Integer, columnIndex As Integer

Set ie = New InternetExplorer
ie.Visible = False

Worksheets("Investimentos").Range("J3:K45").Select
Selection.NumberFormat = "General"
Worksheets("Investimentos").Range("A1").Select

For Each c In Worksheets("Investimentos").Range("custodia").Rows
 
    papel = c.Cells(1, 1).Value
 
    If papel = "" Then
        GoTo DoNothing
    End If
    
    url = "https://www.infomoney.com.br/cotacoes/b3/fii/fundos-imobiliarios-" & papel & "/"
    
    If papel = "RZAG11" Then
        url = "https://www.infomoney.com.br/cotacoes/b3/fii/fiagro-" & papel & "/"
    End If
    If papel = "SNAG11" Then
        url = "https://www.infomoney.com.br/cotacoes/b3/fii/suno-agro-snag11/"
    End If
    If papel = "EGAF11" Then
        url = "https://www.infomoney.com.br/cotacoes/b3/fii/ecoagro-egaf11/"
    End If
    If papel = "LIFE11" Then
        url = "https://fiis.com.br/life11/"
    End If
    If papel = "RZAT11" Then
        url = "https://www.infomoney.com.br/cotacoes/b3/fii/rzat11/"
    End If
    
    ie.Navigate url
    
    Do Until ie.ReadyState = READYSTATE_COMPLETE
        DoEvents: Sleep 100
    Loop
    
    If papel <> "LIFE11" Then
        c.Cells(1, 3).Value = CCur(ie.Document.getElementsByClassName("typography__display--2-noscale typography--numeric spacing--mr1")(0).innerText)
        c.Cells(1, 4).Value = CCur(ie.Document.getElementsByClassName("typography__body--2 typography--wmedium")(0).innerText)
    Else
        c.Cells(1, 3).Value = CCur(ie.Document.getElementsByClassName("wrapper indicators")(0).Children(2).Children(0).Children(1).innerText)
        c.Cells(1, 4).Value = CCur(ie.Document.getElementsByClassName("indicators")(0).Children(1).Children(0).Children(1).innerText)
    End If

DoNothing:
    
Next
    
ie.Quit
    
Worksheets("Investimentos").Range("J3:K45").Select
Selection.Style = "Currency"
Worksheets("Investimentos").Range("A1").Select
    
End Sub



