Attribute VB_Name = "Módulo1"
Sub ColetarValorPapel()

Dim papel As String
Dim navegadorChrome As New Selenium.ChromeDriver
Dim url As String

url = "https://www.fundsexplorer.com.br/ranking"
navegadorChrome.Get (url)

Worksheets("Investimentos").Range("J3:K45").Select
Selection.NumberFormat = "General"
Worksheets("Investimentos").Range("A1").Select

For Each c In Worksheets("Investimentos").Range("custodia").Rows
    papel = c.Cells(1, 1).Value
    If papel <> "" Then
        For Each tr In navegadorChrome.FindElementById("upTo--default-fiis-table").FindElementsByTag("tr")
                If Split(tr.Attribute("innerText"))(0) Like "*" & papel & "*" Then
                        Debug.Print tr.FindElementsByTag("td")(1).Attribute("innerText")
                        c.Cells(1, 3).Value = CCur(tr.FindElementsByTag("td")(3).Attribute("innerText"))
                        c.Cells(1, 4).Value = CCur(tr.FindElementsByTag("td")(6).Attribute("innerText"))
                End If
        Next tr
    End If
Next
    
navegadorChrome.Quit
    
Worksheets("Investimentos").Range("J3:K45").Select
Selection.Style = "Currency"
Worksheets("Investimentos").Range("A1").Select
    
End Sub
