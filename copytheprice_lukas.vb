Sub Find_Price()
    'Variables declaration
    Dim WordApp As Word.Application
    Dim ExcelApp As Excel.Application
    
    Dim ws As Worksheet
    Dim TextToFind As String
    Dim ApartmentPrice As String
    Dim Rng As Word.Range
    Dim StartPos As Long
    Dim EndPos As Long
    Application.ScreenUpdating = False
    
    TextToFind = "cenę brutto w kwocie "             'this text length is 16 caracters
        
    Set WordApp = GetObject(, "Word.Application")
    Set Rng = WordApp.ActiveDocument.Content
    
    
    StartPos = InStr(1, Rng, TextToFind)        'here we get 2269, we're looking 4 "TextToFind"
    EndPos = InStr(StartPos, Rng, ",00zł")      'here we get 2292, we're looking 4 ",00zł"
        
    If StartPos = 0 Or EndPos = 0 Then
        MsgBox ("Apartment price was not found!")
    Else
        StartPos = StartPos + Len(TextToFind)   'now start position is reassigned at 2285
                                                'this is where the first digit of the price is  :-)
        ApartmentPrice = Replace(Mid(Rng, StartPos, EndPos - StartPos), ".", "")
        
        'MsgBox "Price is " & ApartmentPrice & " pln."
        
        Set ExcelApp = GetObject(, "Excel.Application")
        Set ws = ExcelApp.ActiveSheet
        ws.Range("E27").Value = ApartmentPrice
        
        ExcelApp.Application.Visible = True
    End If
End Sub
