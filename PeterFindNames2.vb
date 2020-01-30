Option Explicit
Sub PeterFindNames2()
    'Variables declaration
    Dim textToFind As String
    Dim startPos As Long
    Dim endPos As Long
    Dim searchArea As Word.Range
    
    'set object variables
    Set searchArea = Word.Application.ActiveDocument.Content
    'Set searchArea = WordApp.ActiveDocument.Content
    textToFind = "REGON 364061169, NIP 951-24-09-783,"
    startPos = InStr(1, searchArea, textToFind) - 1  'here we get 1421, we're looking 4 "TextToFind"
    If (startPos = 0) Then Exit Sub

    'adjust the searchArea to start from where we found the text, until the end of the document
    searchArea.SetRange Start:=startPos, End:=searchArea.End

    '---we want the name at the start of the very next paragraph
    '   (the current paragraph with the text to find is paragraph 1)
    Dim theParagraph As Word.Paragraph
    Dim scndParagraph As Word.Paragraph
    Dim thrdParagraph As Word.Paragraph
    Set theParagraph = searchArea.Paragraphs(2)
    Set scndParagraph = searchArea.Paragraphs(3)
    Set thrdParagraph = searchArea.Paragraphs(4)

    Dim itemNumber As Long
    Dim firstName As String
    Dim lastName As String
    Dim firstSurname As String
    'Debug.Print theParagraph.Range.Words(1)
    'Debug.Print scndParagraph.Range.Words(1)
    'Debug.Print thrdParagraph.Range.Words(1)
    
    'the VBA CLng function converts an expression into a Long data type
    itemNumber = CLng(Trim(theParagraph.Range.Words(1)))
    firstName = Trim$(theParagraph.Range.Words(3))
    lastName = Trim$(theParagraph.Range.Words(4))
    firstSurname = Trim$(theParagraph.Range.Words(5))

    'Debug.Print "Name = " & firstName & " " & lastName & " in Item #" & itemNumber
    Debug.Print firstName & " " & lastName & " " & firstSurname
End Sub
