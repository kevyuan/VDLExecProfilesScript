Option Explicit

Sub ExportIntoHTMLforExecProfiles()

    ' Declare variables
    Dim x As Range
    Dim y As Range
    Dim i As Integer
    
    Set x = Sheets("Sheet 1").Range("A3")
    i = 0
            
    Do While (x.Value <> "")
         Call ExportOneExecProfileFromRow(x, i)
         
         i = i + 1
         Set x = x.Offset(1, 0)
         
         
    Loop
    
End Sub

Sub ExportOneExecProfileFromRow(startingCell As Range, i As Integer)

    Dim line1, line2, line3, line4, line5, linebreak, entireBlurb As String
        
    Dim strUsername As String
    Dim strAlignment As String
    Dim strImageName As String
    Dim strName As String
    Dim strPosition As String
    Dim strNights As String
    Dim strEmail As String
    Dim strRandomFact As String
    
    
'    strUsername = "kyuan"
'    If (i Mod 2) = 0 Then
'        strAlignment = "alignright"
'
'    ElseIf (i Mod 2) = 1 Then
'        strAlignment = "alignleft"
'
'    End If
'    strName = "Kevin ""Fat"" Yuan"
'    strPosition = "HR Coordinator"
'    strNights = "Tuesday (Fat et al. (2013)), Wednesday (Dodge Rules Everything Around Me)"
'    strEmail = strUsername + "[at] vdldodgeball.ca"
'    strRandomFact = "asdf"
    
    
    strUsername = startingCell.Offset(0, 0).Value
    If (i Mod 2) = 0 Then
        strAlignment = "alignright"
    
    ElseIf (i Mod 2) = 1 Then
        strAlignment = "alignleft"
    
    End If
    strImageName = startingCell.Offset(0, 1).Value
    If (strImageName <> "") Then
        strImageName = strUsername + ".jpg"
    Else
        strImageName = "default.png"
    End If
    
    
    strName = startingCell.Offset(0, 2).Value
    strPosition = startingCell.Offset(0, 3).Value
    strNights = startingCell.Offset(0, 4).Value
    strEmail = strUsername + " [at] vdldodgeball.ca"
    strRandomFact = startingCell.Offset(0, 6).Value
    If (strRandomFact = "") Then
        strRandomFact = GrabRandomQuote()
    End If
    
    
    line1 = "<span style=""color: #ff6600;""><img class=""{strAlignment} size-full"" src=""http://vdldodgeball.ca/wp-content/uploads/2009/10/{strImageName}"" />Name</span>: {strName}"
    line2 = "<span style=""color: #ff6600;"">Position</span>: {strPosition}"
    line3 = "<span style=""color: #ff6600;"">Nights of Play</span>: {strNights}"
    line4 = "<span style=""color: #ff6600;"">Contact</span>: {strEmail}"
    line5 = "<em>{strRandomFact}</em>"
    linebreak = "<hr />"

    entireBlurb = line1 + vbCrLf + line2 + vbCrLf + line3 + vbCrLf + line4 + vbCrLf + vbCrLf + line5 + vbCrLf + vbCrLf + linebreak + vbCrLf

    entireBlurb = Replace(entireBlurb, "{strUsername}", strUsername)
    entireBlurb = Replace(entireBlurb, "{strName}", strName)
    entireBlurb = Replace(entireBlurb, "{strImageName}", strImageName)
    entireBlurb = Replace(entireBlurb, "{strAlignment}", strAlignment)
    entireBlurb = Replace(entireBlurb, "{strPosition}", strPosition)
    entireBlurb = Replace(entireBlurb, "{strNights}", strNights)
    entireBlurb = Replace(entireBlurb, "{strEmail}", strEmail)
    entireBlurb = Replace(entireBlurb, "{strRandomFact}", strRandomFact)
    

    Debug.Print entireBlurb
    
    'Sheets("output").Range("A1").Value = Sheets("output").Range("A1").Value + entireBlurb
    'startingCell.Offset(0, 10).Value = entireBlurb
    Call AppendToFile(entireBlurb)

End Sub

Function GrabRandomQuote()

    Dim quoteSheet As Worksheet
    Dim randomQuote As String
    
    Dim index As Integer
    index = Int((23) * Rnd)

    GrabRandomQuote = Sheets("quotes").Range("A1").Offset(index, 0).Value
            
    'Debug.Print "quote: " + randomQuote + ", index: " + Str(index)
    
End Function

Sub AppendToFile(text As String)
    Dim fs
    Dim a

    ' Create new file system object
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    ' Open existing text file for appending
    Set a = fs.OpenTextFile(ThisWorkbook.Path + "\html_markup.txt", 8, 1)
    
    ' Write String into file
    a.WriteLine (text)
    
    ' Close the text file
    a.Close
End Sub
