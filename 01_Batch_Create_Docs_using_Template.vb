Sub GenerateDocs()
    ' Important!!! Must add reference "Microsoft Word Object Library"
    ' 
    ' Usage:
    ' Create a word document "template.docx" containing basic template with dynamic fields / tokens
    ' Place this function in an macro enabled excel file, for instance, "tool.xlsm"
    ' On the main worksheet, configure the document names and token name + values in tabular format
    ' Start from A1 cell, construct a table looks like following:
    ' 
    ' Document Name                           $Name               $Date       $Amount  $Salutation
    ' debtor creditor  confirmation 1.docx    Jennifer            18/05/2018  51,486   Phd
    ' debtor creditor  confirmation 2.docx    Josephine           19/05/2018  12,345   Mr
    ' debtor creditor  confirmation 3.docx    Reddy Prasakahan    20/05/2018  942,512  Mrs
    ' debtor creditor  confirmation 4.docx    Prabakaran          21/05/2018  123,219  Miss
    ' debtor creditor  confirmation 5.docx    Robert Ng           22/05/2018  99,021   Baby
    ' 
    ' Run the script and it will create 5 new documents with the defined name
    ' The token (ex: $Name) will be replaced with the defined field in the new created file

    Dim i As Integer
    Dim j As Integer
    Dim key As Variant
    Dim counter As Integer
    
    Dim newDocName As String
    Dim token As String
    Dim value As String
    
    Dim wApp As Word.Application
    Dim tokenMapper As Object
    
    Set wApp = CreateObject("Word.Application")
    
    ' Create token mapper
    ' For example: 2 -> '$Name', 3 -> '$Date'
    Set tokenMapper = CreateObject("Scripting.Dictionary")
    For j = 2 To ActiveSheet.UsedRange.Columns.count
        tokenMapper.Add j, Cells(1, j).Text
    Next
    
    ' Iterate and process files
    counter = 0
    For i = 2 To ActiveSheet.UsedRange.Rows.count
        newDocName = Cells(i, 1).Text
        
        ' Initialize from template
        Set wDoc = wApp.Documents.Open(Filename:=ActiveWorkbook.Path & "\template.docx", ReadOnly:=True)

        ' Create file from template
        wDoc.SaveAs Filename:=ActiveWorkbook.Path & "\" & newDocName
        
        ' Replace token in new file
        For Each key In tokenMapper.Keys
            j = key
            token = tokenMapper(j)
            value = Cells(i, j).Text
            wDoc.Content.Find.Execute FindText:=token, ReplaceWith:=value, Format:=True, Replace:=wdReplaceAll
        Next key
        
        ' Create file from template
        wDoc.SaveAs Filename:=ActiveWorkbook.Path & "\" & newDocName
        
        ' Update counter
        counter = counter + 1
    Next i
    
    wApp.Quit
    
    MsgBox Str(counter) & " docx file(s) are generated successfully!"
    
End Sub
