Sub SubDocument()
  ' Create an Excel application
  Dim xlApp As Object
  On Error Resume Next
  Set xlApp = GetObject(, "Excel.Application")
  If Err Then
    Set xlApp = CreateObject("Excel.Application")
  End If
  On Error Goto 0
    xlApp.Visible = True

    ' Create a New workbook
    Dim xlWB As Object
    Set xlWB = xlApp.Workbooks.Add

    ' Initialize variables
    Dim i As Integer
    Dim commentDate As String
    Dim commentText As String
    Dim headingName As String
    Dim pageNumber As String
    Dim commenterFullName As String

    Dim napasNames As Variant
    napasNames = Array("NAPAS", "Nguyen Thi Mai Thuong", "Ly Dinh Quang", "Nguyen Hung Cuong", "Thuong Nguyen", "Do Thi Ha")

    Dim obeNames As Variant
    obeNames = Array("Paul", "Peter")

    Dim count As Integer
    count = 1

    Dim ancestorLineNumber(364) As Boolean
    Dim ancestorCount As Integer
    ancestorCount = 0

    ' === Main worksheet code ===
    With xlWB.Worksheets(1)
      ' Add column headers
      .Cells(count, 1).Value = "STT"
      .Cells(count, 2).Value = "From"
      .Cells(count, 3).Value = "Date"
      .Cells(count, 4).Value = "Comment"
      .Cells(count, 5).Value = "Response"
      .Cells(count, 6).Value = "Status"
      .Cells(count, 7).Value = "Heading"
      .Cells(count, 8).Value = "Page"
      .Cells(count, 9).Value = "Commenter"
      .Cells(count, 10).Value = "Deadline"


      ' Loop through all comments in the active document
      For i = 1 To ActiveDocument.Comments.count
        count = count + 1

        commentDate = ActiveDocument.Comments(i).Date
        commentText = ActiveDocument.Comments(i).Range.Text
        commenterFullName = ActiveDocument.Comments(i).Author

        ' Check For NAPAS names
        Dim matchFound As Boolean
        matchFound = False ' Initialize flag

        For Each nameToCheck In napasNames
          If InStr(UCase(commenterFullName), UCase(nameToCheck)) > 0 Then
            .Cells(count, 2).Value = "NAPAS"
            matchFound = True
           Exit For
          End If
        Next nameToCheck
        ' Check For OBE names
        If Not matchFound Then ' Only check OBE If no NAPAS match was found
          For Each nameToCheck In obeNames
            If InStr(UCase(commenterFullName), UCase(nameToCheck)) > 0 Then
              .Cells(count, 2).Value = "OBE"
              matchFound = True
             Exit For
            End If
          Next nameToCheck
        End If
        ' Default To SAVIS If no match in any group
        If Not matchFound Then
          .Cells(count, 2).Value = "SAVIS"
        End If

        ' Check For comments With no ancestor
        If (ActiveDocument.Comments(i).Ancestor Is Nothing) Then
          ancestorLineNumber(count) = True
          ancestorCount = ancestorCount + 1

          ' Find nearest heading
          Set headingRange = ActiveDocument.Comments(i).Reference.Goto(What:=wdGoToHeading, Which:=wdGoToPrevious)
          If Not headingRange Is Nothing Then
            headingName = headingRange.Paragraphs(1).Range.Text
          Else
            headingName = "No Heading Found"
          End If
          pageNumber = ActiveDocument.Comments(i).Scope.Information(wdActiveEndAdjustedPageNumber)
          ' Check If the comment is resolve, If it is Then Set the Status column To "Resolved" Else "Pending"
          If (ActiveDocument.Comments(i).Done) Then
            .Cells(count, 6).Value = "Resolved"
          Else
            .Cells(count, 6).Value = "Pending"
          End If

          ' Fill in the "Comment" column
          .Cells(count, 1).Value = ancestorCount
          .Cells(count, 4).Value = commentText
        Else
          ' Fill in the "Response" column (For comments With ancestors)
          .Cells(count, 5).Value = ActiveDocument.Comments(i).Range.Text
        End If

        ' Populate the Excel sheet
        .Cells(count, 3).Value = Format(commentDate, "mm/dd/yyyy")

        .Cells(count, 7).Value = headingName
        .Cells(count, 8).Value = pageNumber
        .Cells(count, 9).Value = commenterFullName

        ' Reset all values To null
        commentDate = Empty
        commentText = Empty
        headingName = Empty
        pageNumber = Empty
        commenterFullName = Empty
      Next i
    End With

    ' Format the Excel sheet
    With xlWB.Worksheets(1)
      .Columns("D:E").WrapText = True
      ' Set vertical align middle For all cells
      .Cells.VerticalAlignment = xlCenter
      ' Set horizontal align center For A, B, C, F, G, H, I columns
      .Columns("A:C").HorizontalAlignment = xlCenter
      .Columns("F:I").HorizontalAlignment = xlCenter
      ' Make D And E columns wider
      .Columns("D:E").ColumnWidth = 50
      ' AutoFit the columns
      .Columns("A:J").AutoFit
      ' AutoFit rows height from 1 To count
      .Rows("1:" & count).AutoFit
    End With

    ' xlWB.Worksheets(1).Range("A1:C5").Select
    ' Add yellow background For row which index match is True in the ancestorLineNumber array
    For i = 1 To UBound(ancestorLineNumber)
      If ancestorLineNumber(i) Then
        xlWB.Worksheets(1).Range("A" & i & ":J" & i).Interior.Color = RGB(255, 235, 156)
      End If
    Next i
    ' Clean up
    Set xlWB = Nothing
    Set xlApp = Nothing
End Sub
