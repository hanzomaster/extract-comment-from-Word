Sub MainDocument()
  ' Create an Excel application
  Dim xlApp As Object
  On Error Resume Next
  Set xlApp = GetObject(, "Excel.Application")
  If Err Then
    Set xlApp = CreateObject("Excel.Application")
  End If
  On Error GoTo 0
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
    Dim commentHeader As String

    Dim napasNames As Variant
    napasNames = Array("NAPAS", "Nguyen Thi Mai Thuong", "Ly Dinh Quang", "Nguyen Hung Cuong", "Thuong Nguyen", "Do Thi Ha", "Nguyen Cao Cuong", "Steve", "Pham Duc Nhon")

    Dim obeNames As Variant
    obeNames = Array("Paul", "Peter", "Jim")

    Dim prefix As String
    Dim count As Integer
    count = 1

    Dim ancestorLineNumber(364) As Boolean
    Dim ancestorCountColumn As Integer
    Dim responseCount As Integer

    ancestorCount = 8
    responseCount = 0

    ' === Main worksheet code ===
    With xlWB.Worksheets(1)
      ' Add column headers
      .Cells(count, 1).Value = "STT"
      .Cells(count, 2).Value = "Version"
      .Cells(count, 3).Value = "Heading"
      .Cells(count, 4).Value = "Page"
      .Cells(count, 5).Value = "Status"
      .Cells(count, 6).Value = "Note"
      .Cells(count, 7).Value = "Commenter"
      .Cells(count, 8).Value = "Comment"


      ' Loop through all comments in the active document
      For i = 1 To ActiveDocument.Comments.count
        count = count + 1

        commentDate = ActiveDocument.Comments(i).Date
        formattedDate = Format(commentDate, "dd/mm/yyyy")
        commentText = ActiveDocument.Comments(i).Range.Text
        commenterFullName = ActiveDocument.Comments(i).Author

        ' Check For NAPAS names
        Dim matchFound As Boolean
        matchFound = False ' Initialize flag

        For Each nameToCheck In napasNames
          If InStr(UCase(commenterFullName), UCase(nameToCheck)) > 0 Then
            prefix = "NAPAS"
            matchFound = True
            Exit For
          End If
        Next nameToCheck

        ' Check For OBE names
        If Not matchFound Then ' Only check OBE If no NAPAS match was found
          For Each nameToCheck In obeNames
            If InStr(UCase(commenterFullName), UCase(nameToCheck)) > 0 Then
              prefix = "OBE"
              matchFound = True
              Exit For
            End If
          Next nameToCheck
        End If

        ' Default To SAVIS If no match in any group
        If Not matchFound Then
          prefix = "SAVIS"
        End If

        commentHeader = prefix & " - " & commenterFullName & " (" + formattedDate + "):"

        ' Check For comments With no ancestor
        If (ActiveDocument.Comments(i).Ancestor Is Nothing) Then
          ' ancestorLineNumber(count) = True
          ancestorCount = 8

          ' Find nearest heading
          Set headingRange = ActiveDocument.Comments(i).Reference.GoTo(What:=wdGoToHeading, Which:=wdGoToPrevious)
          If Not headingRange Is Nothing Then
            headingName = headingRange.Paragraphs(1).Range.Text
          Else
            headingName = "No Heading Found"
          End If

          pageNumber = ActiveDocument.Comments(i).Scope.Information(wdActiveEndAdjustedPageNumber)
          ' Check If the comment is resolve, If it is Then Set the Status column To "Resolved" Else "Pending"
          If (ActiveDocument.Comments(i).Done) Then
            .Cells(count, 5).Value = "Resolved"
          Else
            .Cells(count, 5).Value = "Pending"
          End If

          .Cells(count, 8).Value = commentHeader & vbCrLf & commentText
        Else
          ancestorCount = ancestorCount + 1
          .Cells(count, ancestorCount).Value = commentHeader & vbCrLf & commentText
          ' check if the first row of the ancestorCount column is empty, if it is then set the column to "Response" + responseCount
          If .Cells(1, ancestorCount).Value = "" Then
            responseCount = responseCount + 1
            .Cells(1, ancestorCount).Value = "Response " & responseCount
          End If
        End If

        ' Populate the Excel sheet
        .Cells(count, 1).Value = count
        .Cells(count, 3).Value = headingName
        .Cells(count, 4).Value = pageNumber
        .Cells(count, 7).Value = commenterFullName

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
      .Columns("F:AA").WrapText = True
      .Columns("C").WrapText = True
      ' Set vertical align middle For all cells
      .Cells.VerticalAlignment = xlCenter
      .Columns("A:E").HorizontalAlignment = xlCenter
      .Columns("G").HorizontalAlignment = xlCenter
      .Columns("C").ColumnWidth = 50
      .Columns("F:AA").ColumnWidth = 50
      ' AutoFit the columns
      .Columns("A:B").AutoFit
      .Columns("D:AA").AutoFit
      ' AutoFit rows height from 1 To count
      .Rows("1:" & count).AutoFit
    End With

    ' xlWB.Worksheets(1).Range("A1:C5").Select
    ' Add yellow background For row which index match is True in the ancestorLineNumber array
    'For i = 1 To UBound(ancestorLineNumber)
    '  If ancestorLineNumber(i) Then
    '    xlWB.Worksheets(1).Range("A" & i & ":K" & i).Interior.Color = RGB(255, 235, 156)
    '  End If
    'Next i

    ' Clean up
    Set xlWB = Nothing
    Set xlApp = Nothing
End Sub
