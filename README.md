# Excel-VBA-fit-image-to-cell

Excel does not have native support for fitting images to cell. To have this option in excel open an excel file and save file as Macro-enabled workbook. Open VBA window from developer tools and add below code. Assign this `FitImageToCell` macro to a key combination (e.g `ctrl + d`). So you are ready to use. Insert an image, click the image first, press your key combination `ctrl + d`, then click a cell and again press your key combination `ctrl + d`, you ara done!

```vba
Sub FitImageToCell()

  ' Declare variables
  Dim imgName As Range
  Dim imgHeight As Range
  Dim imgWidth As Range
  
  Dim cellName As Range
  Dim cellHeight As Range
  Dim cellWidth As Range
  
  Dim calcTop As Double
  Dim calcLeft As Double
  
  Dim imgRatio As Double
  Dim cellRatio As Double
  Dim padding As Double
  
  Dim sheetName As String: sheetName = "__selection"
  Dim sheetExists As Boolean: sheetExists = False
  
  
  ' Check if sheet exists
  For Each sh In Worksheets
    If sh.Name = sheetName Then
      sheetExists = True
      Exit For
    End If
  Next sh
  
  
  If Not sheetExists Then
  
    ' Create sheet
    Sheets.Add.Name = sheetName
    
    ' Create table
    Sheets(sheetName).Range("A1") = "Image Fit To Cell"
    Sheets(sheetName).Range("A2") = "Name"
    Sheets(sheetName).Range("A3") = "Height"
    Sheets(sheetName).Range("A4") = "Width"
    Sheets(sheetName).Range("B1") = "Image"
    Sheets(sheetName).Range("C1") = "Cell"

    ' Set gaps
    Rows.RowHeight = 22.5
    Columns("A:C").ColumnWidth = 16.43
    
    ' Select active sheet
    Sheets(sheetName).Select
    
    ' Select range
    Range("A1:C4").Select
    
    ' Align cells to center and middle
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    ' Draw borders
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    
    ' Deselect range
    Range("A1").Select
    
    ' Hide sheet
    Worksheets(sheetName).Visible = False
    
  End If
  
  
  ' Set ranges
  Set imgName = Sheets(sheetName).Range("B2")
  Set imgHeight = Sheets(sheetName).Range("B3")
  Set imgWidth = Sheets(sheetName).Range("B4")
  
  Set cellName = Sheets(sheetName).Range("C2")
  Set cellHeight = Sheets(sheetName).Range("C3")
  Set cellWidth = Sheets(sheetName).Range("C4")
  
  
  ' Check if selection is an image
  If TypeName(Selection) = "Picture" Then
  
   ' Preserve active image information
    imgName = Selection.Name
    imgHeight = Selection.Height
    imgWidth = Selection.Width
    
    ' Truncate operation data
    cellName.ClearContents
    cellHeight.ClearContents
    cellWidth.ClearContents
  
  End If
  
  
  ' Check if selection is a cell and stored image name cell is not empty
  If TypeName(Selection) = "Range" And Not IsEmpty(imgName) Then
    
    ' Preserve active cell information
    cellName = ActiveCell.Address
    cellHeight = ActiveSheet.Cells(ActiveCell.Row + 1, 1).Top - ActiveSheet.Cells(ActiveCell.Row, 1).Top
    cellWidth = ActiveSheet.Cells(1, ActiveCell.Column + 1).Left - ActiveSheet.Cells(1, ActiveCell.Column).Left
    
    ' Select image
    ActiveSheet.Shapes.Range(Array(imgName)).Select
  
    ' Calculate ratio
    cellRatio = cellWidth / cellHeight
    imgRatio = imgWidth / imgHeight
    
    ' Define padding ratio
    padding = 0.8
    
    ' Scale image
    If cellRatio > imgRatio Then
      Selection.ShapeRange.Height = cellHeight * padding
    Else
      Selection.ShapeRange.Width = cellWidth * padding
    End If
    
    ' Set scaled image values
    imgHeight = Selection.Height
    imgWidth = Selection.Width
    
    ' Calculate center
    calcTop = ActiveSheet.Range(cellName).Top + (cellHeight - imgHeight) / 2
    calcLeft = ActiveSheet.Range(cellName).Left + (cellWidth / 2) - (imgWidth / 2)
    
    ' Move to cell center
    Selection.ShapeRange.Top = calcTop
    Selection.ShapeRange.Left = calcLeft
    Selection.Placement = xlMove
    
    ' Re-select preserved cell
    ActiveSheet.Range(cellName).Select
    
    ' Truncate operation data
    imgName = ClearContents
    imgHeight = ClearContents
    imgWidth = ClearContents
    
  End If


End Sub

```

## (Bonus) Pin All Images
Sub PinAllImages()

  For Each pic In ActiveSheet.Pictures
      pic.Placement = xlMove
  Next

End Sub

