
Sub AddEquationNumber()
'
' v2 Macro
'
'
    Dim equationText As String
    equationText = Selection.Text
    Selection.TypeText Text:=vbTab
    
    ' Insert the equation number
     Selection.InsertCaption Label:="(", TitleAutoText:="InsertCaption3", Title _
        :="", Position:=wdCaptionPositionBelow, ExcludeLabel:=0
    Selection.TypeText Text:=")"
    
    ' Select the whole line
     Selection.GoTo What:=wdGoToBookmark, Name:="\para"
     
     ' Convert to Table
     Dim currentTalbe As Table
     Set currentTalbe = Selection.ConvertToTable(Separator:=wdSeparateByTabs, NumColumns:=2, _
        NumRows:=1, AutoFitBehavior:=wdAutoFitFixed)
        
      
    With currentTalbe
        .Style = "Table Grid"
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        
    End With
    
    ' Set all borders to be invisiible
    Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderRight).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderVertical).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
    Selection.Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
    
    ' Allign center middle the text in all cells
    Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    
    ' Set the size of the first column
    Dim columnWidth2 As Double
    Dim firstCellWidth As Double
    
    columnWidth2 = PointsToInches(Selection.Cells.Width) * 2
    firstCellWidth = columnWidth2 - 0.5
    
    Selection.Cells(1).SetWidth columnWidth:=InchesToPoints(firstCellWidth), RulerStyle:=wdAdjustFirstColumn
    
    ' set the spacing after and in the table to have no white space to eliminate blank lines
    Selection.ParagraphFormat.SpaceAfter = 0
    
    ' Select the second cell
    ' Move left to put curse at the beginning of the cell text
    ' Move right past the "("
    ' Delete the white space betweeen the "(" and the caption number
    
    Selection.Cells(2).Select
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    Selection.Delete Unit:=wdCharacter, Count:=1
    
    
     'Selection.Cells(2).Select
    'Selection.MoveRight Unit:=wdCharacter, Count:=1
    ' Selection.Delete Unit:=wdCharacter, Count:=1
  
 
End Sub


