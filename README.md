# excelVBA

==== code starts here

Sub picture()
'Pictures saved with file
'Set column width (ie, pic width) before running macro

'For Each r In Range("AX:AXX" & Cells(Rows.Count, 1).End(xlUp).Row) -- set x and xx to be the rows of the directory

''Left:=Cells(r.Row, 3).Left + Shrink, Top:=Cells(r.Row, 3) -- set the column to show the picture, 3 = column 3th = Column C
''If column 10 then it is L

Dim r As Range, Shrink As Long
Dim shpPic As Shape
Application.ScreenUpdating = False
Shrink = 0 'Provides negative offset from cell borders when > 0

On Error Resume Next
For Each r In Range("A2:A26" & Cells(Rows.Count, 1).End(xlUp).Row)
    If r.Value <> "" Then
        Set shpPic = ActiveSheet.Shapes.AddPicture(Filename:=r.Value, linktofile:=msoFalse, _
            savewithdocument:=msoTrue, Left:=Cells(r.Row, 3).Left + Shrink, Top:=Cells(r.Row, 3).Top + Shrink, _
                Width:=-1, Height:=-1)
        With shpPic
            .LockAspectRatio = msoTrue
            .Width = Columns(12).Width - (2 * Shrink)
            Rows(r.Row).RowHeight = .Height + (2 * Shrink)
        End With
    End If
Next r
Application.ScreenUpdating = True
End Sub

  ' format so the picture fits the cell frame
'        thisPic.Top = .Top + 1
'        thisPic.Left = .Left + 1
'        thisPic.Width = .Width - 2
'        thisPic.Height = .Height - 2
