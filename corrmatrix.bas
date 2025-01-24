Attribute VB_Name = "Module1"
Function CorrMatrix(Data As Range, Optional InColumns As Boolean = True)

Dim r As Integer, nrows As Integer
Dim c As Integer, ncols As Integer

    nrows = Data.Rows.Count
    ncols = Data.Columns.Count

Dim DataArray() As Variant, ColX() As Variant, ColY() As Variant
    DataArray = Data

Dim Matrix() As Variant

If InColumns = True Then
    ReDim Matrix(1 To ncols, 1 To ncols)

    With Application.WorksheetFunction

        For r = 1 To ncols
            ColX = .Index(DataArray, 0, r)
            For c = 1 To ncols
                ColY = .Index(DataArray, 0, c)
                Matrix(r, c) = .Pearson(ColX, ColY)
            Next c
        Next r
    End With

Else
    ReDim Matrix(1 To nrows, 1 To nrows)

    With Application.WorksheetFunction

        For r = 1 To nrows
            ColX = .Index(DataArray, r, 0)
            For c = 1 To nrows
                ColY = .Index(DataArray, c, 0)
                Matrix(r, c) = .Pearson(ColX, ColY)
            Next c
        Next r
    End With
End If
    
CorrMatrix = Matrix
End Function







