Attribute VB_Name = "CorrMod"
' Function: CorrMatrix
' Description: This function calculates the correlation matrix for a given dataset.
' Parameters:
'    Data (Range): The input data range for which the correlation matrix is to be calculated.
'    InColumns (Boolean, Optional): If True, calculates correlation matrix for columns; if False, for rows. Default is True.
'
' Returns:
'    A 2D array containing the correlation coefficients.

Function CorrMatrix(Data As Range, Optional InColumns As Boolean = True)

    ' Variables to store the number of rows and columns in the input data range
    Dim r As Integer, nrows As Integer
    Dim c As Integer, ncols As Integer

    ' Get the number of rows and columns in the input data range
    nrows = Data.Rows.Count
    ncols = Data.Columns.Count

    ' Arrays to store the input data and columns/rows for correlation calculation
    Dim DataArray() As Variant, ColX() As Variant, ColY() As Variant
    DataArray = Data

    ' Array to store the correlation matrix
    Dim Matrix() As Variant

    ' Check if correlation matrix is to be calculated for columns
    If InColumns = True Then
        ' Resize the Matrix array to store the correlation coefficients for columns
        ReDim Matrix(1 To ncols, 1 To ncols)

        ' Use WorksheetFunction to calculate Pearson correlation
        With Application.WorksheetFunction

            ' Loop through each pair of columns to calculate their correlation
            For r = 1 To ncols
                ColX = .Index(DataArray, 0, r)
                For c = 1 To ncols
                    ColY = .Index(DataArray, 0, c)
                    Matrix(r, c) = .Pearson(ColX, ColY)
                Next c
            Next r
        End With

    Else
        ' Resize the Matrix array to store the correlation coefficients for rows
        ReDim Matrix(1 To nrows, 1 To nrows)

        ' Use WorksheetFunction to calculate Pearson correlation
        With Application.WorksheetFunction

            ' Loop through each pair of rows to calculate their correlation
            For r = 1 To nrows
                ColX = .Index(DataArray, r, 0)
                For c = 1 To nrows
                    ColY = .Index(DataArray, c, 0)
                    Matrix(r, c) = .Pearson(ColX, ColY)
                Next c
            Next r
        End With
    End If
    
    ' Return the correlation matrix
    CorrMatrix = Matrix
End Function
