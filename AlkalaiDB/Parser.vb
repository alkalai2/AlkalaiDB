'
'Written by J.Alkalai on 12/20/2014
'

Module Parser
    ' Module handles parsing of Excel regions to populate Entry public vars (tname, attributes, values, etc.)

    Public xlApp As Excel.Application = Globals.ThisAddIn.Application


    Public Function getTableValues(ByRef e As Entry)
        ' populates double array with values from selection (values frow two rows under table name)

        Dim vals(e.rows - 2, e.cols) As String
        For i As Integer = 2 To e.rows - 1 
            For j As Integer = 0 To e.cols - 1
                Dim str As String
                str = e.loc.Cells(i + 1, j + 1).value.ToString
                vals(i - 2, j) = str
            Next
        Next
        getTableValues = vals
    End Function


    Public Function getTableAttributes(ByRef e As Entry)
        ' populates array with attribute names (one row under table name)

        Dim attr(e.cols - 1) As String
        Dim str As String
        For i As Integer = 0 To e.cols - 1
            str = e.loc.Cells(2, i + 1).value.ToString
            attr(i) = str

        Next
        getTableAttributes = attr
    End Function

    Public Function getTableTypes(ByRef e As Entry)
        ' Uses first row of values to determine dataype. Converts Excel String to Character(30)
        Dim types(e.cols - 1) As String
        Dim type As String
        For i As Integer = 0 To e.cols - 1
            type = TypeName(e.loc.Cells(3, i + 1).value)
            If type = "String" Then
                type = "character(30)"
            End If
            If type = "Double" Then
                type = "real"
            End If
            types(i) = type

        Next
        getTableTypes = types
    End Function

    
    Public Function getRowValues(r As Excel.Range, len As Integer)
        ' populates the values from a selected single row (used in Insert/Delete Rows)

        If ((r Is Nothing) = False) Then
            Dim vals(len) As String
            For i As Integer = 0 To len - 1
                vals(i) = r.Cells(1, i + 1).value.ToString
            Next
            Return vals
        End If
        Return Nothing
    End Function
End Module
