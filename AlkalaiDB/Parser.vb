Module Parser
    Public xlApp As Excel.Application = Globals.ThisAddIn.Application


    Public Function getTableValues(ByRef e As Entry)

        Dim vals(e.rows - 2, e.cols) As String
        For i As Integer = 2 To e.rows - 1 ' loc at first row, collect data to send to parser (sep attr from vals)
            For j As Integer = 0 To e.cols - 1
                Dim str As String
                str = e.loc.Cells(i + 1, j + 1).value.ToString
                vals(i - 2, j) = str
            Next
        Next
        getTableValues = vals
    End Function


    Public Function getTableAttributes(ByRef e As Entry)
        Dim attr(e.cols - 1) As String
        Dim str As String
        For i As Integer = 0 To e.cols - 1
            str = e.loc.Cells(2, i + 1).value.ToString
            attr(i) = str

        Next
        getTableAttributes = attr
    End Function

    Public Function getTableTypes(ByRef e As Entry)
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

    ' param: the left-most cell of excel row ; length of inputted row
    Public Function getRowValues(r As Excel.Range, len As Integer)

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
