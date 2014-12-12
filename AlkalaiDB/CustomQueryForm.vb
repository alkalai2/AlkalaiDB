Imports Npgsql
Imports System.Drawing

Public Class CustomQueryForm
    Public newTable As Boolean

    Private Sub btn_createCustomQuery_Click(sender As Object, e As EventArgs) Handles btn_createCustomQuery.Click
        Dim reader As NpgsqlDataReader
        Dim range As Excel.Range
        Dim entry As Entry
        Dim loc As Excel.Range = xlApp.ActiveCell



        Dim offset As Integer = 0
        Dim newTname As String = Nothing

        newTable = check_customNewTable.Checked

        If (newTable) Then
            newTname = txt_customName.Text
            offset = 2
        End If


        reader = executeSQL(txt_customQuery.Text, newTname)

        Dim newAttributes(reader.FieldCount) As String
        Dim count As Integer
        While reader.Read()
            For i As Integer = 0 To reader.FieldCount - 1
                newAttributes(i) = reader.GetName(i)

                With loc.Cells(count + 1 + offset, i + 1)
                    .value = reader.Item(i)
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Color = Color.LightGray
                    .Borders(Excel.XlBordersIndex.xlEdgeTop).Color = Color.LightGray
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).Color = Color.LightGray
                    .Borders(Excel.XlBordersIndex.xlEdgeLeft).Color = Color.LightGray

                End With

                ' loc.Cells(count + 3, i + 1).value = reader.Item(i)
                ' Borders(Excel.XlBordersIndex.xlEdgeBottom).Color = Color.Gray
                'apply style to data cells
                ' loc.Cells(count + 3, i + 1).style = "ValueStyle"

            Next
            count = count + 1
        End While

        ' if making a new table, set newtable name + attributes
        If (newTable) Then
            loc.Value = newTname
            For i = 0 To reader.FieldCount - 1
                loc.Cells(2, i + 1).value = newAttributes(i)
            Next
        End If

        range = xlApp.Range(loc, loc.Cells(count + offset, reader.FieldCount))
        Dim v As String = range.Address
        ' create entry element for new table
        entry = New Entry(range)
        Me.Close()
    End Sub

    
    Private Sub check_customNewTable_CheckedChanged(sender As Object, e As EventArgs) Handles check_customNewTable.CheckedChanged
        If check_customNewTable.CheckState = False Then
            label_customName.Visible = False
            txt_customName.Visible = False
            newTable = False
        Else
            label_customName.Visible = True
            txt_customName.Visible = True
            newTable = True
        End If
    End Sub

    Private Sub CustomQueryForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class