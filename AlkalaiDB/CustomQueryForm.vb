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
        If (IsNothing(reader)) Then
            Return
        End If


        Dim newAttributes(reader.FieldCount - 1) As String
        Dim newTypes(reader.FieldCount - 1) As String
        Dim count As Integer
        While reader.Read()
            For i As Integer = 0 To reader.FieldCount - 1
                newAttributes(i) = reader.GetName(i)
                newTypes(i) = reader.GetProviderSpecificFieldType(i).ToString
                'MsgBox(newAttributes(i))
                'MsgBox(newTypes(i))
                With loc.Cells(count + 1 + offset, i + 1)
                    .value = reader.Item(i)
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Color = Color.LightGray
                    .Borders(Excel.XlBordersIndex.xlEdgeTop).Color = Color.LightGray
                    .Borders(Excel.XlBordersIndex.xlEdgeRight).Color = Color.LightGray
                    .Borders(Excel.XlBordersIndex.xlEdgeLeft).Color = Color.LightGray

                End With

            Next
            count = count + 1
        End While

        range = xlApp.Range(loc, loc.Cells(count + offset, reader.FieldCount))

        ' if making a new table, set newtable name + attributes
        If (newTable) Then


            With loc
                .Value = newTname.ToUpper
                .Font.Bold = True
            End With


            For i = 0 To reader.FieldCount - 1
                With loc.Cells(2, i + 1)
                    .value = newAttributes(i)
                    .Borders(Excel.XlBordersIndex.xlEdgeBottom).Color = Color.Black
                End With
            Next

            entry = New Entry(range, newAttributes)
            entry.onChangeEvent()
            entry.allowEventChanges = True
            entry.types = getTableTypes(entry)
            list_of_entries.Add(entry)


            MsgBox("Table '" + entry.tname.ToUpper + "' created")
            MsgBox("Table address : " + entry.range.Address)
            'styles
            range.Interior.Color = ColorTranslator.FromHtml("#F2F8FC")
        End If

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