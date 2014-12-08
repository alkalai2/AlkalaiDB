Imports System.Collections

Module Handler
    ' keep track of created entries

    Public list_of_entries As New ArrayList()


    ' inserting : Flag that indicates an insertion is being requested. For Deletes, the flag is set to False
    Public Function find_entry(loc As Excel.Range, inserting As Boolean)
        Dim found As Entry
        For Each e As Entry In list_of_entries
            If (inserting = True) Then
                ' for row insertions, row must be directly underneath existing table
                If xlApp.Intersect(loc.Cells(0, 1), e.range) IsNot Nothing And xlApp.Intersect(loc.Cells(1, 1), e.range) Is Nothing Then
                    'check that row of correct size is inputted
                    If (loc.Columns.Count = e.cols) Then
                        found = e
                        MsgBox("found: " + e.tname)
                        Return found
                    End If
                End If
            Else
                ' for row deletions, row must be within table
                If xlApp.Intersect(loc.Cells(1, 1), e.range) IsNot Nothing Then
                    found = e
                    MsgBox("found: " + e.tname)
                    Return found
                End If
            End If
        Next
        MsgBox("Error \n For Insertions please highlight a row directly beneath an existing table \n For Deletions please highilight a row in an existing table")
        Return Nothing
    End Function

    Public Function remove_entry(loc As Excel.Range)
        Dim found As Entry
        For i As Integer = 0 To list_of_entries.Count
            If xlApp.Intersect(loc, list_of_entries(i).range) IsNot Nothing Then
                MsgBox("found table " + list_of_entries(i).tname)
                found = list_of_entries(i)
                list_of_entries.RemoveAt(i)
                Return found
            End If
        Next
        MsgBox(list_of_entries.Count)
        Return Nothing
    End Function
    Public Sub delegateEvent()
        Dim EventDel_CellsChange As Excel.DocEvents_ChangeEventHandler
        Dim xlSheet As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
        EventDel_CellsChange = New Excel.DocEvents_ChangeEventHandler(AddressOf CellsChange)
        AddHandler xlSheet.Change, EventDel_CellsChange
    End Sub

    Private Sub CellsChange(ByVal Target As Excel.Range)
        'This is called when a cell or cells on a worksheet are changed.

        ' see if changed cell is part of a table
        Dim found As Entry
        For Each e As Entry In list_of_entries
            If xlApp.Intersect(Target, e.range) IsNot Nothing Then

                MsgBox("found: " + e.tname)
                found = e
            End If
        Next
    End Sub


End Module
