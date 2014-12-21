' 
'Written by J.Alkalai on 12/20/14
'

Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1



    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub
   
    
    ' Create a new Table
    Private Sub btn_createTable_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_createTable.Click
         ' Fires up the Create Table Form, the form handles the creation

        Dim createTableForm As CreateTableForm
        createTableForm = New CreateTableForm
        createTableForm.Show()

        
    End Sub

    ' Insert a Row - user needs to highlight a new row of the same width as the table (eg selection.width = table.width)
    Private Sub btn_insertRow_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_insertRow.Click
        ' search through our list of previous entries (in handler.vb)
        Dim found As Entry = find_entry(xlApp.Selection, True)
        
        If (IsNothing(found)) Then
            Return
        End If

        ' insert row into DB table
        found.insertRow(xlApp.ActiveCell, xlApp.Selection.columns.count)
        'update range/size of entry
        'found.range = xlApp.Range(found.loc, found.loc.Cells(found.rows + 1, found.cols))

        ' clear area and repopulate from DB Table
        found.populateTableValues()
        'MsgBox(found.range.Address)
        

    End Sub

    ' Delete a Row - user needs to highlight an existing row to remove 
    Private Sub btn_deleteRow_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_deleteRow.Click
        ' search through our list of previous entries (in handler.vb)
        Dim found As Entry = find_entry(xlApp.Selection, False)
        
        If (IsNothing(found)) Then
            Return
        End If

        ' delete row from DB table
        found.deleteRow(xlApp.ActiveCell, xlApp.Selection.columns.count)
        'update range/size of entry
        
        ' clear area and repopulate from DB table
        found.populateTableValues()
        'MsgBox(found.range.Address)
        
    End Sub


    Private Sub btn_deleteTable_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_deleteTable.Click

        'search through our list of previous entries (in handler.vb)
        Dim found As Entry = remove_entry(xlApp.ActiveCell)
        If IsNothing(found) = False Then
            ' remove DB table from DB
            found.deleteTable()
            found = Nothing
        End If
    End Sub

    
    Private Sub btn_customQuery_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_customQuery.Click
        ' fire Custom Query Form

        Dim customQueryForm As CustomQueryForm
        customQueryForm = New CustomQueryForm
        customQueryForm.Show()
    End Sub
End Class
