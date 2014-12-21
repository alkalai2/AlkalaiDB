Imports Npgsql
Imports System.Diagnostics
Imports System.Collections
Imports System.Drawing



Public Class Entry

    '   
    '   Will hold all the data pretaining to an inputted table object
    '       loc    - location of table ( upper-left cell of selection area )
    '       tname  - Table Name
    '       attr   - array of attribute values (columns)
    '       types  - array of attribute types
    '       constr - array of attribute constraints (PRIMARY KEY, NOT NULL)
    '       rows   - row dimension of table
    '       cols   - column dimenions of table

    Public xlApp As Excel.Application = Globals.ThisAddIn.Application



    Public loc As Excel.Range
    Public range As Excel.Range
    Public tname As String
    Public attr As String()
    Public types As String()
    Public constr As String()
    Public rows As Integer
    Public cols As Integer

    ' to be able to enable / disable change listener
    Public allowEventChanges = False


    Public Sub New(r As Excel.Range, Optional ByVal attr As String() = Nothing)
        range = r
        loc = r(1)
        tname = r(1).value.ToString
        rows = r.Rows.Count
        cols = r.Columns.Count


        If (IsNothing(attr)) Then
            Me.attr = getTableAttributes(Me)
        Else
            Me.attr = attr
        End If

        types = getTableTypes(Me)

        ' intialize constr to empty strings
        Dim temp(cols) As String
        For i As Integer = 0 To cols - 1
            temp(i) = " "
        Next
        constr = temp


    End Sub




    ' ============================== Database Functions =====================================
    '
    '   createTable - creates new DB table. Then fills DB table with values
    '   insertRow
    '   deleteRow
    '   updateRow
    '   populateTableValues - pulls data from DB table and  populates selected Excel area. 
    '   executeQuery
    '
    '
    '   Note: the initial creation of a table in Excel results in a createTable() call and a populateTableValues() call 
    '

    Public Function createTable()

        Dim connection As NpgsqlConnection
        Dim command As NpgsqlCommand
        Dim sql As String
        Try
            connection = New NpgsqlConnection()
           
        Catch e As NpgsqlException
            MsgBox(e.BaseMessage)
            Return Nothing
        End Try


        connection.ConnectionString = "Server=localhost;Port=5432;Database=MyDB;User Id=postgres;Password=Oijoij123;"
        connection.Open()


        Dim size As Integer = cols

        ' use parser to form arrays
        attr = getTableAttributes(Me)
        types = getTableTypes(Me)

        ' create SQL statement
        sql = "CREATE TABLE " + tname + " ("
        Dim sep As String = ""
        For i As Integer = 0 To size - 1
            sql = sql + sep + attr(i) + " " + types(i) + " " + constr(i)
            sep = ", "
        Next
        sql = sql + ");"

        ' execute SQL
        Command = New NpgsqlCommand(sql, connection)
        Try
            Command.ExecuteNonQuery()
            'MsgBox("executed  " + sql)
        Catch e As NpgsqlException
            MsgBox(e.BaseMessage)
            createTable = 0
            Exit Function
        End Try


        ' Table created, now use values to populate DB table
        ' Dim vals As String(,) = getTableValues(Me)
        Dim vals(cols) As String
        For i As Integer = 3 To rows

            vals = getRowValues(range.Cells(i, 1), cols)
            ' create SQL insert statement for each row
            sep = ""
            sql = "INSERT INTO " + tname + " Values("
            For j As Integer = 0 To cols - 1
                If InStr(types(j), "character(30)") > 0 Then
                    sql = sql + sep + " ' " + vals(j) + " ' "
                Else
                    sql = sql + sep + vals(j)

                End If
                sep = ", "
            Next

            'execute SQL
            sql = sql + ");"
            Command = New NpgsqlCommand(sql, connection)
            Try
                Command.ExecuteNonQuery()
                'MsgBox("executed  " + sql)
            Catch e As NpgsqlException
                MsgBox(e.BaseMessage)
                createTable = 0
                Exit Function
            End Try
        Next

        MsgBox("Table '" + tname.ToUpper + "' created")

        'add entry to our list
        Me.onChangeEvent()
        allowEventChanges = True
        list_of_entries.Add(Me)
        createTable = 1
    End Function

    Public Function deleteTable()

        Dim connection As NpgsqlConnection = New NpgsqlConnection()
        Dim command As NpgsqlCommand
        Dim sql As String

        Try
            connection.ConnectionString = "Server=localhost;Port=5432;Database=MyDB;User Id=postgres;Password=Oijoij123;"
            connection.Open()

            sql = "DROP table " + tname + ";"

            command = New NpgsqlCommand(sql, connection)

            command.ExecuteNonQuery()
            ' MsgBox("executed  " + sql)
        Catch ex As NpgsqlException
            MsgBox(ex.BaseMessage)
            Return 0
        End Try

        'removing table, disable event handler
        allowEventChanges = False
        range.Clear()
        MsgBox(tname + " deleted")
        deleteTable = 1

    End Function

     
    Public Function insertRow(r As Excel.Range, len As Integer, Optional ByVal input As String() = Nothing)
        Dim connection As NpgsqlConnection = New NpgsqlConnection()
        Dim command As NpgsqlCommand
        Dim sql As String

        Try
            allowEventChanges = False
            connection.ConnectionString = "Server=localhost;Port=5432;Database=MyDB;User Id=postgres;Password=Oijoij123;"
            connection.Open()
        Catch e As NpgsqlException
            MsgBox(e.BaseMessage)
            Return Nothing
        End Try
        Dim vals(len) As String
        If (IsNothing(input)) Then
            vals = getRowValues(r, len)
        Else
            vals = input
        End If

        ' create SQL insert statement for each row
        Dim sep As String = ""
        sql = "INSERT INTO " + tname + " Values("
        For j As Integer = 0 To cols - 1
            If InStr(types(j), "character(30)") > 0 Then
                sql = sql + sep + " ' " + vals(j) + " ' "
            Else
                sql = sql + sep + vals(j)

            End If
            sep = ", "
        Next

        'execute SQL
        sql = sql + ");"
        command = New NpgsqlCommand(sql, connection)
        Try
            command.ExecuteNonQuery()
            'MsgBox("executed  " + sql)
        Catch ex As NpgsqlException
            MsgBox(ex.BaseMessage)
            insertRow = 0
            Exit Function
        End Try

        ' increase table range by 1
        Me.range = xlApp.Range(loc, loc.Cells(rows + 1, cols))
        Me.rows = rows + 1
        insertRow = 1
    End Function


    Public Function deleteRow(r As Excel.Range, len As Integer)
        Dim connection As NpgsqlConnection = New NpgsqlConnection()
        Dim command As NpgsqlCommand
        Dim sql As String
        Dim temp As String

        'disable event listener
        allowEventChanges = False

        Try
            connection.ConnectionString = "Server=localhost;Port=5432;Database=MyDB;User Id=postgres;Password=Oijoij123;"
            connection.Open()
        Catch e As NpgsqlException
            MsgBox(e.BaseMessage)
            Return Nothing
        End Try
        Dim vals As String() = getRowValues(r, len)
        temp = r.Address


        ' construct SQL

        sql = "DELETE from " + tname + " where "

        Dim foundPK = False
        Dim sep As String = ""
        For i As Integer = 0 To constr.Count - 1
            If (InStr(constr(i), "PRIMARY KEY")) Then
                foundPK = True
                If InStr(types(i), "character(30)") > 0 Then
                    ' add quotes for strings
                    sql = sql + sep + attr(i) + " = '" + vals(i) + "' "
                Else
                    sql = sql + sep + attr(i) + " = " + vals(i)
                End If

                sep = "AND"
            End If
        Next

        ' if no primary key, use first attribute
        If (foundPK = False) Then
            If InStr(types(0), "character(30)") > 0 Then
                ' add quotes for strings
                sql = sql + sep + attr(0) + " = '" + vals(0) + "' "
            Else
                sql = sql + sep + attr(0) + " = " + vals(0)
            End If
        End If

        ' Execute SQL
        command = New NpgsqlCommand(sql, connection)
        '  MsgBox("executed  " + sql)
        Try
            command.ExecuteNonQuery()

        Catch ex As NpgsqlException
            MsgBox(ex.BaseMessage)
            deleteRow = 0
            Exit Function
            Return Nothing

        End Try

        'clear range
        Dim toDelete As Excel.Range = xlApp.Range(r, r.Cells(rows, cols))
        toDelete.Clear()

        'range decreases by 1
        Me.range = xlApp.Range(loc, loc.Cells(rows - 1, cols))
        Me.rows = rows - 1
        allowEventChanges = True
        deleteRow = 1
    End Function

    Public Sub populateTableValues()
        ' used to repopulate an Excel region with DB table values

        Dim sheet As Excel.Worksheet = xlApp.ActiveSheet
        
        Dim connection As NpgsqlConnection = New NpgsqlConnection()
        Dim command As NpgsqlCommand
        Dim sql As String

        Try
            connection.ConnectionString = "Server=localhost;Port=5432;Database=MyDB;User Id=postgres;Password=Oijoij123;"
            connection.Open()
        Catch ex As NpgsqlException
            MsgBox(ex.BaseMessage)
            Return
        End Try

            ' Construct SQL, get all values from table

            sql = "SELECT * FROM " + tname + ";"

            command = New NpgsqlCommand(sql, connection)
            Try
                command.ExecuteNonQuery()
                ' MsgBox("populating: " + sql)
            Catch e As NpgsqlException
                MsgBox(e.BaseMessage)
                Exit Sub
            End Try

            ' use reader to get ouput of SQL above
            Dim reader As NpgsqlDataReader = command.ExecuteReader


            Dim count As Integer = 0

            ' Read value from db and populate excel
            While reader.Read()
                For i As Integer = 0 To reader.FieldCount - 1
                    allowEventChanges = False
                    With loc.Cells(count + 3, i + 1)
                        .value = reader.Item(i)
                        'set styles
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


            ' set more styles
            range.Interior.Color = ColorTranslator.FromHtml("#F2F8FC")
            'range.Borders(Excel.XlLineStyle.xlContinuous).Color = Color.LightGray
            loc.Font.Bold = True
            loc.Value = Me.tname.ToUpper
            Dim r As Excel.Range = xlApp.Range(loc.Cells(2, 1), loc.Cells(2, reader.FieldCount)) ' attrib
            r.Borders(Excel.XlBordersIndex.xlEdgeBottom).Color = Color.Black

            allowEventChanges = True
    End Sub



    ' Listening Event setup, used to provide DB changes when a single Excel value is changed

    Public Sub onChangeEvent()
        Dim EventDel_CellsChange As Excel.DocEvents_ChangeEventHandler
        Dim xlSheet As Excel.Worksheet = Globals.ThisAddIn.Application.ActiveSheet
        EventDel_CellsChange = New Excel.DocEvents_ChangeEventHandler(AddressOf CellsChange)
        AddHandler xlSheet.Change, EventDel_CellsChange
    End Sub

    Private Sub CellsChange(ByVal Target As Excel.Range)
        'This is called when a cell or cells on a worksheet are changed.
        Dim temp = allowEventChanges
        If (allowEventChanges = True) Then
            allowEventChanges = False

            ' see if changed cell is part of a table
            Dim valueRange As Excel.Range = xlApp.Range(loc.Cells(3, 1), loc.Cells(rows, cols))
            'MsgBox(valueRange.Address)
            If xlApp.Intersect(Target, valueRange) IsNot Nothing Then
                'MsgBox("found: " + tname)
                Dim rowOffset = Target.Row - loc.Row + 1
                Dim updateRange = xlApp.Range(loc.Cells(rowOffset, 1), loc.Cells(rowOffset, cols))
                'MsgBox("updating range: " + updateRange.Address)

                'collect values in updating range
                Dim vals(cols - 1) As String
                For i = 0 To cols - 1
                    vals(i) = updateRange.Cells(1, i + 1).value
                Next


                deleteRow(updateRange, cols)
                insertRow(updateRange, cols, vals)
                populateTableValues()


            End If
            allowEventChanges = True
        End If


        Exit Sub
    End Sub


End Class

