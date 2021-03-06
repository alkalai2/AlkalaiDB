﻿'
'Written by J.Alkalai 12/20/14
'

Imports System.Collections
Imports Npgsql

Module Handler
    
    ' Module holds list of previously created entries
    ' also includes a test 'Execute SQL' funct that is only used in the Custom Query Form (meant to implement throughout)


    Public list_of_entries As New ArrayList()

    ' Search through list for prev created tables (used in Insert/Delete rows)
    ' params: loc  - location of first cell of selection
    '         inserting - True if called by InsertRow, false if by DeleteRow
    Public Function find_entry(loc As Excel.Range, inserting As Boolean)
        Dim found As Entry
        For Each e As Entry In list_of_entries
            If (inserting = True) Then
                ' for row insertions, row must be directly underneath existing table
                If xlApp.Intersect(loc(1).Cells(0, 1), e.range) IsNot Nothing And xlApp.Intersect(loc(1).Cells(1, 1), e.range) Is Nothing Then
                    'check that row of correct size is inputted

                    If (loc.Columns.Count = e.cols) Then
                        found = e
                        'MsgBox("found: " + e.tname)
                        Return found
                    End If
                End If
            Else
                ' for row deletions, row must be within table
                If xlApp.Intersect(loc(1).Cells(1, 1), e.range) IsNot Nothing Then
                    found = e
                    'MsgBox("found: " + e.tname)
                    Return found
                End If
            End If
        Next
        MsgBox("Error: For Insertions please highlight a row directly beneath an existing table. For Deletions please highilight a row in an existing table")
        Return Nothing
    End Function

    ' find Entry in list and remove it. Used in 'DeleteTable'
    Public Function remove_entry(loc As Excel.Range)
        Dim found As Entry
        For i As Integer = 0 To list_of_entries.Count
            If xlApp.Intersect(loc, list_of_entries(i).range) IsNot Nothing Then
                'MsgBox("found table " + list_of_entries(i).tname)
                found = list_of_entries(i)
                list_of_entries.RemoveAt(i)
                Return found
            End If
        Next
        MsgBox(list_of_entries.Count)
        Return Nothing
    End Function


    ' Test function to put execution of SQL in single place. So far only used in CustomQuery.vb
    Public Function executeSQL(sql As String, Optional ByVal newTname As String = Nothing)

        Dim connection As NpgsqlConnection = New NpgsqlConnection()
        Dim command, command2 As NpgsqlCommand
        Dim reader As NpgsqlDataReader
        connection.ConnectionString = "Server=localhost;Port=5432;Database=MyDB;User Id=postgres;Password=Oijoij123;"
        connection.Open()

        'command = New NpgsqlCommand(sql, connection)

        Try
            'command.ExecuteNonQuery()

            'execute and populate from query
            If (IsNothing(newTname)) Then
                command = New NpgsqlCommand(sql, connection)

                reader = command.ExecuteReader()
            Else
                sql = Replace(sql.ToLower, "from", " into table " + newTname + " from ")
                command = New NpgsqlCommand(sql, connection)
                command.ExecuteNonQuery()
                command2 = New NpgsqlCommand("select * from " + newTname + ";", connection)
                command2.ExecuteNonQuery()
                reader = command2.ExecuteReader

            End If
        Catch ex As NpgsqlException
            MsgBox(ex.BaseMessage)
            Return Nothing
            Exit Function
        End Try
        MsgBox("Executing:  ' " + sql + " ' ")


        Return reader

    End Function


End Module
