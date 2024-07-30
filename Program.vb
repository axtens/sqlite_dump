Imports System

Module Program
    Sub Main(args As String())
        If args.Count < 2 Then
            Console.WriteLine("sqlite_dump -csv|-sql <string>")
            End
        End If
        Dim result As (String, String)
        If args(0) = "-csv" Then
            result = BODB.DBCSV(args(1))
        ElseIf args(0) = "-sql" Then
            result = BODB.DBSQLite(args(1))
        Else
            Console.WriteLine(args(0) + " method unknown.")
            End
        End If

        If result.Item1 <> vbNullString then
            Console.WriteLine(result.Item1)
        else
            Console.WriteLine(result.Item2)
        end if
    End Sub
End Module
