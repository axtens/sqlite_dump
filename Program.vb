Imports System

Module Program
    Sub Main(args As String())
        if args.Count < 1 Then
            Console.WriteLine("sqlite_dump <string>")
            End
        End If
        dim result = BODB.DBSQLite(args(0))
        if result.Item1 <> vbNullString then
            Console.WriteLine(result.Item1)
        else
            Console.WriteLine(result.Item2)
        end if
    End Sub
End Module
