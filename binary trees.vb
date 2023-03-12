Module Module1
    Dim Data(5) As String
    Dim LeftPointer(5) As Integer
    Dim RightPointer(5) As Integer
    Dim SearchItem As String
    Sub CreateArrays()
        Data(0) = "maths"
        Data(1) = "polit"
        Data(2) = "biolog"
        Data(3) = "physics"
        Data(4) = "comput"

        LeftPointer(0) = 2
        LeftPointer(1) = 4
        LeftPointer(2) = -1
        LeftPointer(3) = -1
        LeftPointer(4) = -1

        RightPointer(0) = 1
        RightPointer(1) = -1
        RightPointer(2) = 4
        RightPointer(3) = -1
        RightPointer(4) = -1
    End Sub
    Sub DisplayArrays()
        For i = 0 To 5
            Console.WriteLine(LeftPointer(i) & vbTab & Data(i) & vbTab & RightPointer(i))
        Next
    End Sub
    Sub SearchArray(SearchItem)
        Dim current As Integer
        Dim found As Boolean = False
        Console.WriteLine("Enter item to search for")
        SearchItem = Console.ReadLine()
        current = 0
        Dim Null As Integer = -1
        Do While current <> -1 And found = False
            If Data(current) = SearchItem Then
                found = True
                Console.WriteLine(SearchItem & " was found")
            Else
                If Data(current) > SearchItem Then
                    current = LeftPointer(current)
                Else
                    current = RightPointer(current)
                End If
            End If
        Loop
        If found = False Then
            Console.WriteLine(SearchItem & " was not found")
        End If
    End Sub
    Sub InsertItem()
        Dim Current, PreviousIndex, NullPointer, Freespace As Integer
        Dim Found As Boolean
        Dim NewData As String
        Freespace = 5
        NullPointer = -1
        Found = False
        Console.WriteLine("enter new item")
        NewData = Console.ReadLine
        If Data(0) = "" Then
            Data(0) = NewData
        Else
            Current = 0
            Do While Current <> -1 And Found = False
                If Data(Current) > NewData Then
                    PreviousIndex = Current
                    Current = LeftPointer(Current)
                Else
                    PreviousIndex = Current
                    Current = RightPointer(Current)
                End If
            Loop
            Data(Freespace) = NewData
            Console.WriteLine(PreviousIndex & vbTab & Current & " Loop finished")
        End If
    End Sub
    Sub Main()
        Call CreateArrays()
        Call DisplayArrays()
        Call SearchArray(SearchItem)
        Call InsertItem()
        Call DisplayArrays()
        Console.ReadLine()
    End Sub

End Module
