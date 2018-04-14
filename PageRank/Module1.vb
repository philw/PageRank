Module Module1

    Sub Main()

        Dim PageLinks = {{0, 1, 1, 0},
                         {1, 0, 0, 0},
                         {0, 1, 0, 1},
                         {0, 0, 0, 0}}

        Dim LinkCount = {1, 2, 1, 1}

        Dim PageRank() As Double = {1, 1, 1, 1}
        Dim NumPages As Integer = PageRank.Length
        Dim Rank As Double
        Dim Damping As Double = 0.85
        Dim UserInput As String

        Console.WriteLine("Press enter key to calculate next line")
        Console.WriteLine()

        PrintHeaderRow(PageRank.Length)
        PrintRow(PageRank)
        Console.ReadLine()

        Do
            ' recalculate page ranks
            For Page = 0 To NumPages - 1
                Rank = 0.0
                For Link = 0 To NumPages - 1
                    If Link <> Page And PageLinks(Page, Link) = 1 Then
                        Rank += PageRank(Link) / LinkCount(Link)
                    End If
                Next
                PageRank(Page) = Rank * Damping + (1 - Damping)
            Next
            PrintRow(PageRank)
            UserInput = Console.ReadLine()
        Loop While UserInput = ""


    End Sub

    Sub PrintHeaderRow(Cols)
        Dim C As Char

        For N = 0 To Cols - 1
            C = Chr(Asc("A") + N)
            Console.Write("{0,-10}", C)
        Next
        Console.WriteLine()

    End Sub

    Sub PrintRow(PageRank() As Double)

        For N = 0 To PageRank.Length - 1
            Console.Write("{0,-10:N3}", PageRank(N))
        Next
        'Console.WriteLine()
    End Sub

End Module
