Attribute VB_Name = "Module1"
Function CollatzStep(n As Long) As Long
        If n Mod 2 = 1 Then
            CollatzStep = (3 * n) + 1
        Else
            CollatzStep = n / 2
        End If
    End Function
    

Function CountCollatzSteps(n As Long) As Long
            Dim count As Long
            count = 0
            Do While n <> 1
                n = CollatzStep(n)
                count = count + 1
            Loop
            CountCollatzSteps = count
    End Function

Attribute VB_Name = "Module3"
Function FindMaxIters(maxRow As Long, maxCol As Long) As Long
    Dim maxIters As Long
maxIters = 0
    
    For r = 1 To maxRow
        Dim thisRowIters As Long
        thisRowIters = 0
        For c = 1 To maxCol
            If Cells(r, c) = 1 Then
                thisRowIters = c
                Exit For
            End If
        Next c
        If thisRowIters > maxIters Then
            maxIters = maxRow
        End If
    Next r
    FindMaxIters = maxIters
End Function


Function CollatzSteps(n As Long) As Long
        If n Mod 2 = 1 Then
            CollatzSteps = (3 * n) + 1
        Else
            CollatzSteps = n / 2
        End If
    End Function
    

  


