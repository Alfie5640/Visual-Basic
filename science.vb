Module Module1
    Structure sciMark
        Dim subject As String
        Dim year As Integer
        Dim numCandidates As Integer
        Dim numPasses As Integer
    End Structure
    Sub Main()
        Dim sciMarks(23) As sciMark
        Dim passRate(23) As Decimal
        Dim pCount, cCount, bCount As Integer

        getFileData(sciMarks)

        Console.WriteLine("The number of all class passes was " & findPasses(sciMarks))

        Console.WriteLine("The number of times biology had between 10 and 15 pupils was " & findBiologyCand(sciMarks))
        Console.WriteLine("The number of times chemistry had between 10 and 15 pupils was " & findChemistryCand(sciMarks))
        Console.WriteLine("The number of times physics had between 10 and 15 pupils was " & findPhysicsCand(sciMarks))
        Console.WriteLine("")
        calculateBestPasses(sciMarks, pCount, cCount, bCount)

        Console.WriteLine("Physics had a better pass number " & pCount & " times")
        Console.WriteLine("Biology had a better pass number " & bCount & " times")
        Console.WriteLine("Chemistry had a better pass number " & cCount & " times")

        Console.WriteLine("")

        findPassRate(sciMarks, passRate)
        bestWorstPassRate(sciMarks, passRate)
        leastCandidates(sciMarks)
        findMaximum(sciMarks)
        findMinimum(sciMarks)
        findMin2016(sciMarks)
    End Sub

    Sub calculateBestPasses(ByVal sciMarks() As sciMark, ByRef pCount As Integer, ByRef cCount As Integer, ByRef bCount As Integer)
        Dim phyPass, chemPass, bioPass As Integer
        pCount = 0
        cCount = 0
        bCount = 0
        For selectYear As Integer = 2010 To 2017
            phyPass = 0
            bioPass = 0
            chemPass = 0
            For index As Integer = 0 To 23
                If selectYear = sciMarks(index).year Then
                    If sciMarks(index).subject = "Physics" Then
                        phyPass = sciMarks(index).numPasses
                    ElseIf sciMarks(index).subject = "Chemistry" Then
                        chemPass = sciMarks(index).numPasses
                    ElseIf sciMarks(index).subject = "Biology" Then
                        bioPass = sciMarks(index).numPasses
                    End If
                End If
            Next
            If phyPass > chemPass And phyPass > bioPass Then
                pCount = pCount + 1
            ElseIf chemPass > phyPass And chemPass > bioPass Then
                cCount = cCount + 1
            ElseIf bioPass > phyPass And bioPass > chemPass Then
                bCount = bCount + 1
            End If
        Next
    End Sub
    Function findChemistryCand(ByVal sciMarks() As sciMark)
        Dim numOfChem As Integer = 0
        For index As Integer = 0 To 23
            If sciMarks(index).numCandidates >= 10 And sciMarks(index).numCandidates <= 15 And sciMarks(index).subject = "Chemistry" Then
                numOfChem = numOfChem + 1
            End If
        Next
        Return numOfChem
    End Function

    Function findPhysicsCand(ByVal sciMarks() As sciMark)
        Dim numOfPhy As Integer = 0
        For index As Integer = 0 To 15
            If sciMarks(index).numCandidates >= 10 And sciMarks(index).numCandidates <= 15 And sciMarks(index).subject = "Physics" Then
                numOfPhy = numOfPhy + 1
            End If
        Next
        Return numOfPhy
    End Function
    Function findBiologyCand(ByVal sciMarks() As sciMark)
        Dim numOfBio As Integer = 0
        For index As Integer = 0 To 7
            If sciMarks(index).numCandidates >= 10 And sciMarks(index).numCandidates <= 15 Then
                numOfBio = numOfBio + 1
            End If
        Next
        Return numOfBio
    End Function
    Function findPasses(ByVal sciMarks() As sciMark) As Integer
        Dim numOfPasses As Integer = 0
        For index = 0 To 23
            If sciMarks(index).numCandidates = sciMarks(index).numPasses Then
                numOfPasses = numOfPasses + 1
            End If
        Next
        Return numOfPasses
    End Function

    Sub leastCandidates(ByVal sciMarks() As sciMark)
        Dim totalCandidates() As Integer = {0, 0, 0, 0, 0, 0, 0, 0}
        Dim years() As Integer = {2017, 2016, 2015, 2014, 2013, 2012, 2011, 2010}
        Dim currentYear As Integer = 2017
        Dim num As Integer = 0
        Dim leastCand As Integer = 1000000
        Dim leastPos As Integer = 0
        Do
            For index As Integer = 0 To 23
                If sciMarks(index).year = currentYear Then
                    totalCandidates(num) = totalCandidates(num) + sciMarks(index).numCandidates
                End If
            Next
            num = num + 1
            currentYear = currentYear - 1
        Loop Until currentYear < 2010
        For index As Integer = 0 To 7
            If totalCandidates(index) < leastCand Then
                leastCand = totalCandidates(index)
                leastPos = index
            End If
        Next
        Console.WriteLine("The year " & years(leastPos) & " had the least number of candidates with only " & totalCandidates(leastPos) & " people.")
    End Sub

    Sub bestWorstPassRate(ByVal sciMarks() As sciMark, ByVal passRate() As Decimal)
        Dim maxRate As Decimal = passRate(0)
        Dim maxPos As Integer = 0
        For index As Integer = 1 To 23
            If passRate(index) > maxRate Then
                maxRate = passRate(index)
                maxPos = index
            End If
        Next
        Console.WriteLine("The best pass rate was " & maxRate & "% in " & sciMarks(maxPos).year & " in " & sciMarks(maxPos).subject)

        Dim minRate As Decimal = passRate(0)
        Dim minPos As Integer = 0
        For index As Integer = 1 To 23
            If passRate(index) < minRate Then
                minRate = passRate(index)
                minPos = index
            End If
        Next
        Console.WriteLine("The worst pass rate was " & minRate & "% in " & sciMarks(minPos).year & " in " & sciMarks(minPos).subject)
    End Sub
    Sub findPassRate(ByVal sciMarks() As sciMark, ByRef passRate() As Decimal)
        For index As Integer = 0 To 23
            passRate(index) = (sciMarks(index).numPasses / sciMarks(index).numCandidates) * 100
        Next
    End Sub
    Sub findMin2016(ByVal sciMarks() As sciMark)
        Dim minVal As Integer = 100
        Dim minPos As Integer = -1
        For index As Integer = 0 To 23
            If sciMarks(index).numCandidates < minVal And sciMarks(index).year = 2016 Then 'new min found in 2016
                minVal = sciMarks(index).numCandidates
                minPos = index
            End If
        Next
        Console.WriteLine("The lowest number of candidates in 2016 was " & minVal & " in " & sciMarks(minPos).subject)
    End Sub
    Sub findMinimum(ByVal sciMarks() As sciMark)
        Dim minVal As Integer = sciMarks(0).numCandidates 'first value is current min
        Dim minPos As Integer = 0
        For index As Integer = 1 To 23
            If sciMarks(index).numCandidates < minVal Then 'new min found
                minVal = sciMarks(index).numCandidates
                minPos = index
            End If
        Next
        Console.WriteLine("The lowest number of candidates was " & minVal & " in " & sciMarks(minPos).subject & " in " & sciMarks(minPos).year)
    End Sub
    Sub findMaximum(ByVal sciMarks() As sciMark)
        Dim maxVal As Integer = sciMarks(0).numPasses 'first value is current max
        Dim maxPos As Integer = 0
        For index As Integer = 1 To 23
            If sciMarks(index).numPasses > maxVal Then 'new max found
                maxVal = sciMarks(index).numPasses
                maxPos = index
            End If
        Next
        Console.WriteLine("The highest number of passes was " & maxVal & " in " & sciMarks(maxPos).subject & " in " & sciMarks(maxPos).year)
    End Sub
    Sub getFileData(ByRef sciMarks() As sciMark)
        Dim recordIndex As Integer = 0
        Using myReader As New Microsoft.VisualBasic.FileIO.TextFieldParser("../../scienceData.txt")
            myReader.TextFieldType = FileIO.FieldType.Delimited
            myReader.SetDelimiters(",") ' comma is the delimiter in the file
            Dim currentRow() As String
            While Not myReader.EndOfData
                currentRow = myReader.ReadFields()
                sciMarks(recordIndex).subject = currentRow(0)
                sciMarks(recordIndex).year = currentRow(1)
                sciMarks(recordIndex).numCandidates = currentRow(2)
                sciMarks(recordIndex).numPasses = currentRow(3)
                recordIndex = recordIndex + 1
            End While
            myReader.Close()
        End Using
    End Sub
End Module

