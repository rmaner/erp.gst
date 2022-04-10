Attribute VB_Name = "ISBN"
Public Function EAN13(ISBN)
    Dim i, checksum, first, BarCode, tableA As Boolean
    EAN13 = ""
    
    
    ISBN = Replace(ISBN, " ", "")
    ISBN = "978" & Left(Replace(ISBN, "-", ""), 9)
    
    'Check for 12 characters
    If Len(ISBN) = 12 Then
        'And they are really digits
        For i = 1 To 12
            If Asc(Mid$(ISBN, i, 1)) < 48 Or Asc(Mid$(ISBN, i, 1)) > 57 Then
                i = 0
                Exit For
            End If
        Next
        
        If i = 13 Then
            'Calculation of the checksum
            For i = 2 To 12 Step 2
                checksum = checksum + Val(Mid$(ISBN, i, 1))
            Next
            
            checksum = checksum * 3
            For i = 1 To 11 Step 2
                checksum = checksum + Val(Mid$(ISBN, i, 1))
            Next
            
            ISBN = ISBN & (10 - checksum Mod 10) Mod 10
            
            'The first digit is taken just as it is, the second one come from table A
            BarCode = Left$(ISBN, 1) & Chr$(65 + Val(Mid$(ISBN, 2, 1)))
            first = Val(Left$(ISBN, 1))
            For i = 3 To 7
                tableA = False
                Select Case i
                    Case 3
                        Select Case first
                            Case 0 To 3
                            tableA = True
                        End Select
                    Case 4
                        Select Case first
                            Case 0, 4, 7, 8
                            tableA = True
                        End Select
                    Case 5
                        Select Case first
                            Case 0, 1, 4, 5, 9
                            tableA = True
                        End Select
                    Case 6
                        Select Case first
                            Case 0, 2, 5, 6, 7
                            tableA = True
                        End Select
                    Case 7
                        Select Case first
                            Case 0, 3, 6, 8, 9
                            tableA = True
                        End Select
                End Select
                
                If tableA Then
                    BarCode = BarCode & Chr$(65 + Val(Mid$(ISBN, i, 1)))
                Else
                    BarCode = BarCode & Chr$(75 + Val(Mid$(ISBN, i, 1)))
                End If
            Next
            
            ' Add middle separator
            BarCode = BarCode & "*"
            
            For i = 8 To 13
                BarCode = BarCode & Chr$(97 + Val(Mid$(ISBN, i, 1)))
            Next
            
            ' Add end mark
            BarCode = BarCode & "+"
            EAN13 = BarCode
        End If
    End If
End Function

Public Function ParseISBN(ByVal ISBN As String) As String
    ISBN = Replace(ISBN, " ", "")
    ISBN = Replace(ISBN, "-", "")
    ISBN = Left(ISBN, 9)
    ParseISBN = ISBN
End Function
