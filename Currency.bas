Attribute VB_Name = "Currency"

Public Function ConvertCurrencyToEnglish(ByVal MyNumber)
    Dim Temp
    Dim Rs, Paise
    Dim DecimalPlace, Count

    ReDim Place(9) As String
    Place(2) = " Thousand "
    Place(3) = " Lac "
    Place(4) = " Crore "
    Place(5) = " Erb "
    Place(6) = " Kharab "
    Place(7) = " Neel "

    MyNumber = Trim(str(MyNumber))
    DecimalPlace = InStr(MyNumber, ".")
    If DecimalPlace > 0 Then
       Temp = Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2)
       Paise = ConvertTens(Temp)
       MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
    End If

    Count = 1
    Do While MyNumber <> ""
       ' Convert last 3 digits of MyNumber to English Rs.
       If Count = 1 Then
           Temp = ConvertHundreds(Right("000" & MyNumber, 3))
           If Temp <> "" Then Rs = Temp & Place(Count) & Rs
               If Len(MyNumber) > 3 Then
                  ' Remove last 3 converted digits from MyNumber.
                  MyNumber = Left(MyNumber, Len(MyNumber) - 3)
               Else
                  MyNumber = ""
               End If
       Else
           Temp = ConvertTens(Right("00" & MyNumber, 2))
           If Temp <> "" Then Rs = Temp & Place(Count) & Rs
               If Len(MyNumber) > 2 Then
                  ' Remove last 2 converted digits from MyNumber.
                  MyNumber = Left(MyNumber, Len(MyNumber) - 2)
               Else
                  MyNumber = ""
               End If
       End If
       Count = Count + 1
    Loop

    Select Case Rs
       Case ""
          Rs = "Zero Rs"
       Case "One"
          Rs = "One Rs"
       Case Else
          Rs = "Rupees " & Rs
    End Select

    Select Case Paise
       Case ""
          Paise = " "   'REMOVED FOR ZERO PAISE
       Case "One"
          Paise = " And One Paise"
       Case Else
          Paise = " And " & Paise & " Paise"
    End Select
    ConvertCurrencyToEnglish = Rs & Paise & " Only."
End Function

Private Function ConvertHundreds(ByVal MyNumber)
         Dim Result As String
         If Val(MyNumber) = 0 Then Exit Function
         MyNumber = Right("000" & MyNumber, 3)
         If Left(MyNumber, 1) <> "0" Then
            Result = ConvertDigit(Left(MyNumber, 1)) & " Hundred "
         End If
         If Mid(MyNumber, 2, 1) <> "0" Then
            Result = Result & ConvertTens(Mid(MyNumber, 2))
         Else
            Result = Result & ConvertDigit(Mid(MyNumber, 3))
         End If
         ConvertHundreds = Trim(Result)
End Function

Private Function ConvertTens(ByVal MyTens)
         Dim Result As String
         MyNumber = Right("00" & MyNumber, 2)
         If Val(Left(MyTens, 1)) = 1 Then
            Select Case Val(MyTens)
               Case 10: Result = "Ten"
               Case 11: Result = "Eleven"
               Case 12: Result = "Twelve"
               Case 13: Result = "Thirteen"
               Case 14: Result = "Fourteen"
               Case 15: Result = "Fifteen"
               Case 16: Result = "Sixteen"
               Case 17: Result = "Seventeen"
               Case 18: Result = "Eighteen"
               Case 19: Result = "Nineteen"
               Case Else
            End Select
         Else
            Select Case Val(Left(MyTens, 1))
               Case 2: Result = "Twenty "
               Case 3: Result = "Thirty "
               Case 4: Result = "Forty "
               Case 5: Result = "Fifty "
               Case 6: Result = "Sixty "
               Case 7: Result = "Seventy "
               Case 8: Result = "Eighty "
               Case 9: Result = "Ninety "
               Case Else
            End Select
            Result = Result & ConvertDigit(Right(MyTens, 1))
         End If

         ConvertTens = Result
End Function

Private Function ConvertDigit(ByVal MyDigit)
         Select Case Val(MyDigit)
            Case 1: ConvertDigit = "One"
            Case 2: ConvertDigit = "Two"
            Case 3: ConvertDigit = "Three"
            Case 4: ConvertDigit = "Four"
            Case 5: ConvertDigit = "Five"
            Case 6: ConvertDigit = "Six"
            Case 7: ConvertDigit = "Seven"
            Case 8: ConvertDigit = "Eight"
            Case 9: ConvertDigit = "Nine"
            Case Else: ConvertDigit = ""
         End Select
End Function



