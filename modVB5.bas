Attribute VB_Name = "modVB5"
'*************************************'
'*      Created by Michael Canejo    *'
'*    Email: mikecanejo@hotmail.com  *'
'*           AIM: Mikey3dd           *'
'*************************************'
'* Name: modVB5.bas   April 18, 2003 *'
'*************************************'

'Here are some unsupported vb5 functions
'that vb6 has built in. These functions give
'100% functionality of their vb6 parents.

'There is no commenting but my email is at
'the top if you have questions or need help :)

'Added Round() and StrReverse() below, enjoy!
'Edited on April 18, 2003


Option Explicit


Public Function Round(ByRef Number As Variant, _
    Optional ByRef NumDigitsAfterDecimal As Long = 0) As Variant
    
    Dim lngPos      As Long
    Dim dblRound    As Double
    
    lngPos = InStr(1, Number, ".")
    
    If lngPos = 0 Then
        Round = Number
        Exit Function
    End If
    
    dblRound = Left(Number, lngPos + NumDigitsAfterDecimal)
    
    If Len(Number) - lngPos >= NumDigitsAfterDecimal Then

        If Mid(Number, lngPos + NumDigitsAfterDecimal + 1, 1) > 5 Then
            If NumDigitsAfterDecimal > 0 Then
    
                dblRound = CDbl(dblRound) + CDbl("." _
                & String(NumDigitsAfterDecimal - 1, "0") & "1")
            Else
                dblRound = dblRound + 1
            End If
        End If
    End If

    Round = dblRound

End Function


Public Function Split(ByVal Expression As String, _
    Optional ByVal Delimiter As String = " ", _
    Optional ByVal Limit As Long = -1, _
    Optional ByVal Compare As VbCompareMethod _
                            = vbBinaryCompare) As Variant
    
    Dim lngPos      As Long
    Dim lngIndex    As Long
    Dim arrSplit()  As String
    
    If Right(Expression, Len(Delimiter)) <> Delimiter Then
        Expression = Expression & Delimiter
    End If
    
    Do
        lngPos = InStr(1, Expression, Delimiter, Compare)
        
        If lngPos = 0 Or Limit = lngIndex Then Exit Do
        
        ReDim Preserve arrSplit(lngIndex)
        arrSplit(lngIndex) = Left(Expression, lngPos - 1)
        
        Expression = Mid(Expression, lngPos + Len(Delimiter))
        lngIndex = lngIndex + 1
        lngPos = 0
    
    Loop
    
    Split = arrSplit()
    Erase arrSplit()

End Function


Public Function Join(ByRef SourceArray As Variant, _
    Optional ByVal Delimiter As String = " ") As String
    
    Dim lngLoop     As Long
    Dim strJoined   As String
    
    For lngLoop = LBound(SourceArray) To UBound(SourceArray)
        
        strJoined = strJoined & _
        SourceArray(lngLoop) & Delimiter
    
    Next
    
    Join = Left(strJoined, _
    Len(strJoined) - Len(Delimiter))

End Function


Public Function InStrRev(ByVal StringCheck As String, _
    ByVal StringMatch As String, Optional ByVal Start As Long = -1, _
    Optional Compare As VbCompareMethod = vbBinaryCompare) As Long
    
    Dim lngPos As Long
    Dim lngRev As Long
    
    
    If Start < 1 Then Start = Len(StringCheck)
    
    lngRev = Len(StringCheck)
    StringCheck = Left(StringCheck, Start)
    
    Do
        
        lngPos = InStr(lngRev, StringCheck, StringMatch, Compare)
        
        If lngPos > 0 Or lngRev = 0 Then Exit Do
        
        lngRev = lngRev - 1
    
    Loop
    
    InStrRev = lngPos
    
End Function


Public Function Replace(ByVal Expression As String, ByVal Find As String, _
    ByVal ReplaceWith As String, Optional ByVal Start As Long = 1, Optional ByVal Count As Long = -1, _
    Optional Compare As VbCompareMethod = vbBinaryCompare) As String

    Dim lngPos      As Long
    Dim lngCount    As Long
    
    Expression = Mid(Expression, Start)
    lngPos = InStr(1, Expression, Find)
   
    Do
        
        If lngPos = 0 Or Count = lngCount Then Exit Do
        
        Expression = Left(Expression, lngPos - 1) _
        & ReplaceWith & Mid(Expression, Len(Find) + lngPos)
        
        lngPos = InStr(lngPos, Expression, Find)
        lngCount = lngCount + 1
    
    Loop
    
    Replace = Expression

End Function


Public Function StrReverse(ByVal Expression As String) As String
    
    Dim strRev  As String
    Dim lngLoop As Long
    
    For lngLoop = Len(Expression) To 1 Step -1
        strRev = strRev & Mid(Expression, lngLoop, 1)
    Next
    
    StrReverse = strRev
    
End Function


