Attribute VB_Name = "ExcelProgress"
Option Explicit

Enum xlProgressAction
    xlInitMeter
    xlUpdateMeter
    xlRemoveMeter
End Enum

Dim progressCount As Integer, progressIndex As Integer

Public Sub ProgressBar(mode As xlProgressAction, Optional ByVal message As String, Optional ByVal inputNumber)

    Select Case mode
    
        Case Is = xlInitMeter
            progressCount = inputNumber
            progressIndex = 1
        Case Is = xlUpdateMeter
            progressIndex = inputNumber
            Application.StatusBar = IIf(message <> "", message, "") & ProgressString(progressIndex, progressCount)
        Case Is = xlRemoveMeter
            Application.StatusBar = False
    
    End Select
    
End Sub
'This animates the partially completed block
Function BarProgress(ByVal val, maxVal)

    BarProgress = Int(Round((1 / 6) * Floor(CDbl(Normalize(val, maxVal, 0) * 10 - Floor(Normalize(val, maxVal, 0) * 10)) / (1 / 6)), 1) / (1 / 6))

End Function

Function ProgressString(index, count) As String

    If index <> count Then

        Dim progressPercent
        progressPercent = index / count
        ProgressString = String(Int(progressPercent * 10), ChrW(9609)) _
                        & Array(ChrW(8198) & ChrW(8198) & ChrW(8198), ChrW(9615), ChrW(9614), ChrW(9613), ChrW(9611), ChrW(9610))(BarProgress(index, count)) _
                        & String((9 - Int((progressPercent * 10))) * 3, ChrW(8198)) & ChrW(9615)
     End If
    
End Function

Function Min(i1, i2)

    Min = i1
    If i2 < i1 Then Min = i2

End Function

Function Max(i1, i2)

    Max = i1
    If i2 > i1 Then Max = i2

End Function

Function Normalize(i, maxValue, minValue)
    
    Normalize = (Min(Max(i, minValue), maxValue) - minValue) / CDbl(maxValue - minValue)

End Function

Function Floor(i)

    Floor = Int(i) - 1 * (Int(i) > i)

End Function
