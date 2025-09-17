Attribute VB_Name = "MathMod"
Option Explicit
' MathMod - Centralized Mathematical Functions

'************************************************************
'**************** INTERPOLATION FUNCTIONS ******************
'************************************************************

' Standard linear interpolation between two points
Public Function LinearInterpolation(x As Variant, x1 As Variant, x2 As Variant, y1 As Variant, y2 As Variant) As Variant
    ' Handle exact matches first for performance
    If x = x1 Then
        LinearInterpolation = y1
        Exit Function
    End If
    
    If x = x2 Then
        LinearInterpolation = y2
        Exit Function
    End If
    
    ' Standard linear interpolation formula
    LinearInterpolation = y1 + (x - x1) * (y2 - y1) / (x2 - x1)
End Function

' High-performance linear interpolation for array data
Public Function FastLinearInterpolation(xPoint As Double, DataX As Variant, DataY As Variant) As Double
    Dim i As Long, lowerIdx As Long
    Dim value As Double: value = xPoint
    
    ' Find the bracketing points
    For i = LBound(DataX) To UBound(DataX) - 1
        If DataX(i) <= value And DataX(i + 1) >= value Then
            lowerIdx = i
            Exit For
        End If
    Next i
    
    ' Check for exact matches (within tolerance)
    If Abs(DataX(lowerIdx) - value) < 0.00001 Then
        FastLinearInterpolation = DataY(lowerIdx)
    ElseIf Abs(DataX(lowerIdx + 1) - value) < 0.00001 Then
        FastLinearInterpolation = DataY(lowerIdx + 1)
    Else
        ' Perform linear interpolation
        Dim xDelta As Double: xDelta = DataX(lowerIdx + 1) - DataX(lowerIdx)
        Dim yDelta As Double: yDelta = DataY(lowerIdx + 1) - DataY(lowerIdx)
        FastLinearInterpolation = DataY(lowerIdx) + (value - DataX(lowerIdx)) * (yDelta / xDelta)
    End If
End Function

' Legacy compatibility function (maps to new function name)
Public Function LinInterpolation(xPoint As Double, DataX As Variant, DataY As Variant, Optional ArrayBase As Integer = 1) As Double
    LinInterpolation = FastLinearInterpolation(xPoint, DataX, DataY)
End Function

' Binary search interpolation for large sorted datasets
Public Function BinaryInterpolation(xPoint As Double, DataX As Variant, DataY As Variant) As Double
    Dim low As Long, high As Long, mid As Long
    
    low = LBound(DataX)
    high = UBound(DataX)
    
    ' Handle edge cases
    If xPoint <= DataX(low) Then
        BinaryInterpolation = DataY(low)
        Exit Function
    End If
    
    If xPoint >= DataX(high) Then
        BinaryInterpolation = DataY(high)
        Exit Function
    End If
    
    ' Binary search for bracketing points
    Do While high - low > 1
        mid = (low + high) \ 2
        If DataX(mid) < xPoint Then
            low = mid
        Else
            high = mid
        End If
    Loop
    
    ' Linear interpolation between bracketing points
    Dim xDelta As Double: xDelta = DataX(high) - DataX(low)
    Dim yDelta As Double: yDelta = DataY(high) - DataY(low)
    BinaryInterpolation = DataY(low) + (xPoint - DataX(low)) * (yDelta / xDelta)
End Function

'************************************************************
'**************** STATISTICAL FUNCTIONS ********************
'************************************************************

' Calculate mean of array values
Public Function ArrayMean(arr As Variant) As Double
    Dim sum As Double
    Dim count As Long
    Dim i As Long
    
    For i = LBound(arr) To UBound(arr)
        If IsNumeric(arr(i)) Then
            sum = sum + arr(i)
            count = count + 1
        End If
    Next i
    
    If count > 0 Then
        ArrayMean = sum / count
    Else
        ArrayMean = 0
    End If
End Function

' Calculate standard deviation of array values
Public Function ArrayStdDev(arr As Variant, Optional usePopulation As Boolean = False) As Double
    Dim mean As Double
    Dim sumSquares As Double
    Dim count As Long
    Dim i As Long
    
    mean = ArrayMean(arr)
    
    For i = LBound(arr) To UBound(arr)
        If IsNumeric(arr(i)) Then
            sumSquares = sumSquares + (arr(i) - mean) ^ 2
            count = count + 1
        End If
    Next i
    
    If count > 1 Then
        If usePopulation Then
            ArrayStdDev = Sqr(sumSquares / count)
        Else
            ArrayStdDev = Sqr(sumSquares / (count - 1))
        End If
    Else
        ArrayStdDev = 0
    End If
End Function

' Find minimum value in array
Public Function ArrayMin(arr As Variant) As Double
    Dim minVal As Double
    Dim i As Long
    Dim firstFound As Boolean
    
    firstFound = False
    
    For i = LBound(arr) To UBound(arr)
        If IsNumeric(arr(i)) Then
            If Not firstFound Then
                minVal = arr(i)
                firstFound = True
            ElseIf arr(i) < minVal Then
                minVal = arr(i)
            End If
        End If
    Next i
    
    ArrayMin = minVal
End Function

' Find maximum value in array
Public Function ArrayMax(arr As Variant) As Double
    Dim maxVal As Double
    Dim i As Long
    Dim firstFound As Boolean
    
    firstFound = False
    
    For i = LBound(arr) To UBound(arr)
        If IsNumeric(arr(i)) Then
            If Not firstFound Then
                maxVal = arr(i)
                firstFound = True
            ElseIf arr(i) > maxVal Then
                maxVal = arr(i)
            End If
        End If
    Next i
    
    ArrayMax = maxVal
End Function

'************************************************************
'**************** GENERIC ARRAY MATH ***********************
'************************************************************

' Calculate cumulative sum of array values (generic utility)
Public Function CumulativeSum(values As Variant, Optional reverseOrder As Boolean = False) As Variant
    Dim result As Variant
    Dim i As Long
    Dim runningTotal As Double
    
    ReDim result(LBound(values) To UBound(values))
    
    If reverseOrder Then
        ' Sum from end to beginning
        runningTotal = 0
        For i = UBound(values) To LBound(values) Step -1
            If IsNumeric(values(i)) Then runningTotal = runningTotal + values(i)
            result(i) = runningTotal
        Next i
    Else
        ' Sum from beginning to end
        runningTotal = 0
        For i = LBound(values) To UBound(values)
            If IsNumeric(values(i)) Then runningTotal = runningTotal + values(i)
            result(i) = runningTotal
        Next i
    End If
    
    CumulativeSum = result
End Function

' Calculate differences between adjacent array elements
Public Function ArrayDifferences(values As Variant, Optional absolute As Boolean = False) As Variant
    Dim result As Variant
    Dim i As Long
    Dim diff As Double
    
    If UBound(values) <= LBound(values) Then
        ArrayDifferences = Empty
        Exit Function
    End If
    
    ReDim result(LBound(values) + 1 To UBound(values))
    
    For i = LBound(values) + 1 To UBound(values)
        If IsNumeric(values(i)) And IsNumeric(values(i - 1)) Then
            diff = values(i) - values(i - 1)
            If absolute Then diff = Abs(diff)
            result(i) = diff
        Else
            result(i) = 0
        End If
    Next i
    
    ArrayDifferences = result
End Function

'************************************************************
'**************** ARRAY UTILITY FUNCTIONS ******************
'************************************************************

' Smooth array data using moving average
Public Function MovingAverage(arr As Variant, windowSize As Long) As Variant
    Dim result As Variant
    Dim i As Long, j As Long
    Dim sum As Double
    Dim count As Long
    
    ReDim result(LBound(arr) To UBound(arr))
    
    For i = LBound(arr) To UBound(arr)
        sum = 0
        count = 0
        
        ' Calculate window bounds
        Dim startIdx As Long, endIdx As Long
        startIdx = Application.Max(LBound(arr), i - windowSize \ 2)
        endIdx = Application.Min(UBound(arr), i + windowSize \ 2)
        
        ' Sum values in window
        For j = startIdx To endIdx
            If IsNumeric(arr(j)) Then
                sum = sum + arr(j)
                count = count + 1
            End If
        Next j
        
        If count > 0 Then
            result(i) = sum / count
        Else
            result(i) = arr(i)
        End If
    Next i
    
    MovingAverage = result
End Function

' Scale array values by a factor
Public Function ScaleArray(arr As Variant, scaleFactor As Double) As Variant
    Dim result As Variant
    Dim i As Long
    
    ReDim result(LBound(arr) To UBound(arr))
    
    For i = LBound(arr) To UBound(arr)
        If IsNumeric(arr(i)) Then
            result(i) = arr(i) * scaleFactor
        Else
            result(i) = arr(i)
        End If
    Next i
    
    ScaleArray = result
End Function

' Add offset to array values
Public Function OffsetArray(arr As Variant, offset As Double) As Variant
    Dim result As Variant
    Dim i As Long
    
    ReDim result(LBound(arr) To UBound(arr))
    
    For i = LBound(arr) To UBound(arr)
        If IsNumeric(arr(i)) Then
            result(i) = arr(i) + offset
        Else
            result(i) = arr(i)
        End If
    Next i
    
    OffsetArray = result
End Function

'************************************************************
'**************** CONVERSION FUNCTIONS **********************
'************************************************************

' Convert string to boolean (for flag processing)
Public Function ConvertToBool(str As Variant) As Boolean
    Dim temp As String
    
    If IsEmpty(str) Or IsNull(str) Then
        ConvertToBool = False
        Exit Function
    End If
    
    temp = Replace(Replace(CStr(str), "#", ""), " ", "")
    temp = UCase(temp)
    
    Select Case temp
        Case "TRUE", "YES", "1", "ON", "ENABLED"
            ConvertToBool = True
        Case Else
            ConvertToBool = False
    End Select
End Function

' Convert time decimal to elapsed time units
Public Function DecimalTimeToElapsed(timeValue As Double, startTime As Double, Optional outputUnit As String = "minutes") As Double
    Dim currentTime As Double
    currentTime = timeValue
    
    ' Handle midnight rollover
    If currentTime < startTime Then
        currentTime = currentTime + 1
    End If
    
    ' Convert to elapsed time in specified units
    Dim elapsedDays As Double
    elapsedDays = currentTime - startTime
    
    Select Case LCase(outputUnit)
        Case "days"
            DecimalTimeToElapsed = elapsedDays
        Case "hours"
            DecimalTimeToElapsed = elapsedDays * 24
        Case "minutes"
            DecimalTimeToElapsed = elapsedDays * 1440
        Case "seconds"
            DecimalTimeToElapsed = elapsedDays * 86400
        Case Else
            DecimalTimeToElapsed = elapsedDays * 1440 ' Default to minutes
    End Select
End Function

'************************************************************
'**************** VALIDATION FUNCTIONS **********************
'************************************************************

' Check if array contains only numeric values
Public Function IsNumericArray(arr As Variant) As Boolean
    Dim i As Long
    
    IsNumericArray = True
    
    For i = LBound(arr) To UBound(arr)
        If Not IsNumeric(arr(i)) Then
            IsNumericArray = False
            Exit Function
        End If
    Next i
End Function

' Check if value is within tolerance
Public Function WithinTolerance(value1 As Double, value2 As Double, tolerance As Double) As Boolean
    WithinTolerance = (Abs(value1 - value2) <= tolerance)
End Function

' Round to specified decimal places
Public Function RoundTo(value As Double, decimalPlaces As Long) As Double
    Dim factor As Double
    factor = 10 ^ decimalPlaces
    RoundTo = Int(value * factor + 0.5) / factor
End Function

'************************************************************
'**************** PERFORMANCE FUNCTIONS ********************
'************************************************************

' High-performance array summation
Public Function FastArraySum(arr As Variant) As Double
    Dim sum As Double
    Dim i As Long
    
    For i = LBound(arr) To UBound(arr)
        If IsNumeric(arr(i)) Then
            sum = sum + arr(i)
        End If
    Next i
    
    FastArraySum = sum
End Function

' Find index of maximum value
Public Function ArrayMaxIndex(arr As Variant) As Long
    Dim maxVal As Double
    Dim maxIdx As Long
    Dim i As Long
    Dim firstFound As Boolean
    
    firstFound = False
    
    For i = LBound(arr) To UBound(arr)
        If IsNumeric(arr(i)) Then
            If Not firstFound Then
                maxVal = arr(i)
                maxIdx = i
                firstFound = True
            ElseIf arr(i) > maxVal Then
                maxVal = arr(i)
                maxIdx = i
            End If
        End If
    Next i
    
    ArrayMaxIndex = maxIdx
End Function

' Find index of minimum value
Public Function ArrayMinIndex(arr As Variant) As Long
    Dim minVal As Double
    Dim minIdx As Long
    Dim i As Long
    Dim firstFound As Boolean
    
    firstFound = False
    
    For i = LBound(arr) To UBound(arr)
        If IsNumeric(arr(i)) Then
            If Not firstFound Then
                minVal = arr(i)
                minIdx = i
                firstFound = True
            ElseIf arr(i) < minVal Then
                minVal = arr(i)
                minIdx = i
            End If
        End If
    Next i
    
    ArrayMinIndex = minIdx
End Function
