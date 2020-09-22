Attribute VB_Name = "Module4"
Option Explicit

Public Sub QuickSort(ByRef vntArr As Variant, Optional ByVal lngLeft As Long = -2, Optional ByVal lngRight As Long = -2)
Dim i As Long
Dim j As Long
Dim lngMid As Long
Dim vntTestVal As Variant

If lngLeft = -2 Then lngLeft = LBound(vntArr)
If lngRight = -2 Then lngRight = UBound(vntArr)
If lngLeft < lngRight Then
lngMid = (lngLeft + lngRight) \ 2
vntTestVal = vntArr(lngMid)
i = lngLeft
j = lngRight
Do
Do While vntArr(i) < vntTestVal
i = i + 1
Loop
Do While vntArr(j) > vntTestVal
j = j - 1
Loop
    If i <= j Then
    Call SwapElements(vntArr, i, j)
    i = i + 1
    j = j - 1
    End If
Loop Until i > j

If j <= lngMid Then
Call QuickSort(vntArr, lngLeft, j)
Call QuickSort(vntArr, i, lngRight)
Else
Call QuickSort(vntArr, i, lngRight)
Call QuickSort(vntArr, lngLeft, j)
End If
    End If
End Sub

Private Sub SwapElements(ByRef vntItems As Variant, ByVal lngItem1 As Long, ByVal lngItem2 As Long)
Dim vntTemp As Variant

vntTemp = vntItems(lngItem2)
vntItems(lngItem2) = vntItems(lngItem1)
vntItems(lngItem1) = vntTemp
End Sub
