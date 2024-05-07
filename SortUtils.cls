VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SortUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub directSort(ByVal arr As Variant)

Dim i As Integer
Dim j As Integer
Dim minIndex As Integer
Dim buf As Integer

For i = 0 To UBound(arr)
    minIndex = i
    For j = i + 1 To UBound(arr)
        If arr(j) < arr(minIndex) Then
            minIndex = j
        End If
    Next j
    If i <> minIndex Then
        buf = arr(i)
        arr(i) = arr(minIndex)
        arr(minIndex) = buf
    End If
Next i

End Sub

Public Sub quickSort(ByVal arr As Variant)
If arr = Null Then
    End Sub
End If
Call quickSort1(arr, 0, UBound(arr))

End Sub

'    public static void quickSort(int[] array){
'        if (array == null)
'            return;
'        quickSort(array, 0, array.length - 1);
'    }

Private Sub quickSort1(ByVal arr As Variant, ByVal theStart As Integer, ByVal theEnd As Integer)
 
 Dim theLeft As Integer
 Dim theRight As Integer
 Dim middle As Integer
 Dim buf As Integer

 theLeft = theStart
 theRight = theEnd
 middle = arr(theStart + theEnd / 2)
  Do While arr(theLeft) < middle
    theLeft = theLeft + 1
  Loop
  Do While arr(theRight) > middle
    theRight = theRight + 1
  Loop
                If theLeft <= theRight Then
                    If theLeft < theRight Then
                        buf = arr(theLeft)
                        arr(theLeft) = arr(theRight)
                        arr(theRight) = buf
                    End If

                    theLeft = theLeft + 1
                    theRight = theRight - 1
                End If
    Do While theLeft <= theRight

        If theLeft < theEnd Then
           Call quickSort1(arr, theLeft, theEnd)
        End If

        If theStart < theRight Then
            Call quickSort1(arr, theStart, theRight)
        End If
End Sub

'
'    private static void quickSort(int[] array, int start, int end){
'        int left = start;
'        int right = end;
'
'        int middle = array[(start + end) / 2];
'
'        do {
    '            while (array[left] < middle){
    '                left++;
    '            }
'
    '            while (array[right] > middle){
    '                right--;
    '            }
'
    '            if (left <= right){
    '                if (left < right){
    '                    int buf = array[left];
    '                    array[left] = array[right];
    '                    array[right] = buf;
    '                }
    '                left++;
    '                right--;
    '            }
'           }
'        while (left <= right);
'
'        if (left < end){
'            quickSort(array, left, end);
'        }
'        if (start < right){
'            quickSort(array, start, right);
'        }
'    }
