VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const INIT_SIZE = 16
Private Const CUT_RATE = 4
Private arr() As Variant
Private pointer As Integer
Private myArr As Variant
Private Sub Class_Initialize()
    ReDim arr(INIT_SIZE) As Variant
End Sub

' Adding  a new element to the list
' When the array size is reached,it doubles in size
 Public Sub add(ByVal item As Variant)
'    pointer = 0
    If pointer = UBound(arr) Then
        Call resise(UBound(arr) * 2)
    End If
    arr(pointer) = item
    pointer = pointer + 1
 End Sub
 Public Sub remove(ByVal index As Integer)
    Dim i As Integer
    For i = index To pointer
        arr(i) = arr(i + 1)
    Next i
    arr(pointer) = Null
    pointer = pointer - 1
    If UBound(arr) > INIT_SIZE And pointer < UBound(arr) / CUT_RATE Then
        Call resise(UBound(arr) / 2)
    End If
 End Sub
  
Private Sub resise(ByVal newLenght1 As Integer)
    Dim newArr() As Variant
    ReDim newArr(newLenght1)
    Dim i As Integer
    For i = 0 To pointer
        newArr(i) = arr(i)
    Next i
    ReDim arr(newLenght1)
    arr = newArr
End Sub
Private Function size() As Integer
    size = pointer
End Function

Public Function getItem(ByVal index As Integer) As Variant
If index > pointer Then
    MsgBox "Element with index " & index & " does not exist!", vbExclamation
Else
    getItem = arr(index)
End If
End Function




Public Property Get GetArray() As Variant
    ReDim myArr(size)
    Dim i As Integer
    For i = 0 To size - 1
        myArr(i) = arr(i)
    Next i
    GetArray = myArr
End Property


Private Sub quickSort(ByRef arr As Variant)

    Call quickSort1(arr, 0, UBound(arr) - 1)

End Sub



Private Sub quickSort1(ByRef arr As Variant, ByVal theStart As Integer, ByVal theEnd As Integer)
 
    Dim theLeft As Integer: theLeft = theStart
    Dim theRight As Integer: theRight = theEnd
    Dim middle As String: middle = arr((theStart + theEnd) / 2)
    Dim buf As String
 
    Do
        Do While arr(theLeft) < middle
            theLeft = theLeft + 1
        Loop
        
        Do While arr(theRight) > middle
            theRight = theRight - 1
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
        
    Loop While theLeft <= theRight

        If theLeft < theEnd Then
           Call quickSort1(arr, theLeft, theEnd)
        End If

        If theStart < theRight Then
            Call quickSort1(arr, theStart, theRight)
        End If
End Sub


