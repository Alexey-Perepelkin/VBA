VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayList1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const INIT_SIZE = 16
Private Const CUT_RATE = 4
Private arr(INIT_SIZE) As Variant
Private pointer As Integer

'    ƒобавл€ет новый элемент в список. ѕри достижении размера внутреннего
'    массива происходит его увеличение в два раза.
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
  
Private Sub resise(ByVal newLenght As Integer)
    Dim newArr(newLenght) As Variant
    Dim i As Integer
    For i = 0 To pointer
        newArr(i) = arr(i)
    Next i
    ReDim arr(newLenght)
    arr = newArr
End Sub
Public Function size() As Integer
    size = pointer
End Function

Public Function getItem(ByVal index As Integer) As Variant
If index > pointer Then
    MsgBox "Ёлемент с индексом " & index & " не существует", vbExclamation
Else
    getItem = arr(index)
End If
End Function



