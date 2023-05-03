Attribute VB_Name = "Module1"
Sub lotto()

Dim num As Long
Dim fPlace As Long
Dim sPlace As Long
Dim tPlace As Long
Dim b1 As Long
Dim b2 As Long
Dim b3 As Long

Dim position As Integer

fPlace = 3957481
sPlace = 5865187
tPlace = 2817729

b1 = 2275339
b2 = 5868182
b3 = 1841402

Dim flag As Boolean
flag = False

For i = 2 To 1001
    num = Cells(i, 3)

    If num = fPlace Then
    Cells(2, 6) = Cells(i, 1)
    Cells(2, 7) = Cells(i, 2)
    Cells(2, 8) = fPlace
    
    ElseIf num = sPlace Then
    Cells(3, 6) = Cells(i, 1)
    Cells(3, 7) = Cells(i, 2)
    Cells(3, 8) = sPlace
    
    ElseIf num = tPlace Then
    Cells(4, 6) = Cells(i, 1)
    Cells(4, 7) = Cells(i, 2)
    Cells(4, 8) = tPlace
    
    ElseIf (num = b1 And flag = False) Then
    i = position
    flag = True
    Cells(5, 6) = Cells(i, 1)
    Cells(5, 7) = Cells(i, 2)
    Cells(5, 8) = b1
    
    
    ElseIf (num = b2 And flag = False) Then
    i = position
    flag = True
    Cells(5, 6) = Cells(i, 1)
    Cells(5, 7) = Cells(i, 2)
    Cells(5, 8) = b2
    
    ElseIf (num = b3 And flag = False) Then
    flag = True
    Cells(5, 6) = Cells(i, 1)
    Cells(5, 7) = Cells(i, 2)
    Cells(5, 8) = b3
    
    End If
    

Next i

    
    
End Sub



