Attribute VB_Name = "Module2"
Sub sentenceBreaker():

Dim Words() As String
Dim theSentence As String
Dim wordNum1 As Integer
Dim wordNum2 As Integer
Dim wordNum3 As Integer

Range("F6").Value = "Input a sentence:"
theSentence = Range("G6").Value

Range("F7").Value = "Word Number"

wordNum1 = Range("F8").Value
wordNum2 = Range("F9").Value
wordNum3 = Range("F10").Value

Words = Split(theSentence, " ")
Range("G8").Value = (Words(wordNum1 - 1))
Range("G9").Value = (Words(wordNum2 - 1))
Range("G10").Value = (Words(wordNum3 - 1))

End Sub
