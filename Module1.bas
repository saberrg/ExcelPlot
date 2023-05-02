Attribute VB_Name = "Module1"
Sub myClass():

 Range("F1").Value = "Price"
 Range("G1").Value = "Tax"
 Range("H1").Value = "Quantity"
 Range("I1").Value = "Total"



Dim Price As Double
Dim Tax As Double
Dim Quantity As Double
Dim Total As Double


Price = Range("F2").Value
Tax = Range("G2").Value
Quantity = Range("H2").Value

Total = Price * (1 + Tax) * Quantity

Range("I2").Value = Total

End Sub

