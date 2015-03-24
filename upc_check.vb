'by Andy Bertagnoli
'a function to add the check digit to 11 digit upc's
Function upc_chk(upc As Double) As Double

Dim upc_s As String
upc_s = upc

Dim u1, u2, u3, u4, u5, u6, u7, u8, u9, u10, u11, u12 As Integer
Dim upc_new_s As String
Dim upc_twlv As Double
Dim a, b, c, d As Integer


u1 = Left(upc, 1)
u2 = Right(Left(upc, 2), 1)
u3 = Right(Left(upc, 3), 1)
u4 = Right(Left(upc, 4), 1)
u5 = Right(Left(upc, 5), 1)
u6 = Right(Left(upc, 6), 1)
u7 = Right(Left(upc, 7), 1)
u8 = Right(Left(upc, 8), 1)
u9 = Right(Left(upc, 9), 1)
u10 = Right(Left(upc, 10), 1)
u11 = Right(Left(upc, 11), 1)
u12 = ((((CDbl(u1) + CDbl(u3) + CDbl(u5) + CDbl(u7) + CDbl(u9) + CDbl(u11)) * 3) + (CDbl(u2) + CDbl(u4) + CDbl(u6) + CDbl(u8) + CDbl(u10))) Mod 10)
If u12 > 0 Then
u12 = 10 - u12
End If


'a = ((CDbl(u1) + CDbl(u3) + CDbl(u5) + CDbl(u7) + CDbl(u9) + CDbl(u11)) * 3)
'Debug.Print a & "--a"
'b = (CDbl(u2) + CDbl(u4) + CDbl(u6) + CDbl(u8) + CDbl(u10))
'Debug.Print b & "--b"
'c = (a + b) Mod 10
'Debug.Print c & "--c"

upc_chk = CDbl(u1 & u2 & u3 & u4 & u5 & u6 & u7 & u8 & u9 & u10 & u11 & u12)

End Function
