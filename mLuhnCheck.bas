Attribute VB_Name = "mLuhnCheck"
Option Explicit
' Luhn check module
' From https://www.tek-tips.com/faqs.cfm?fid=6704


' this array holds lookup values for translating digits
'(when necessary)
Dim xL(9) As Integer

Public Sub fillXL()

'this sub is used to fill the public array used for
'translation of digits

'array values in xL are for the index integer
'0 * 2 = 0 --> 0 = 0
'6*2 = 12 --> 1 + 2 = 3
'having this array available saves us from performing string
'conversions and math operations (just lookup by index)
xL(0) = 0
xL(1) = 2
xL(2) = 4
xL(3) = 6
xL(4) = 8
xL(5) = 1
xL(6) = 3
xL(7) = 5
xL(8) = 7
xL(9) = 9

End Sub

Public Function LuhnCheck(ByVal intStr As String) As String

'this function is used to return the check digit to be
'appended to a given number
Dim b() As Byte
Dim x As Integer
Dim sD As Integer
' sD holds sum of digits (as modified by Luhn algorithm)
Dim lD As Integer
' lD is used to store checksum digit (10 - sD Mod 10)

'check for numeric input
If Not IsNumeric(intStr & ".0e0") Then
   LuhnCheck = "X"
   Exit Function
End If

Call fillXL

sD = 0
ReDim b(Len(intStr))

b = StrConv(StrReverse(intStr), vbFromUnicode)

'b(x) - 48 == faster way to get integer value
'from unicode byte value
'first digit (starting from right)is doubled/digits added
'because once check digit is appended this will be the second
For x = LBound(b) To UBound(b)
    If x Mod 2 = 0 Then
        sD = sD + xL(b(x) - 48)
    Else
        sD = sD + (b(x) - 48)
    End If
Next

lD = 10 - (sD Mod 10)

'we don't want to add 10, if lD calculates to 10 then we
'really want to add 0
If lD = 10 Then
    lD = 0
End If

'return string with check digit appended
LuhnCheck = CStr(lD)

End Function

Public Function luhnValid(ByVal intStr As String) As Boolean

'this function is used to check if a number entered is valid
Dim sD As Integer
Dim bl As Boolean
Dim b() As Byte
Dim x As Integer

'check for numeric input
If Not IsNumeric(intStr & ".0e0") Then
   luhnValid = False
   Exit Function
End If

Call fillXL

ReDim b(Len(intStr))

b = StrConv(intStr, vbFromUnicode)

bl = False
sD = 0

'start with last digit, work towards first
For x = UBound(b) To LBound(b) Step -1
    If bl Then
        sD = sD + xL(b(x) - 48)
    Else
        sD = sD + b(x) - 48
    End If
    
    bl = Not (bl)
Next

luhnValid = (sD Mod 10 = 0)

End Function
