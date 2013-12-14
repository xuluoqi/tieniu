<%@Language="VBSCRIPT"%>
<%
db="hyx_dd.mdb"
Set conn = Server.CreateObject("ADODB.Connection")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.Open connstr
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from yuangong where username='"&session("userid")&"' and password='"&session("password")&"'",conn,1,1
if rs.eof and rs.bof then
response.write "<SCRIPT language=JavaScript>alert('对不起，你没有权限进入该页面!');"
response.write"this.location.href='index.asp';</SCRIPT>"
On Error GoTo 0
Err.Raise 9999
end if
rs.close
set rs=nothing
if  session("level")="" then
response.write "<SCRIPT language=JavaScript>alert('对不起，你没有权限进入该页面!');"
response.write"this.location.href='index.asp';</SCRIPT>"
On Error GoTo 0
Err.Raise 9999
end if
session.Timeout=60
Private Const BITS_TO_A_BYTE = 8
Private Const BYTES_TO_A_WORD = 4
Private Const BITS_TO_A_WORD = 32
Private m_lOnBits(30)
Private m_l2Power(30)
Private Function LShift(lValue, iShiftBits)
If iShiftBits = 0 Then
LShift = lValue
Exit Function
ElseIf iShiftBits = 31 Then
If lValue And 1 Then
LShift = &H80000000
Else
LShift = 0
End If
Exit Function
ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
Err.Raise 6
End If
If (lValue And m_l2Power(31 - iShiftBits)) Then
LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000
Else
LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits))
End If
End Function
Private Function RShift(lValue, iShiftBits)
If iShiftBits = 0 Then
RShift = lValue
Exit Function
ElseIf iShiftBits = 31 Then
If lValue And &H80000000 Then
RShift = 1
Else
RShift = 0
End If
Exit Function
ElseIf iShiftBits < 0 Or iShiftBits > 31 Then
Err.Raise 6
End If
RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits)
If (lValue And &H80000000) Then
RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1)))
End If
End Function
Private Function RotateLeft(lValue, iShiftBits)
RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits))
End Function
Private Function AddUnsigned(lX, lY)
Dim lX4
Dim lY4
Dim lX8
Dim lY8
Dim lResult
lX8 = lX And &H80000000
lY8 = lY And &H80000000
lX4 = lX And &H40000000
lY4 = lY And &H40000000
lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)
If lX4 And lY4 Then
lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
ElseIf lX4 Or lY4 Then
If lResult And &H40000000 Then
lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
Else
lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
End If
Else
lResult = lResult Xor lX8 Xor lY8
End If
AddUnsigned = lResult
End Function
Private Function md5_F(x, y, z)
md5_F = (x And y) Or ((Not x) And z)
End Function
Private Function md5_G(x, y, z)
md5_G = (x And z) Or (y And (Not z))
End Function
Private Function md5_H(x, y, z)
md5_H = (x Xor y Xor z)
End Function
Private Function md5_I(x, y, z)
md5_I = (y Xor (x Or (Not z)))
End Function
Private Sub md5_FF(a, b, c, d, x, s, ac)
a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_F(b, c, d), x), ac))
a = RotateLeft(a, s)
a = AddUnsigned(a, b)
End Sub
Private Sub md5_GG(a, b, c, d, x, s, ac)
a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_G(b, c, d), x), ac))
a = RotateLeft(a, s)
a = AddUnsigned(a, b)
End Sub
Private Sub md5_HH(a, b, c, d, x, s, ac)
a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_H(b, c, d), x), ac))
a = RotateLeft(a, s)
a = AddUnsigned(a, b)
End Sub
Private Sub md5_II(a, b, c, d, x, s, ac)
a = AddUnsigned(a, AddUnsigned(AddUnsigned(md5_I(b, c, d), x), ac))
a = RotateLeft(a, s)
a = AddUnsigned(a, b)
End Sub
Private Function ConvertToWordArray(sMessage)
Dim lMessageLength
Dim lNumberOfWords
Dim lWordArray()
Dim lBytePosition
Dim lByteCount
Dim lWordCount
Const MODULUS_BITS = 512
Const CONGRUENT_BITS = 448
lMessageLength = Len(sMessage)
lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD)
ReDim lWordArray(lNumberOfWords - 1)
lBytePosition = 0
lByteCount = 0
Do Until lByteCount >= lMessageLength
lWordCount = lByteCount \ BYTES_TO_A_WORD
lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(Asc(Mid(sMessage, lByteCount + 1, 1)), lBytePosition)
lByteCount = lByteCount + 1
Loop
lWordCount = lByteCount \ BYTES_TO_A_WORD
lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE
lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition)
lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3)
lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29)
ConvertToWordArray = lWordArray
End Function
Private Function WordToHex(lValue)
Dim lByte
Dim lCount
For lCount = 0 To 3
lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) And m_lOnBits(BITS_TO_A_BYTE - 1)
WordToHex = WordToHex & Right("0" & Hex(lByte), 2)
Next
End Function
Public Function MD5(sMessage)
m_lOnBits(0) = CLng(1)
m_lOnBits(1) = CLng(3)
m_lOnBits(2) = CLng(7)
m_lOnBits(3) = CLng(15)
m_lOnBits(4) = CLng(31)
m_lOnBits(5) = CLng(63)
m_lOnBits(6) = CLng(127)
m_lOnBits(7) = CLng(255)
m_lOnBits(8) = CLng(511)
m_lOnBits(9) = CLng(1023)
m_lOnBits(10) = CLng(2047)
m_lOnBits(11) = CLng(4095)
m_lOnBits(12) = CLng(8191)
m_lOnBits(13) = CLng(16383)
m_lOnBits(14) = CLng(32767)
m_lOnBits(15) = CLng(65535)
m_lOnBits(16) = CLng(131071)
m_lOnBits(17) = CLng(262143)
m_lOnBits(18) = CLng(524287)
m_lOnBits(19) = CLng(1048575)
m_lOnBits(20) = CLng(2097151)
m_lOnBits(21) = CLng(4194303)
m_lOnBits(22) = CLng(8388607)
m_lOnBits(23) = CLng(16777215)
m_lOnBits(24) = CLng(33554431)
m_lOnBits(25) = CLng(67108863)
m_lOnBits(26) = CLng(134217727)
m_lOnBits(27) = CLng(268435455)
m_lOnBits(28) = CLng(536870911)
m_lOnBits(29) = CLng(1073741823)
m_lOnBits(30) = CLng(2147483647)
m_l2Power(0) = CLng(1)
m_l2Power(1) = CLng(2)
m_l2Power(2) = CLng(4)
m_l2Power(3) = CLng(8)
m_l2Power(4) = CLng(16)
m_l2Power(5) = CLng(32)
m_l2Power(6) = CLng(64)
m_l2Power(7) = CLng(128)
m_l2Power(8) = CLng(256)
m_l2Power(9) = CLng(512)
m_l2Power(10) = CLng(1024)
m_l2Power(11) = CLng(2048)
m_l2Power(12) = CLng(4096)
m_l2Power(13) = CLng(8192)
m_l2Power(14) = CLng(16384)
m_l2Power(15) = CLng(32768)
m_l2Power(16) = CLng(65536)
m_l2Power(17) = CLng(131072)
m_l2Power(18) = CLng(262144)
m_l2Power(19) = CLng(524288)
m_l2Power(20) = CLng(1048576)
m_l2Power(21) = CLng(2097152)
m_l2Power(22) = CLng(4194304)
m_l2Power(23) = CLng(8388608)
m_l2Power(24) = CLng(16777216)
m_l2Power(25) = CLng(33554432)
m_l2Power(26) = CLng(67108864)
m_l2Power(27) = CLng(134217728)
m_l2Power(28) = CLng(268435456)
m_l2Power(29) = CLng(536870912)
m_l2Power(30) = CLng(1073741824)
Dim x
Dim k
Dim AA
Dim BB
Dim CC
Dim DD
Dim a
Dim b
Dim c
Dim d
Const S11 = 7
Const S12 = 12
Const S13 = 17
Const S14 = 22
Const S21 = 5
Const S22 = 9
Const S23 = 14
Const S24 = 20
Const S31 = 4
Const S32 = 11
Const S33 = 16
Const S34 = 23
Const S41 = 6
Const S42 = 10
Const S43 = 15
Const S44 = 21
x = ConvertToWordArray(sMessage)
a = &H67452301
b = &HEFCDAB89
c = &H98BADCFE
d = &H10325476
For k = 0 To UBound(x) Step 16
AA = a
BB = b
CC = c
DD = d
md5_FF a, b, c, d, x(k + 0), S11, &HD76AA478
md5_FF d, a, b, c, x(k + 1), S12, &HE8C7B756
md5_FF c, d, a, b, x(k + 2), S13, &H242070DB
md5_FF b, c, d, a, x(k + 3), S14, &HC1BDCEEE
md5_FF a, b, c, d, x(k + 4), S11, &HF57C0FAF
md5_FF d, a, b, c, x(k + 5), S12, &H4787C62A
md5_FF c, d, a, b, x(k + 6), S13, &HA8304613
md5_FF b, c, d, a, x(k + 7), S14, &HFD469501
md5_FF a, b, c, d, x(k + 8), S11, &H698098D8
md5_FF d, a, b, c, x(k + 9), S12, &H8B44F7AF
md5_FF c, d, a, b, x(k + 10), S13, &HFFFF5BB1
md5_FF b, c, d, a, x(k + 11), S14, &H895CD7BE
md5_FF a, b, c, d, x(k + 12), S11, &H6B901122
md5_FF d, a, b, c, x(k + 13), S12, &HFD987193
md5_FF c, d, a, b, x(k + 14), S13, &HA679438E
md5_FF b, c, d, a, x(k + 15), S14, &H49B40821
md5_GG a, b, c, d, x(k + 1), S21, &HF61E2562
md5_GG d, a, b, c, x(k + 6), S22, &HC040B340
md5_GG c, d, a, b, x(k + 11), S23, &H265E5A51
md5_GG b, c, d, a, x(k + 0), S24, &HE9B6C7AA
md5_GG a, b, c, d, x(k + 5), S21, &HD62F105D
md5_GG d, a, b, c, x(k + 10), S22, &H2441453
md5_GG c, d, a, b, x(k + 15), S23, &HD8A1E681
md5_GG b, c, d, a, x(k + 4), S24, &HE7D3FBC8
md5_GG a, b, c, d, x(k + 9), S21, &H21E1CDE6
md5_GG d, a, b, c, x(k + 14), S22, &HC33707D6
md5_GG c, d, a, b, x(k + 3), S23, &HF4D50D87
md5_GG b, c, d, a, x(k + 8), S24, &H455A14ED
md5_GG a, b, c, d, x(k + 13), S21, &HA9E3E905
md5_GG d, a, b, c, x(k + 2), S22, &HFCEFA3F8
md5_GG c, d, a, b, x(k + 7), S23, &H676F02D9
md5_GG b, c, d, a, x(k + 12), S24, &H8D2A4C8A
md5_HH a, b, c, d, x(k + 5), S31, &HFFFA3942
md5_HH d, a, b, c, x(k + 8), S32, &H8771F681
md5_HH c, d, a, b, x(k + 11), S33, &H6D9D6122
md5_HH b, c, d, a, x(k + 14), S34, &HFDE5380C
md5_HH a, b, c, d, x(k + 1), S31, &HA4BEEA44
md5_HH d, a, b, c, x(k + 4), S32, &H4BDECFA9
md5_HH c, d, a, b, x(k + 7), S33, &HF6BB4B60
md5_HH b, c, d, a, x(k + 10), S34, &HBEBFBC70
md5_HH a, b, c, d, x(k + 13), S31, &H289B7EC6
md5_HH d, a, b, c, x(k + 0), S32, &HEAA127FA
md5_HH c, d, a, b, x(k + 3), S33, &HD4EF3085
md5_HH b, c, d, a, x(k + 6), S34, &H4881D05
md5_HH a, b, c, d, x(k + 9), S31, &HD9D4D039
md5_HH d, a, b, c, x(k + 12), S32, &HE6DB99E5
md5_HH c, d, a, b, x(k + 15), S33, &H1FA27CF8
md5_HH b, c, d, a, x(k + 2), S34, &HC4AC5665
md5_II a, b, c, d, x(k + 0), S41, &HF4292244
md5_II d, a, b, c, x(k + 7), S42, &H432AFF97
md5_II c, d, a, b, x(k + 14), S43, &HAB9423A7
md5_II b, c, d, a, x(k + 5), S44, &HFC93A039
md5_II a, b, c, d, x(k + 12), S41, &H655B59C3
md5_II d, a, b, c, x(k + 3), S42, &H8F0CCC92
md5_II c, d, a, b, x(k + 10), S43, &HFFEFF47D
md5_II b, c, d, a, x(k + 1), S44, &H85845DD1
md5_II a, b, c, d, x(k + 8), S41, &H6FA87E4F
md5_II d, a, b, c, x(k + 15), S42, &HFE2CE6E0
md5_II c, d, a, b, x(k + 6), S43, &HA3014314
md5_II b, c, d, a, x(k + 13), S44, &H4E0811A1
md5_II a, b, c, d, x(k + 4), S41, &HF7537E82
md5_II d, a, b, c, x(k + 11), S42, &HBD3AF235
md5_II c, d, a, b, x(k + 2), S43, &H2AD7D2BB
md5_II b, c, d, a, x(k + 9), S44, &HEB86D391
a = AddUnsigned(a, AA)
b = AddUnsigned(b, BB)
c = AddUnsigned(c, CC)
d = AddUnsigned(d, DD)
Next
'
MD5=LCase(WordToHex(b) & WordToHex(c))  '
End Function
Public Function MD5_16(sMessage)
m_lOnBits(0) = CLng(1)
m_lOnBits(1) = CLng(3)
m_lOnBits(2) = CLng(7)
m_lOnBits(3) = CLng(15)
m_lOnBits(4) = CLng(31)
m_lOnBits(5) = CLng(63)
m_lOnBits(6) = CLng(127)
m_lOnBits(7) = CLng(255)
m_lOnBits(8) = CLng(511)
m_lOnBits(9) = CLng(1023)
m_lOnBits(10) = CLng(2047)
m_lOnBits(11) = CLng(4095)
m_lOnBits(12) = CLng(8191)
m_lOnBits(13) = CLng(16383)
m_lOnBits(14) = CLng(32767)
m_lOnBits(15) = CLng(65535)
m_lOnBits(16) = CLng(131071)
m_lOnBits(17) = CLng(262143)
m_lOnBits(18) = CLng(524287)
m_lOnBits(19) = CLng(1048575)
m_lOnBits(20) = CLng(2097151)
m_lOnBits(21) = CLng(4194303)
m_lOnBits(22) = CLng(8388607)
m_lOnBits(23) = CLng(16777215)
m_lOnBits(24) = CLng(33554431)
m_lOnBits(25) = CLng(67108863)
m_lOnBits(26) = CLng(134217727)
m_lOnBits(27) = CLng(268435455)
m_lOnBits(28) = CLng(536870911)
m_lOnBits(29) = CLng(1073741823)
m_lOnBits(30) = CLng(2147483647)
m_l2Power(0) = CLng(1)
m_l2Power(1) = CLng(2)
m_l2Power(2) = CLng(4)
m_l2Power(3) = CLng(8)
m_l2Power(4) = CLng(16)
m_l2Power(5) = CLng(32)
m_l2Power(6) = CLng(64)
m_l2Power(7) = CLng(128)
m_l2Power(8) = CLng(256)
m_l2Power(9) = CLng(512)
m_l2Power(10) = CLng(1024)
m_l2Power(11) = CLng(2048)
m_l2Power(12) = CLng(4096)
m_l2Power(13) = CLng(8192)
m_l2Power(14) = CLng(16384)
m_l2Power(15) = CLng(32768)
m_l2Power(16) = CLng(65536)
m_l2Power(17) = CLng(131072)
m_l2Power(18) = CLng(262144)
m_l2Power(19) = CLng(524288)
m_l2Power(20) = CLng(1048576)
m_l2Power(21) = CLng(2097152)
m_l2Power(22) = CLng(4194304)
m_l2Power(23) = CLng(8388608)
m_l2Power(24) = CLng(16777216)
m_l2Power(25) = CLng(33554432)
m_l2Power(26) = CLng(67108864)
m_l2Power(27) = CLng(134217728)
m_l2Power(28) = CLng(268435456)
m_l2Power(29) = CLng(536870912)
m_l2Power(30) = CLng(1073741824)
Dim x
Dim k
Dim AA
Dim BB
Dim CC
Dim DD
Dim a
Dim b
Dim c
Dim d
Const S11 = 7
Const S12 = 12
Const S13 = 17
Const S14 = 22
Const S21 = 5
Const S22 = 9
Const S23 = 14
Const S24 = 20
Const S31 = 4
Const S32 = 11
Const S33 = 16
Const S34 = 23
Const S41 = 6
Const S42 = 10
Const S43 = 15
Const S44 = 21
x = ConvertToWordArray(sMessage)
a = &H67452301
b = &HEFCDAB89
c = &H98BADCFE
d = &H10325476
For k = 0 To UBound(x) Step 16
AA = a
BB = b
CC = c
DD = d
md5_FF a, b, c, d, x(k + 0), S11, &HD76AA478
md5_FF d, a, b, c, x(k + 1), S12, &HE8C7B756
md5_FF c, d, a, b, x(k + 2), S13, &H242070DB
md5_FF b, c, d, a, x(k + 3), S14, &HC1BDCEEE
md5_FF a, b, c, d, x(k + 4), S11, &HF57C0FAF
md5_FF d, a, b, c, x(k + 5), S12, &H4787C62A
md5_FF c, d, a, b, x(k + 6), S13, &HA8304613
md5_FF b, c, d, a, x(k + 7), S14, &HFD469501
md5_FF a, b, c, d, x(k + 8), S11, &H698098D8
md5_FF d, a, b, c, x(k + 9), S12, &H8B44F7AF
md5_FF c, d, a, b, x(k + 10), S13, &HFFFF5BB1
md5_FF b, c, d, a, x(k + 11), S14, &H895CD7BE
md5_FF a, b, c, d, x(k + 12), S11, &H6B901122
md5_FF d, a, b, c, x(k + 13), S12, &HFD987193
md5_FF c, d, a, b, x(k + 14), S13, &HA679438E
md5_FF b, c, d, a, x(k + 15), S14, &H49B40821
md5_GG a, b, c, d, x(k + 1), S21, &HF61E2562
md5_GG d, a, b, c, x(k + 6), S22, &HC040B340
md5_GG c, d, a, b, x(k + 11), S23, &H265E5A51
md5_GG b, c, d, a, x(k + 0), S24, &HE9B6C7AA
md5_GG a, b, c, d, x(k + 5), S21, &HD62F105D
md5_GG d, a, b, c, x(k + 10), S22, &H2441453
md5_GG c, d, a, b, x(k + 15), S23, &HD8A1E681
md5_GG b, c, d, a, x(k + 4), S24, &HE7D3FBC8
md5_GG a, b, c, d, x(k + 9), S21, &H21E1CDE6
md5_GG d, a, b, c, x(k + 14), S22, &HC33707D6
md5_GG c, d, a, b, x(k + 3), S23, &HF4D50D87
md5_GG b, c, d, a, x(k + 8), S24, &H455A14ED
md5_GG a, b, c, d, x(k + 13), S21, &HA9E3E905
md5_GG d, a, b, c, x(k + 2), S22, &HFCEFA3F8
md5_GG c, d, a, b, x(k + 7), S23, &H676F02D9
md5_GG b, c, d, a, x(k + 12), S24, &H8D2A4C8A
md5_HH a, b, c, d, x(k + 5), S31, &HFFFA3942
md5_HH d, a, b, c, x(k + 8), S32, &H8771F681
md5_HH c, d, a, b, x(k + 11), S33, &H6D9D6122
md5_HH b, c, d, a, x(k + 14), S34, &HFDE5380C
md5_HH a, b, c, d, x(k + 1), S31, &HA4BEEA44
md5_HH d, a, b, c, x(k + 4), S32, &H4BDECFA9
md5_HH c, d, a, b, x(k + 7), S33, &HF6BB4B60
md5_HH b, c, d, a, x(k + 10), S34, &HBEBFBC70
md5_HH a, b, c, d, x(k + 13), S31, &H289B7EC6
md5_HH d, a, b, c, x(k + 0), S32, &HEAA127FA
md5_HH c, d, a, b, x(k + 3), S33, &HD4EF3085
md5_HH b, c, d, a, x(k + 6), S34, &H4881D05
md5_HH a, b, c, d, x(k + 9), S31, &HD9D4D039
md5_HH d, a, b, c, x(k + 12), S32, &HE6DB99E5
md5_HH c, d, a, b, x(k + 15), S33, &H1FA27CF8
md5_HH b, c, d, a, x(k + 2), S34, &HC4AC5665
md5_II a, b, c, d, x(k + 0), S41, &HF4292244
md5_II d, a, b, c, x(k + 7), S42, &H432AFF97
md5_II c, d, a, b, x(k + 14), S43, &HAB9423A7
md5_II b, c, d, a, x(k + 5), S44, &HFC93A039
md5_II a, b, c, d, x(k + 12), S41, &H655B59C3
md5_II d, a, b, c, x(k + 3), S42, &H8F0CCC92
md5_II c, d, a, b, x(k + 10), S43, &HFFEFF47D
md5_II b, c, d, a, x(k + 1), S44, &H85845DD1
md5_II a, b, c, d, x(k + 8), S41, &H6FA87E4F
md5_II d, a, b, c, x(k + 15), S42, &HFE2CE6E0
md5_II c, d, a, b, x(k + 6), S43, &HA3014314
md5_II b, c, d, a, x(k + 13), S44, &H4E0811A1
md5_II a, b, c, d, x(k + 4), S41, &HF7537E82
md5_II d, a, b, c, x(k + 11), S42, &HBD3AF235
md5_II c, d, a, b, x(k + 2), S43, &H2AD7D2BB
md5_II b, c, d, a, x(k + 9), S44, &HEB86D391
a = AddUnsigned(a, AA)
b = AddUnsigned(b, BB)
c = AddUnsigned(c, CC)
d = AddUnsigned(d, DD)
Next
MD5_16 = LCase(WordToHex(a) & WordToHex(b) & WordToHex(c) & WordToHex(d))
End Function
function trimd(msg)
if msg="" then
trimd=" "
else
trimd=msg
end if
end function
function HTMLEncode2(fString)
fString = Replace(fString, CHR(13), "")
fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
fString = Replace(fString, CHR(10), "<BR>")
HTMLEncode2 = fString
end function
Function encode(str)
if str<>"" then
str=replace(str,chr(10),"<br>")
str=replace(str,chr(32),"&nbsp;")
end if
encode=str
End Function
Function encode2(str)
if str<>""then
str=replace(str,"&nbsp;",CHR(32))
str=replace(str,"<BR>",CHR(10))
str=replace(str,"</P>",CHR(10))
str=replace(str,"<P>",CHR(10))
end if
encode2=str
End Function
function HTMLEncode3(strContent)
dim objRegExp
Set objRegExp=new RegExp
objRegExp.IgnoreCase =true
objRegExp.Global=True
'
objRegExp.Pattern="(\[br\])"
strContent=objRegExp.Replace(strContent,"<br/>")
'
objRegExp.Pattern="(\[u\])(.+?)(\[\/u\])"
strContent=objRegExp.Replace(strContent,"<u>$2</u>")
'
objRegExp.Pattern="(\[i\])(.+?)(\[\/i\])"
strContent=objRegExp.Replace(strContent,"<i>$2</i>")
'
objRegExp.Pattern="(\[b\])(.+?)(\[\/b\])"
strContent=objRegExp.Replace(strContent,"<b>$2</b>")
'
objRegExp.Pattern="(\[QUOTE\])(.+?)(\[\/QUOTE\])"
strContent=objRegExp.Replace(strContent,"<BLOCKQUOTE><font size=2 face=""Verdana, Arial"">引用:</font><HR>$2<HR></BLOCKQUOTE>")
'
objRegExp.Pattern="(\[red\])(.+?)(\[\/red\])"
strContent=objRegExp.Replace(strContent,"<FONT COLOR=""#ff0000"">$2</FONT>")
'
objRegExp.Pattern="(\[gray\])(.+?)(\[\/gray\])"
strContent=objRegExp.Replace(strContent,"<FONT COLOR=""#77ACAC"">$2</FONT>")
'
objRegExp.Pattern="(\[green\])(.+?)(\[\/green\])"
strContent=objRegExp.Replace(strContent,"<FONT COLOR=""#009933"">$2</FONT>")
'
objRegExp.Pattern="(\[blue\])(.+?)(\[\/blue\])"
strContent=objRegExp.Replace(strContent,"<FONT COLOR=""#0055ff"">$2</FONT>")
'
objRegExp.Pattern="(\[color\=)(.+?)(\])(.+?)(\[\/color\])"
strContent=objRegExp.Replace(strContent,"<FONT COLOR=""$2"">$4</FONT>")
'
objRegExp.Pattern="(\[EMAIL\])(\S+\@\S+?)(\[\/EMAIL\])"
strContent= objRegExp.Replace(strContent,"<A HREF=""mailto:$2"">$2</A>")
'
objRegExp.Pattern="(\[URL\])(http:\/\/\S+?)(\[\/URL\])"
strContent= objRegExp.Replace(strContent,"<A HREF=""$2"" TARGET=_blank>$2</A>")
objRegExp.Pattern="(\[URL\])(\S+?)(\[\/URL\])"
strContent= objRegExp.Replace(strContent,"<A HREF=""http://$2"" TARGET=_blank>$2</A>")
'
objRegExp.Pattern="(\[marquee\])(.+?)(\[\/marquee\])"
strContent=objRegExp.Replace(strContent,"<marquee scrollamount='3' id=xxskybbs onmouseover=xxskybbs.stop() onmouseout=xxskybbs.start()>$2</marquee>")
'
objRegExp.Pattern="(\[marqueea\])(.+?)(\[\/marqueea\])"
strContent=objRegExp.Replace(strContent,"<marquee behavior=""alternate"" scrollamount='3' id=xxskybbs onmouseover=xxskybbs.stop() onmouseout=xxskybbs.start()>$2</marquee>")
'
objRegExp.Pattern="(\[IMGurl\=)(http:\/\/\S+?)(\])(http:\/\/\S+?)(\[\/IMGurl\])"
strContent=objRegExp.Replace(strContent,"<a href=""$2"" target=_blank><IMG SRC=""$4"" border=0 onload=""javascript:if(this.width>screen.width-366)this.width=screen.width-366""></a>")
objRegExp.Pattern="(\[IMGurl\=)(\S+?)(\])(\S+?)(\[\/IMGurl\])"
strContent=objRegExp.Replace(strContent,"<a href=""http://$2"" target=_blank><IMG SRC=""http://$4"" border=0 onload=""javascript:if(this.width>screen.width-366)this.width=screen.width-366""></a>")
'
objRegExp.Pattern="(\[IMG\])(\S+?)(\[\/IMG\])"
strContent=objRegExp.Replace(strContent,"<IMG SRC=""$2"">")
set objRegExp=Nothing
HTMLEncode3=strContent
end function
Function long_str(txt,length)
txt=trim(txt)
x = len(txt)
j = 0
if x >= 1 then
for ii = 1 to x
if asc(mid(txt,ii,1)) < 0 or asc(mid(txt,ii,1)) >255 then '
j = j + 2
else
j = j + 1
end if
if j >= length then
txt = left(trim(txt),ii)&".." '
exit for
end if
next
long_str = txt
else
long_str = ""
end if
End Function
on error resume next
dim wed_upload_path, gift_upload_path
wed_upload_path = "../wed/"
gift_upload_path = "../upload/gift/"
Call CheckSystemStatu()
Function U2UTF8(Byval a_iNum)
Dim sResult,sUTF8
Dim iTemp,iHexNum,i
iHexNum = Trim(a_iNum)
If iHexNum = "" Then
Exit Function
End If
sResult = ""
If (iHexNum < 128) Then
sResult = sResult & iHexNum
ElseIf (iHexNum < 2048) Then
sResult = ChrB(&H80 + (iHexNum And &H3F))
iHexNum = iHexNum \ &H40
sResult = ChrB(&HC0 + (iHexNum And &H1F)) & sResult
ElseIf (iHexNum < 65536) Then
sResult = ChrB(&H80 + (iHexNum And &H3F))
iHexNum = iHexNum \ &H40
sResult = ChrB(&H80 + (iHexNum And &H3F)) & sResult
iHexNum = iHexNum \ &H40
sResult = ChrB(&HE0 + (iHexNum And &HF)) & sResult
End If
U2UTF8 = sResult
End Function
Function GB2UTF(Byval a_sStr)
Dim sGB,sResult,sTemp
Dim iLen,iUnicode,iTemp,i
sGB = Trim(a_sStr)
iLen = Len(sGB)
For i = 1 To iLen
sTemp = Mid(sGB,i,1)
iTemp = Asc(sTemp)
If (iTemp>127 or iTemp<0) Then
iUnicode = AscW(sTemp)
If iUnicode<0 Then
iUnicode = iUnicode + 65536
End If
Else
iUnicode = iTemp
End If
sResult = sResult & U2UTF8(iUnicode)
Next
GB2UTF = sResult
End Function
Function FinalMoneySum(xmid,flag)
'
dim money1,money2,money3,money4,money5,finalmoney
money1 = conn.execute("select sum(money) from fujia where xiangmu_id="&xmid)(0)
if money1="" or isnull(money1) then money1=0
money2 = conn.execute("select sum(money) from fujia2 where xiangmu_id="&xmid)(0)
if money2="" or isnull(money2) then money2=0
money3 = conn.execute("select sum(money) from goumai where xiangmu_id="&xmid)(0)
if money3="" or isnull(money3) then money3=0
money4=conn.execute("select jixiang_money from shejixiadan where id="&xmid)(0)
if money4="" or isnull(money4) then money4=0
money5 = conn.execute("select sum(money) from save_money where isdelete=false and xiangmu_id="&xmid)(0)
if isnull(money5) then money5=0
dim rskh
set rskh=server.createobject("adodb.recordset")
rskh.open "select kehu_id from shejixiadan where id="&xmid,conn,1,1
if not (rskh.eof and rskh.bof) then
Call ChangeCustomerRank(rskh(0))
end if
rskh.close
set rskh  = nothing
if flag then	'
finalmoney = money1+money2+money3+money4-money5
dim rs_xm
set rs_xm=server.createobject("adodb.recordset")
rs_xm.open "select shejiwancheng, ReceivablesFlag from shejixiadan where id="&xmid,conn,1,3
if not (rs_xm.eof and rs_xm.bof) then
if rs_xm("shejiwancheng")=1 then
if finalmoney<=0 then
rs_xm("ReceivablesFlag")=2
else
rs_xm("shejiwancheng")=0
rs_xm("ReceivablesFlag")=0
end if
elseif rs_xm("shejiwancheng")=0 then
if finalmoney<=0 then
rs_xm("ReceivablesFlag")=1
else
rs_xm("ReceivablesFlag")=0
end if
end if
rs_xm.update()
end if
rs_xm.close()
set rs_xm=nothing
else
FinalMoneySum = money1+money2+money3+money4-money5
end if
end function
Function QujianMoneySum(xmid)
'
dim money1,money2,money4,money5,finalmoney
money1 = conn.execute("select sum(money) from fujia where xiangmu_id="&xmid)(0)
if money1="" or isnull(money1) then money1=0
money2 = conn.execute("select sum(money) from fujia2 where xiangmu_id="&xmid)(0)
if money2="" or isnull(money2) then money2=0
money4=conn.execute("select jixiang_money from shejixiadan where id="&xmid)(0)
if money4="" or isnull(money4) then money4=0
money5 = conn.execute("select sum(money) from save_money where [type]<>4 and isdelete=false and xiangmu_id="&xmid)(0)
if isnull(money5) then money5=0
QujianMoneySum = money1+money2+money4-money5
end function
function ChangeCustomerRank(khid)
Dim rsxm:Set rsxm = Server.CreateObject("ADODB.RECORDSET")
Dim xmid:xmid = ""
rsxm.open "select id from shejixiadan where kehu_id="&khid,conn,1,1
do while not rsxm.eof
xmid = xmid & "," & rsxm("id")
rsxm.movenext
loop
rsxm.close
if xmid<>"" then xmid = mid(xmid,2)
Set rsxm = nothing
dim money1,money2,money3,money4,moneys
money1 = conn.execute("select sum(money) from fujia where xiangmu_id in ("&xmid&")")(0)
if money1="" or isnull(money1) then money1=0
money2 = conn.execute("select sum(money) from fujia2 where xiangmu_id in ("&xmid&")")(0)
if money2="" or isnull(money2) then money2=0
money3 = conn.execute("select sum(money) from goumai where xiangmu_id in ("&xmid&")")(0)
if money3="" or isnull(money3) then money3=0
money4=conn.execute("select sum(jixiang_money) from shejixiadan where id in ("&xmid&")")(0)
if money4="" or isnull(money4) then money4=0
moneys = money1 + money2 + money3 + money4
dim rsrank
set rsrank = conn.execute("select id from customerrank where minmoney<="&moneys&" and maxmoney>="&moneys)
if not (rsrank.eof and rsrank.bof) then
conn.execute("update kehu set CustomerRank="&rsrank(0)&" where id="&khid)
end if
rsrank.close
set rsrank = nothing
'
'
'
'
'
'
'
end function
function CheckTaskEnd(xmid)
'
dim rscheck
set rscheck = server.createobject("adodb.recordset")
rscheck.open "select * from shejixiadan where id="&xmid,conn,1,1
if CheckNull(rscheck("hz_name")) or CheckNull(rscheck("cp_name")) then
CheckTaskEnd = False
else
CheckTaskEnd = True
end if
rscheck.close()
set rscheck=nothing
end function
function CheckQujianIsComplete(xmid)
err.clear()
dim rs,id,count11,counts,i,flag
dim y_id,y_sl,yj_id,yj,sl,qj_id,qj_sl
dim arr_y_id,arr_y_sl,arr_yj_id,arr_qj_id,arr_yj_sl,arr_qj_sl
set rs=server.CreateObject("adodb.recordset")
'
rs.open "select yunyong,sl from shejixiadan where id="&xmid,conn,1,1
if not rs.eof then
id=split(rs("yunyong"),", ")
sl=split(rs("sl"),", ")
count11=ubound(id)+1
for yy=1 to count11
set rs_yunyong=conn.execute("select [type],id,yunyong from yunyong where id="&id(yy-1))
if rs_yunyong("type")=1 then
if not rs_yunyong.eof then
y_id = y_id & "," & rs_yunyong("id")
y_sl = y_sl & "," & sl(yy-1)
end if
end if
rs_yunyong.close
set rs_yunyong = nothing
next
end if
if y_id<>"" then
y_id=mid(y_id,2)
y_sl=mid(y_sl,2)
end if
rs.close
rs.open "SELECT D.ProID,D.ProVol,L.vType FROM VerifyProDetails D INNER JOIN VerifyProList L ON D.MainID = L.ID WHERE L.Xiangmu_ID="&xmid&" AND D.Types=0",conn,1,1
do while not rs.eof
select case rs("vType")
case 0
yj_id = yj_id & "," & rs("ProID")
yj_sl = yj_sl & "," & rs("ProVol")
case 1
qj_id = qj_id & "," & rs("ProID")
qj_sl = qj_sl & "," & rs("ProVol")
end select
rs.movenext
loop
if yj_id<>"" then
yj_id=mid(yj_id,2)
yj_sl=mid(yj_sl,2)
end if
if qj_id<>"" then
qj_id=mid(qj_id,2)
qj_sl=mid(qj_sl,2)
end if
rs.close
'
arr_y_id = split(y_id,",")
arr_y_sl = split(y_sl,",")
if yj_id<>"" then
arr_yj_id = split(yj_id,",")
arr_yj_sl = split(yj_sl,",")
for i = 0 to ubound(arr_y_id)
flag = false
for k = 0 to ubound(arr_yj_id)
if CInt(arr_y_id(i)) = CInt(arr_yj_id(k)) then
flag = true
exit for
end if
next
if not flag then
'
CheckQujianIsComplete = false
exit function
end if
next
elseif y_id<>"" then
'
CheckQujianIsComplete = false
exit function
end if
if qj_id<>"" then
arr_qj_id = split(qj_id,",")
arr_qj_sl = split(qj_sl,",")
for i = 0 to ubound(arr_y_id)
flag = false
for k = 0 to ubound(arr_qj_id)
if CInt(arr_y_id(i)) = CInt(arr_qj_id(k)) then
flag = true
exit for
end if
next
if not flag then
'
CheckQujianIsComplete = false
exit function
end if
next
elseif y_id<>"" then
'
CheckQujianIsComplete = false
exit function
end if
y_id = ""
y_sl = ""
'
rs.open "select fujia.* from fujia inner join yunyong on fujia.jixiang=yunyong.id where yunyong.type=1 and fujia.xiangmu_id="&xmid&" order by times",conn,1,1
if not (rs.eof and rs.bof) then
do while not rs.eof
y_id = y_id & "," & rs("jixiang")
y_sl = y_sl & "," & rs("sl")
rs.movenext
loop
end if
if y_id<>"" then
y_id=mid(y_id,2)
y_sl=mid(y_sl,2)
end if
rs.close
yj_id = ""
yj_sl = ""
qj_id = ""
qj_sl = ""
rs.open "SELECT D.ProID,D.ProVol,L.vType FROM VerifyProDetails D INNER JOIN VerifyProList L ON D.MainID = L.ID WHERE L.Xiangmu_ID="&xmid&" AND D.Types=1",conn,1,1
do while not rs.eof
select case rs("vType")
case 0
yj_id = yj_id & "," & rs("ProID")
yj_sl = yj_sl & "," & rs("ProVol")
case 1
qj_id = qj_id & "," & rs("ProID")
qj_sl = qj_sl & "," & rs("ProVol")
end select
rs.movenext
loop
if yj_id<>"" then
yj_id=mid(yj_id,2)
yj_sl=mid(yj_sl,2)
end if
if qj_id<>"" then
qj_id=mid(qj_id,2)
qj_sl=mid(qj_sl,2)
end if
rs.close
set rs = nothing
arr_y_id = split(y_id,",")
arr_y_sl = split(y_sl,",")
if yj_id<>"" then
arr_yj_id = split(yj_id,",")
arr_yj_sl = split(yj_sl,",")
for i = 0 to ubound(arr_y_id)
flag = false
for k = 0 to ubound(arr_yj_id)
if CInt(arr_y_id(i)) = CInt(arr_yj_id(k)) then
flag = true
exit for
end if
next
if not flag then
'
CheckQujianIsComplete = false
exit function
end if
next
elseif y_id<>"" then
'
CheckQujianIsComplete = false
exit function
end if
if qj_id<>"" then
arr_qj_id = split(qj_id,",")
arr_qj_sl = split(qj_sl,",")
for i = 0 to ubound(arr_y_id)
flag = false
for k = 0 to ubound(arr_qj_id)
if CInt(arr_y_id(i)) = CInt(arr_qj_id(k)) then
flag = true
exit for
end if
next
if not flag then
'
CheckQujianIsComplete = false
exit function
end if
next
elseif y_id<>"" then
'
CheckQujianIsComplete = false
exit function
end if
CheckQujianIsComplete = true
end function
Function CheckNonCompleteQujian()
dim EnrolBackGroups(8)
dim StrProID,StrProMoney,StrProVol,StrProMemo,StrCompName,StrMemo,StrOther
Dim StrSourcePro,StrSourceMemo,ArrSourcePro,ArrSourceMemo,tmp_memo,tmp_arrmemo,tmp_flag,tmp_counter
dim ArrProID,ArrProMoney,ArrProVol,ArrProMemo
dim rsenrol,rsve
Dim rsorder,cc
cc=0
Set rsorder = server.CreateObject("adodb.recordset")
rsorder.open "select * from ProcessEnrolOrder where isnull(wctime) order by id desc",conn,1,3
Do While Not rsorder.eof
StrProID = ""
StrProMoney = ""
StrProVol = ""
StrProMemo = ""
StrCompName = ""
StrMemo = ""
StrOther = ""
StrSourcePro = ""
StrSourceMemo = ""
Set rsenrol = server.createobject("adodb.recordset")
rsenrol.open "SELECT o.OrderNo, o.CompID, o.LcTime, o.Memo, d.* FROM ProcessEnrolOrder o INNER JOIN ProcessEnrolDetails d ON o.ID = d.OrderID where d.orderid="&rsorder("id"),conn,1,1
do while not rsenrol.eof
tmp_flag = true
tmp_memo = Trim(rsenrol("ProMemo"))
If InStr(tmp_memo,",")>0 Then
tmp_arrmemo = Split(tmp_memo,",")
tmp_memo = ""
For tmp_counter = 0 To UBound(tmp_arrmemo)
tmp_arrmemo(tmp_counter) = Trim(tmp_arrmemo(tmp_counter))
If tmp_arrmemo(tmp_counter)<>"" And IsNumeric(tmp_arrmemo(tmp_counter)) Then
tmp_arrmemo(tmp_counter) = Trim(tmp_arrmemo(tmp_counter))
Set rsve = conn.execute("SELECT D.ProID, D.ProVol FROM VerifyProDetails D INNER JOIN VerifyProList L ON D.MainID = L.ID WHERE D.ProID="&rsenrol("ProID")&" AND L.vType=0 AND L.Xiangmu_ID="&tmp_arrmemo(tmp_counter))
If rsve.eof And rsve.bof Then
tmp_memo = tmp_memo & "," & tmp_arrmemo(tmp_counter)
End If
rsve.Close
Set rsve = Nothing
End If
Next
If tmp_memo <> "" Then
tmp_memo = Mid(tmp_memo,2)
Else
tmp_flag = false
End If
Else
If tmp_memo<>"" And IsNumeric(tmp_memo) then
Set rsve = conn.execute("SELECT D.ProID, D.ProVol FROM VerifyProDetails D INNER JOIN VerifyProList L ON D.MainID = L.ID WHERE D.ProID="&rsenrol("ProID")&" AND L.vType=0 AND L.Xiangmu_ID="&tmp_memo)
If Not (rsve.eof And rsve.bof) Then
tmp_flag = false
End If
rsve.Close
Set rsve = Nothing
End if
End If
If tmp_flag Then
StrProID = StrProID & "|" & rsenrol("ProID")
StrProMoney = StrProMoney & "|" & rsenrol("ProMoney")
StrProVol = StrProVol & "|" & rsenrol("ProVol")
StrProMemo = StrProMemo & "|" & tmp_memo
StrCompName = rsenrol("CompID")
StrMemo = rsenrol("Memo")
StrOther = orderid&", "&rsenrol("OrderNo")&", "&rsenrol("LcTime")&", "
StrSourcePro = StrSourcePro & "|" & rsenrol("ProID")
StrSourceMemo = StrSourceMeMo & "|" & tmp_memo
End If
rsenrol.movenext
loop
rsenrol.close
set rsenrol = Nothing
If StrProID = "" Then
cc=cc+1
response.write cc&"."&rsorder("orderno")&"<br>"
rsorder("wctime")=Now()
rsorder("wcadmin")=rsorder("adminid")
rsorder.update
End If
rsorder.movenext
Loop
rsorder.close
Set rsorder = Nothing
response.write "更新完成"
End Function
function CheckNull(str)
if str="" or isnull(str) then
CheckNull = True
else
CheckNull = False
end if
end function
function CheckWedSurplus(id)
dim arr(3)
allsl=conn.execute("select sl from huensha where id="&id)(0)
yzsl = conn.execute("select sum(volume) from chuzhu_details where AnnexWedID="&id&" and [flag]=1")(0)
if yzsl="" or isnull(yzsl) or yzsl<=0 then
sysl = allsl
yzsl=0
else
sysl = allsl-yzsl
end if
washsl = conn.execute("select sum(sl) from hs_washlist where hs_id="&id&" and flag=0")(0)
if washsl>0 then
sysl=sysl-washsl
else
washsl=0
end if
arr(0)=allsl
arr(1)=yzsl
arr(2)=washsl
arr(3)=sysl
CheckWedSurplus=arr
end function
function unencode(str)
str=replace(str,"&amp;","&")
str=replace(str,"&quot;",chr(34))
str=replace(str,"&lt;","<")
str=replace(str,"&gt;",">")
unencode=str
end function
function smsencode(str)
str=replace(str,"""","“")
str=replace(str,"'","‘")
str=replace(str," ","")
str=replace(str,"<BR>","")
str=replace(str,vbcrlf,"")
smsencode=str
end function
Sub CheckSystemStatu()
dim sys_statu,rssys
set rssys = server.createobject("adodb.recordset")
rssys.open "select * from sysconfig",conn,1,1
if session("level")<>10 then
if rssys("SystemStatu") = 1 then
response.write "<script language='javascript'>top.location.href='../showerr.asp';</script>"
On Error GoTo 0
Err.Raise 9999
end if
end if
if not isnull(rssys("ExMaxNumDate")) then
if (rssys("ExMaxNumDate")<date()) then
response.write "<script language='javascript'>top.location.href='../showerr.asp';</script>"
On Error GoTo 0
Err.Raise 9999
end if
end if
rssys.close
set rssys = nothing
End Sub
function getPerStep(xmid)
set xmrs=server.CreateObject("adodb.recordset")
xmrs.open "select * from shejixiadan where id="&xmid,conn,1,1
'
'
'
'
if not isnull(xmrs("lc_xp2")) then
getPerStep="取件"'
elseif not isnull(xmrs("xg_sj")) then
getPerStep="精修外发"
elseif not isnull(xmrs("lc_sj")) then
getPerStep="看版"
elseif not isnull(xmrs("lc_ky")) then
getPerStep="设计"
elseif not isnull(xmrs("lc_xp")) then
getPerStep="选片"
elseif not isnull(xmrs("lc_cp")) and not isnull(xmrs("lc_hz")) then
getPerStep="调色"
'
'
else
getPerStep="化妆/摄影"
end if
xmrs.close
set xmrs=nothing
end function
function GetFlowInfo(flow)
dim result(2)
result(0)=flow
select case flow
case "hzsy"
result(1)="化妆/摄影"
result(2)="lc_hz|lc_cp"
case "xp"
result(1)="调色"
result(2)="lc_xp"
case "ky"
result(1)="选片"
result(2)="lc_ky"
case "sj"
result(1)="设计"
result(2)="lc_sj"
case "xg"
result(1)="看版"
result(2)="xg_sj"
case "xp2"
result(1)="精修外发"
result(2)="lc_xp2"
case "qj"
result(1)="取件"
result(2)="lc_wc"
end select
GetFlowInfo=result
end function
function GetUserInfo(setfield,getfield,fieldvalue)
dim rsuser
set rsuser = conn.execute("select "&getfield&" from yuangong where "&setfield&"='"&fieldvalue&"'")
if not rsuser.eof then
GetUserInfo = rsuser(0)
else
GetUserInfo = "N/A"
end if
rsuser.close()
set rsuser=nothing
end function
function GetUserGroupName(userid)
'
dim gn,ul,rslev
set rslev = conn.execute("select [level] from yuangong where username='"&userid&"'")
if not rslev.eof then
ul = rslev(0)
rslev.close()
if ul=10 then
GetUserGroupName = "总经理"
exit function
end If
gn = GetDutyName(ul)
GetUserGroupName = gn
else
GetUserGroupName = "N/A"
end if
end function
Function CheckWorkCompFlag()
Dim UserLevel,AllOutString,IsCSShoot
UserLevel = Session("level")
AllOutString = ""
IsCSShoot=conn.execute("select IsCSShoot from sysconfig")(0)
Select Case UserLevel
Case 1	'
'
Call DatabaseBackup()
If instr(session("level2"),"108")>0 Then outString1 = CheckRecord("pz_time","lc_cp","","cp_name","",False,1*24*60*60,"客服拍摄方案","摄影")
xgInvis=conn.execute("select xgInvis from sysconfig")(0)
if xgInvis=1 then
outString2 = CheckRecord("xg_time","xg_sj","","xg_name,xp2_name","",False,1*24*60*60,"看版","看版")
end if
if instr(session("level2"),"109")>0 Then outString3 = CheckRecord("kj_time","lc_ky","","","",False,1*24*60*60,"选片","选片")
if instr(session("level2"),"110")>0 Then outString4 = CheckRecord("qj_time","lc_wc","","","",False,1*24*60*60,"取件","取件")
AllOutString = StringPortfolio(outString1&"|"&outString2&"|"&outString3&"|"&outString4)
Case 2	'
tsInvis=conn.execute("select tsInvis from sysconfig")(0)
if tsInvis=2 or tsInvis=6 then
'
outString1 = CheckRecord("ts_time","lc_xp","xp_name","","",False,1*24*60*60,"调色","限制考勤")
end if
outString7 = CheckRecord("jx_time","lc_jx","jx_name","","",False,1*24*60*60,"精修","限制考勤")
xgInvis=conn.execute("select xgInvis from sysconfig")(0)
if xgInvis=2 then
outString2 = CheckRecord("xg_time","xg_sj","sj_name|lc_sj","xg_name,xp2_name","",False,1*24*60*60,"看版","看版")
end if
outString3 = CheckRecord("sc_time","lc_sj","sj_name","","",False,1*24*60*60,"看版催单","限制考勤")
outString4 = CheckRecord("sc_time","lc_sj","","sj_name","lc_ky",False,1*24*60*60,"设计","限制考勤")
if instr(session("level2"),"101")>0 then
outString5 = CheckRecord("xp2_time","lc_xp2","","xp2_name","xg_sj",False,1*24*60*60,"外发","外发")
end if
if instr(session("level2"),"716")>0 then outString6 = CheckEnrol()
AllOutString = StringPortfolio(outString1&"|"&outString7&"|"&outString2&"|"&outString3&"|"&outString4&"|"&outString5&"|"&outString6)
Case 13	'
if instr(session("level2"),"716")>0 then
outString1 = CheckEnrol()
end if
outString2 = CheckRecord("qj_time","lc_wc","","","",False,1*24*60*60,"取件","取件")
'
AllOutString = StringPortfolio(outString1&"|"&outString2)
Case 4	'
If IsCSShoot=0 Then
outString1 = CheckRecord("pz_time","lc_cp","","cp_name","",False,1*24*60*60,"摄影","摄影")
End If
tsInvis=conn.execute("select tsInvis from sysconfig")(0)
if tsInvis=4 or tsInvis=6 then
outString2 = CheckRecord("kj_time","lc_xp","","","lc_cp",False,1*24*60*60,"调色","调色")
end if
AllOutString = StringPortfolio(outString1&"|"&outString2)
Case 12	'
If IsCSShoot=0 Then
outString1 = CheckRecord("pz_time","cpzl_name","","lc_xp","",False,1*24*60*60,"摄影助理","摄影")
End if
tsInvis=conn.execute("select tsInvis from sysconfig")(0)
if tsInvis=4 or tsInvis=6 then
outString2 = CheckRecord("kj_time","lc_xp","","","",False,1*24*60*60,"调色","调色")
end if
AllOutString = StringPortfolio(outString1&"|"&outString2)
Case 5	'
If IsCSShoot=0 Then
outString1 = CheckRecord("pz_time","hz_name","","","",False,1*24*60*60,"拍照化妆","拍照化妆")
End If
outString2 = CheckRecord("hz_time","hz_userid","","","",False,1*24*60*60,"结婚化妆","结婚化妆")
AllOutString = StringPortfolio(outString1&"|"&outString2)
Case 6	'
If IsCSShoot=1 Then
outString1 = CheckRecord("pz_time","lc_cp","","cp_name","",False,1*24*60*60,"客服拍摄方案","摄影")
End If
AllOutString = StringPortfolio(outString1)
Case 7	'
'
Call DatabaseBackup()
if session("level")=7 and instr(session("level2"),"711")>0 then
outString = CheckEnrol()
end if
AllOutString = StringPortfolio(outString)
Case 14	'
If IsCSShoot=0 Then
outString1 = CheckRecord("pz_time","hz_name","","","",False,1*24*60*60,"拍照化妆","拍照化妆")
End If
outString2 = CheckRecord("hz_time","hz_userid","","","",False,1*24*60*60,"结婚化妆","结婚化妆")
AllOutString = StringPortfolio(outString1&"|"&outString2)
End Select
if AllOutString<>"" then
Call OutputMsg(AllOutString)
CheckWorkCompFlag = False
Else
CheckWorkCompFlag = True
End If
End Function
Function CheckEnrol()
Dim rs,Seconds,sql,rsgys,gysname,outString,EnrolDisabled
EnrolDisabled = GetFieldDataBySQL("select EnrolDisabled from sysconfig","int",0)
if EnrolDisabled=1 then CheckEnrol=""
Set rs = Server.CreateObject("ADODB.RECORDSET")
Seconds = 1*24*60*60
sql = "SELECT TOP 10 * FROM ProcessEnrolOrder WHERE 1=1"
sql = sql & " AND DateDiff('s',LcTime,now())>="&Seconds&" AND NOT ISNULL(LcTime)"
sql = sql & " AND ISNULL(WcTime)"
sql = sql & " ORDER BY LcTime"
rs.open sql,conn,1,1
If NOT rs.eof then
outString = "考勤失败,您有未完成的 送制作回件 记录,部分如下:\t\t\n\n"
outString = outString & "单号\t\t供应商\t\t预设回件日期\n"
outString = outString & "---------------------------------------          \n"
Do While Not rs.eof
num = num + 1
set rsgys = conn.execute("select companyname from changshang where id="&rs("compid"))
if not rsgys.eof then
gysname = rsgys("companyname")
else
gysname = "N/A"
end if
rsgys.close()
outString = outString & rs("orderno") & "\t"&gysname
if len(gysname)<=4 then
outString = outString & "\t\t"
else
outString = outString & "\t"
end if
outString = outString & rs("LcTime")
outString = outString & "\n"
rs.movenext
if num >= 10 then Exit Do
Loop
rs.close()
set rs = nothing
CheckEnrol = outString
Else
rs.close()
set rs = nothing
CheckEnrol = ""
End If
End Function
'
'
'
'
'
'
'
'
Function CheckRecord(chkFieldName, wcFieldName, userFieldName, nullFieldList, notNullFieldList, chkMenshi, Seconds, tName, tName2)
Set rs = Server.CreateObject("ADODB.RECORDSET")
sql = "SELECT TOP 10 s.id,s.kehu_id"
if chkFieldName<>"" then sql = sql & ",s." & chkFieldName
sql = sql & " FROM [shejixiadan] s INNER JOIN [kehu] k on s.kehu_id=k.id WHERE 1=1"
'
if chkFieldName<>"" Then
'
If Instr(chkFieldName,",")>0 then
arrChkFieldName = Split(chkFieldName,",")
For k = 0 to UBound(arrChkFieldName)
if Seconds>=0 then
chksql = chksql & " OR (DateDiff('s',s."&arrChkFieldName(k)&",now())>="&Seconds&" AND NOT ISNULL(s."&arrChkFieldName(k)&"))"
else
chksql = chksql & " OR ((DateDiff('s',s."&arrChkFieldName(k)&",now())<"&abs(Seconds)&" OR s."&arrChkFieldName(k)&"<#"&Now()+(Seconds/24/60/60)&"#) AND NOT ISNULL(s."&arrChkFieldName(k)&"))"
end if
if k = 0 then checkField = arrChkFieldName(k)
if k = UBound(arrChkFieldName) then orderField = arrChkFieldName(k)
Next
Else
if Seconds>=0 then
chksql = chksql & " OR (DateDiff('s',s."&chkFieldName&",now())>="&Seconds&" AND NOT ISNULL(s."&chkFieldName&"))"
else
chksql = chksql & " OR ((DateDiff('s',s."&chkFieldName&",now())<"&abs(Seconds)&" OR s."&chkFieldName&"<#"&Now()+(Seconds/24/60/60)&"#) AND NOT ISNULL(s."&chkFieldName&"))"
end if
checkField = chkFieldName
orderField = chkFieldName
'
End If
end if
sql = sql & " AND (" & mid(chksql,5) & ")"
if wcFieldName<>"" then sql = sql & " AND ISNULL(s."&wcFieldName&")"
sql = sql & " AND ISNULL(s.wc_name) AND ISNULL(s.lc_wc) AND s.shejiwancheng=0"
if userFieldName<>"" then
dim newUserFieldName
if instr(userFieldName,"|")>0 then
newUserFieldName = split(userFieldName,"|")
sql = sql & " AND ((ISNULL(s." & newUserFieldName(0) & ") AND ISNULL(s." & newUserFieldName(1) & ")) OR (s." & newUserFieldName(0) & "='"&session("username")&"' AND ISNULL(s." & newUserFieldName(1) & ")))"
else
sql = sql & " AND s." & userFieldName & "='"&session("username")&"'"
end if
end if
if chkMenshi then
sql = sql & " AND (s.userid='"&session("userid")&"' OR s.userid2='"&session("userid")&"' OR s.userid3='"&session("userid")&"')"
end if
If nullFieldList<>"" Then
If Instr(nullFieldList,",")>0 then
arrNullFieldList = Split(nullFieldList,",")
For k = 0 to UBound(arrNullFieldList)
sql = sql & " AND ISNULL(s."&arrNullFieldList(k)&")"
Next
Else
sql = sql & " AND ISNULL(s."&nullFieldList&")"
End If
End If
If notNullFieldList<>"" Then
If Instr(notNullFieldList,",")>0 then
arrNotNullFieldList = Split(notNullFieldList,",")
For k = 0 to UBound(arrNotNullFieldList)
sql = sql & " AND NOT ISNULL(s."&arrNotNullFieldList(k)&")"
Next
Else
sql = sql & " AND NOT ISNULL(s."&notNullFieldList&")"
End If
End If
'
'
'
'
sql = sql & " AND k.shopid=" & session("UserShopID")
'
if orderField<>"" then sql = sql & " ORDER BY s."&orderField
'
'
rs.Open sql,conn,1,1
'
'
'
'
'
num = 0
If NOT rs.eof then
outString = "考勤失败,您有未完成的 "&tName&" 记录,部分如下:\n\n"
outString = outString & "单号\t客户姓名\t预设"&tName2&"日期\n"
outString = outString & "---------------------------------------          \n"
Do While Not rs.eof
num = num + 1
khname = GetFieldDataBySQL("select lxpeple from kehu where id="&rs("kehu_id"),"str","N/A")
outString = outString & rs("id") & "\t" & khname
if chkfieldname<>"" then outString = outString & "\t\t" & rs(checkField)
outString = outString & "\n"
rs.movenext
if num >= 10 then Exit Do
Loop
rs.close()
set rs = nothing
'
'
'
CheckRecord = outString
Else
rs.close()
set rs = nothing
CheckRecord = ""
End If
End Function
Sub OutputMsg(str)
StepControl = conn.execute("select StepControl from sysconfig")(0)
if str<>"" and StepControl=0 then
Response.Write "<script language=javascript>"
Response.Write "alert('"&str&"');"
'
Response.Write "</script>"
end if
End Sub
Function StringPortfolio(strings)
If Instr(strings,"|")>0 then
AllOutString = ""
arr_str = split(strings,"|")
for i = 0 to UBound(arr_str)
If arr_str(i)<>"" then
if AllOutString<>"" then
AllOutString = AllOutString & "\n\n" & arr_str(i)
else
AllOutString = AllOutString & arr_str(i)
end if
End If
Next
StringPortfolio = AllOutString
Else
If strings<>"" then StringPortfolio = strings
End If
End Function
function GetWedVol(xmid)
dim rs,yyid,yysl,rsflag,lfcount
lfcount=0
set rs=conn.execute("select yunyong,sl from shejixiadan where id="&xmid)
if not (rs.eof and rs.bof) then
if rs("yunyong")<>"" and not isnull(rs("yunyong")) then
yyid=split(rs("yunyong"),", ")
yysl=split(rs("sl"),", ")
for yy=0 to ubound(yyid)
set rsflag = conn.execute("select [type3] from yunyong where id="&yyid(yy))
if not rsflag.eof and rsflag("type3")=1 then
lfcount=lfcount+yysl(yy)
end if
rsflag.close()
set rsflag=nothing
next
end if
end if
rs.close
set rs=nothing
GetWedVol=lfcount
end function
Function EditedTimeSaveToReport(xmid,e,task,time1,time2)
dim userid,peplename,n_time1,n_time2,n_task
if session("username")<>"" then
userid=session("userid")
peplename=session("username")
else
userid=conn.execute("select userid from shejixiadan where id="&xmid)(0)
peplename=conn.execute("select peplename from yuangong where username='"&userid&"'")(0)
end if
if time1="" then
n_time1="空"
else
n_time1=time1
end if
if time2="" then
n_time2="空"
else
n_time2=time2
end if
select case task
case "pz"
n_task = "摄影1"
case "pz2"
n_task = "摄影2"
case "pz3"
n_task = "摄影3"
case "hz"
n_task = "化妆"
case "hhz"
n_task = "回婚妆"
case "pzlf"
n_task = "拍照礼服"
case "jhlf"
n_task = "结婚礼服"
case "qj"
n_task = "取件"
case "qj2"
n_task = "取件2"
case "xg"
n_task = "看版"
case "xp2"
n_task = "外发"
case "kj"
n_task = "选片"
case "jx"
n_task = "精修"
case else
n_task = "未知"
end select
conn.execute("insert into sjs_baobiao (xiangmu_id,EventID,userid,baobiao,topeple,times) values ("&xmid&","&e&",'"&userid&"','"&peplename&" 将 "&n_task&"日期 "&n_time1&" 更换为 "&n_time2&"','所有人',#"&now()&"#)")
End Function
Function EditedTimeSaveToReport2(xmid,dict)
dim n_task,desc
if dict.Count>0 then
desc = "下单预设流程时间如下：<br />"
for each fn in dict
select case fn
case "pz"
n_task = "摄影1"
case "pz2"
n_task = "摄影2"
case "hz"
n_task = "化妆"
case "hhz"
n_task = "回婚妆"
case "pzlf"
n_task = "拍照礼服"
case "jhlf"
n_task = "结婚礼服"
case "qj"
n_task = "取件"
case "xg"
n_task = "看版"
case "xp2"
n_task = "精修外发"
case "kj"
n_task = "选片"
case else
n_task = "未知"
end select
if dict(fn)<>"" and not isnull(dict(fn)) then
desc = desc & "&nbsp;&nbsp;&nbsp;&nbsp;" & n_task & "：" & dict(fn) & "<br />"
end if
next
end if
conn.execute("insert into sjs_baobiao (xiangmu_id,userid,baobiao,topeple,times) values ("&xmid&",'"&userid&"','"&peplename&" 将 "&n_task&"日期 "&n_time1&" 更换为 "&n_time2&"','所有人',#"&now()&"#)")
End Function
Function EditedMoneySaveToReport(xmid,e,money1,money2)
dim msg
msg = session("username")&" 将 套系价格 "&money1&" 更换为 "&money2
conn.execute("insert into sjs_baobiao (xiangmu_id,EventID,userid,baobiao,topeple,times) values ("&xmid&","&e&",'"&session("userid")&"','"&msg&"','所有人',#"&now()&"#)")
End Function
Function EditedJhzstyleSaveToReport(xmid,e,s1,s2)
dim info1,info2,msg
if instr(s2,"1")>0 then info1 = info1 & ", 收费妆"
if instr(s2,"2")>0 then info1 = info1 & ", 免费妆"
if instr(s1,"1")>0 then info2 = info2 & ", 收费妆"
if instr(s1,"2")>0 then info2 = info2 & ", 免费妆"
if info1<>"" then
info1 = mid(info1, 3)
else
info1 = "无"
end if
if info2<>"" then
info2 = mid(info2, 3)
else
info2 = "无"
end if
msg = session("username")&" "&" 调整配送结婚&nbsp;&nbsp;原&nbsp;"&info2&"&nbsp;修改为&nbsp;"&info1
conn.execute("insert into sjs_baobiao (xiangmu_id,EventID,userid,baobiao,topeple,times) values ("&xmid&","&e&",'"&session("userid")&"','"&msg&"','所有人',#"&now()&"#)")
End Function
Function EditedSl2SaveToReport(xmid,e,s1,s2)
msg = session("username")&" "&" 调整拍摄多款张数&nbsp;&nbsp;原&nbsp;"&s1&" 张,&nbsp;修改为&nbsp;"&s2&" 张"
conn.execute("insert into sjs_baobiao (xiangmu_id,EventID,userid,baobiao,topeple,times) values ("&xmid&","&e&",'"&session("userid")&"','"&msg&"','所有人',#"&now()&"#)")
End Function
Function DelCustomerInfoSaveToReport(e,msg)
dim userid
if session("userid")<>"" then
userid=session("userid")
else
userid=conn.execute("select userid from shejixiadan where id="&xmid)(0)
end if
conn.execute("insert into sjs_baobiao (EventID,userid,baobiao,times) values ("&e&",'"&userid&"','"&msg&"',#"&now()&"#)")
End Function
Function EditedCpvolumeSaveToReport(xmid,e,msg)
dim userid
if session("userid")<>"" then
userid=session("userid")
else
userid=conn.execute("select userid from shejixiadan where id="&xmid)(0)
end if
conn.execute("insert into sjs_baobiao (xiangmu_id,EventID,userid,baobiao,topeple,times) values ("&xmid&","&e&",'"&userid&"','"&msg&"','所有人',#"&now()&"#)")
End Function
Function EditedYunyongSaveToReport(xmid,e,msg)
dim userid
if session("userid")<>"" then
userid=session("userid")
else
userid=conn.execute("select userid from shejixiadan where id="&xmid)(0)
end if
conn.execute("insert into sjs_baobiao (xiangmu_id,EventID,userid,baobiao,topeple,times) values ("&xmid&","&e&",'"&userid&"','"&msg&"','所有人',#"&now()&"#)")
End Function
'
'
'
Function CheckEvent_Add(xmid,eventtype,tablename,fieldname,value1,value2)
err.clear()
dim userid,peplename
if session("username")<>"" then
userid=session("userid")
peplename=session("username")
else
userid=conn.execute("select userid from shejixiadan where id="&xmid)(0)
peplename=conn.execute("select peplename from yuangong where username='"&userid&"'")(0)
end if
Dim rsce,temp,eid
Set rsce = Server.CreateObject("adodb.recordset")
rsce.open "select top 1 * from CheckEvent",conn,1,3
rsce.addnew
rsce("xiangmu_id")=xmid
rsce("eventtype")=eventtype
rsce("tablename")=tablename
rsce("fieldname")=fieldname
rsce("value1")=value1
rsce("value2")=value2
rsce("peplename")=peplename
rsce.update
temp = rsce.bookmark
rsce.bookmark = temp
eid=rsce("ID")
if err.number>0 then
CheckEvent_Add = -1
err.clear()
else
CheckEvent_Add = eid
end if
rsce.close
set rsce = nothing
End Function
'
Function CheckEvent_Edit(eventid,checkflag)
conn.execute("update CheckEvent set CheckFlag="&checkflag&",CheckAdmin='"&session("username")&"' where CheckFlag=0 and id in ("&eventid&")")
End Function
Function CheckEvent_Del(eventid)
conn.execute("delete from CheckEvent where id in ("&eventid&")")
conn.execute("update sjs_baobiao set eventid=0 where eventid in ("&eventid&")")
End Function
'
Function CheckEvent_Rollback(eventid)
End Function
'
Function DatabaseBackup()
Set MyFileFolder = New FileFolderCls
spath = server.mappath("/")
BaseName = "hyx_dd.mdb"
if MyFileFolder.ReportFolderStatus(spath&"\backup") = 1 Then
Call MyFileFolder.RenameFolder(spath&"\backup",spath&"\1")
End If
sMonth = Month(Date())
sDay = Day(Date())
sHour = Hour(Time())
If sMonth<10 Then sMonth = "0"&sMonth
If sDay<10 Then sDay = "0"&sDay
FileName = Year(Date())&sMonth&sDay
FileList = MyFileFolder.ShowFileList(spath&"\1")
If Instr(FileList,Year(Date())&sMonth&sDay)>0 Then Exit Function
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
DatabaseBackup = MyFileFolder.CopyAFile(spath&"\"&BaseName,spath&"\1\"&FileName&".bc")
End Function
Function GetTelNo(no)
dim str_dep1,str_dep2,CustSourEncryp
str_dep1 = "6,10"
str_dep2 = "2,4,5,12,13,14"
CustSourEncryp = GetFieldDataBySQL("select isEncryp from sysconfig","int",0)
if no="" or isnull(no) then
GetTelNo = ""
else
if CustSourEncryp=1 then
if instr(","&str_dep1&",",","&session("level")&",")>0 then
GetTelNo = no
else
GetTelNo = "********"
end if
elseif CustSourEncryp=2 then
if instr(","&str_dep2&",",","&session("level")&",")<=0 then
GetTelNo = no
else
GetTelNo = "********"
end if
else
GetTelNo = no
end if
end if
End Function
Function GetTelNoFlag(no)
dim str_dep1,str_dep2,CustSourEncryp
str_dep1 = "6,10"
str_dep2 = "2,4,5,12,13,14"
CustSourEncryp = GetFieldDataBySQL("select isEncryp from sysconfig","int",0)
if no="" or isnull(no) then
GetTelNoFlag = "text"
else
if CustSourEncryp=1 then
if instr(","&str_dep1&",",","&session("level")&",")>0 then
GetTelNoFlag = "text"
else
GetTelNoFlag = "password"
end if
elseif CustSourEncryp=2 then
if instr(","&str_dep2&",",","&session("level")&",")<=0 then
GetTelNoFlag = "text"
else
GetTelNoFlag = "password"
end if
else
GetTelNoFlag = "text"
end if
end if
End Function
Function GetTaskName(smstype)
Select Case smstype
Case "new"
GetTaskName="接单"
Case "ky"
GetTaskName="选片"
Case "qj"
GetTaskName="取件"
Case "js"
GetTaskName="介绍"
Case "lott"
GetTaskName="抽奖"
End Select
End Function
Function GetAutoPostFlag(SmsType)
Dim rsas
set rsas=conn.execute("select * from smsautosend where smstype='"&SmsType&"'")
if not rsas.eof Then
If rsas("autopost")=True Then
GetAutoPostFlag = 1
Else
GetAutoPostFlag = 0
End if
Else
GetAutoPostFlag = -1
end if
rsas.close
set rsas=nothing
End Function
Function FormatPhonenumber(pn)
if len(pn)=12 and instr("013,015,018",left(pn,3))>0 then pn=mid(pn,2)
if (len(pn)=11 and instr("13,15,18",left(pn,2))>0) or (left(pn,1)="0" and (len(pn)=11 or len(pn)=12)) then
FormatPhonenumber=pn
else
FormatPhonenumber=""
end if
End Function
Function SMSAutoPost(SmsType,xmid,khid,username)
Dim rsas,rsxm,rskh,rsdx,rshistory
Dim kehu_id,lxpeple,lxpeple1,lxpeple2,telephone,telephone1,telephone2,js_khid
Dim Content,Delay
Dim SMS_id,SMS_sn,SMS_pw,SMS_SendTime
If khid = 0 then
Set rsxm=conn.execute("select kehu_id from shejixiadan where id="&xmid)
If Not rsxm.eof Then
kehu_id = rsxm("kehu_id")
Else
Exit Function
End If
rsxm.close
Set rsxm = Nothing
Else
kehu_id = khid
End If
Set rskh = conn.execute("select lxpeple,lxpeple2,telephone,telephone2,js_id from kehu where id="&kehu_id)
If Not rskh.eof Then
lxpeple1 = rskh("lxpeple")
lxpeple2 = rskh("lxpeple2")
telephone1 = rskh("telephone")
telephone2 = rskh("telephone2")
js_khid = rskh("js_id")
Else
Exit Function
End If
rskh.close
Set rskh=Nothing
'
If SmsType = "js" And CStr(js_khid)<>"" And CStr(js_khid)<>"0" Then
Set rskh = conn.execute("select telephone,telephone2 from kehu where id="&js_khid)
If Not rskh.eof Then
telephone1 = rskh("telephone")
telephone2 = rskh("telephone2")
End If
rskh.close
Set rskh=Nothing
End If
If lxpeple1 <> "" And Not IsNull(lxpeple1) Then lxpeple = lxpeple1
If lxpeple2 <> "" And Not IsNull(lxpeple2) Then lxpeple = lxpeple & " " & lxpeple2
lxpeple = Trim(lxpeple)
If telephone1 <> "" And Not IsNull(telephone1) Then telephone = telephone1
If telephone2 <> "" And Not IsNull(telephone2) Then telephone = telephone & "," & telephone2
If Left(telephone,1) = "," Then telephone = Mid(telephone,2)
Dim times,years,months,days,hours,minutes,seconds
Set rsas=conn.execute("select * from smsautosend where smstype='"&SmsType&"'")
If Not rsas.EOF Then
Content = rsas("Content")
Delay = rsas("Delay")
Content = Replace(Content, "%n", lxpeple)
Content = Replace(Content, "%a", username)
If Delay > 0 Then
times = DateAdd("n", Delay, Now())
years = CStr(Year(times))
months = Month(times)
days = Day(times)
hours = Hour(times)
minutes = Minute(times)
seconds = Second(times)
If months < 10 Then months = "0" & months
If days < 10 Then days = "0" & days
If hours < 10 Then hours = "0" & hours
If minutes < 10 Then minutes = "0" & minutes
If seconds < 10 Then seconds = "0" & seconds
SMS_SendTime = years & months & days & hours & minutes & seconds
End If
Else
Exit Function
End If
rsas.close
Set rsas = Nothing
Set rsdx = conn.execute("select * from duanxin where statu=1")
If Not rsdx.eof Then
SMS_id = rsdx("id")
SMS_sn = rsdx("zhuce_id")
SMS_pw = rsdx("pass_word")
Else
Exit Function
End If
'
'
Dim volumes,SMS_Object,re
Set SMS_Object = new SMS_Class
SMS_Object.SmsCompanyID = SMS_id
SMS_Object.Create()
re=SMS_Object.SendMessage(SMS_sn, SMS_pw, telephone, Content, SMS_SendTime, "")
If re = 1 Then
volumes=1
if instr(telephone,",")>0 then volumes=2
Set rshistory = Server.CreateObject("ADODB.RECORDSET")
rshistory.open "select top 1 * from SmsHistory",conn,1,3
rshistory.addnew
rshistory("UserName") = username
rshistory("SmsType") = 0
rshistory("SendTime") = Now()
rshistory("SmsVolume") = volumes
rshistory("TaskName") = SmsType
rshistory("Xiangmu_ID") = xmid
rshistory("Content1") = Content
rshistory.update
rshistory.close
Set rshistory = Nothing
End If
Set SMS_Object = Nothing
End Function
'
Function ProcessEnrolComplete(orderid)
Dim rs_orderlist
Set rs_orderlist = Server.CreateObject("ADODB.RECORDSET")
rs_orderlist.open "select * from ProcessEnrolDetails where orderid="&orderid,conn,1,1
If Not (rs_orderlist.eof And rs_orderlist.bof) Then
Call ProcessPro(rs_orderlist("ProID"),rs_orderlist("ProVol"),rs_orderlist("ProMemo"),0,rs_orderlist("ProType"))
End If
rs_orderlist.close
Set rs_orderlist = Nothing
Dim rs_order
Set rs_order = Server.CreateObject("ADODB.RECORDSET")
rs_order.open "select * from ProcessEnrolOrder where id="&orderid,conn,1,3
If Not (rs_order.eof And rs_order.bof) Then
rs_order("WcTime") = Now
rs_order("WcAdmin") = session("adminid")
rs_order.update
End If
rs_order.close
Set rs_order = Nothing
End Function
'
'
'
'
'
'
'
Function ProcessPro(ProID,ProVol,ProMemo,PType,ProType)
Dim ArrXiangmuID,ChkXmProExists,OrderID,ArrYunyong,ArrProVol
Dim ProJxVol,ProHqVol
Dim RsXiangmu,rsyjqj,RsVerify,RsFujia
OrderID = 0
ProJxVol = 0
ProHqVol = 0
'
'
'
'
'
if trim(ProMemo)<>"" then
ArrXiangmuID = Split(ProMemo,",")
For i = 0 to UBound(ArrXiangmuID)
'
'
'
ArrXiangmuID(i) = Trim(ArrXiangmuID(i))
If ArrXiangmuID(i)<>"" And IsNumeric(ArrXiangmuID(i)) Then
'
If ProType = 0 Then
Set RsXiangmu = Server.CreateObject("ADODB.RECORDSET")
RsXiangmu.open "select yunyong,sl from shejixiadan where instr(', '+yunyong+',', ', "&ProID&",')>0 and id="&ArrXiangmuID(i),conn,1,1
if Not (RsXiangmu.Eof And RsXiangmu.Bof) Then
ArrYunyong = Split(RsXiangmu("yunyong"),", ")
ArrProVol = Split(RsXiangmu("sl"),", ")
For k=0 To UBound(ArrYunyong)
If CInt(ArrYunyong(k))=CInt(ProID) And ProType = 0 Then
ProJxVol = ArrProVol(k)
Exit For
End If
Next
Set RsVerify = Server.CreateObject("ADODB.RECORDSET")
RsVerify.Open "select d.* from VerifyProDetails d inner join VerifyProList o on d.mainid=o.id where o.vType="&PType&" and o.Xiangmu_ID="&ArrXiangmuID(i)&" and d.proid="&ProID&" and d.ProType=0",conn,1,3
if RsVerify.eof and RsVerify.bof Then
If Trim(DataDict(CStr(ArrXiangmuID(i))))="" Or Not IsNumeric(DataDict(CStr(ArrXiangmuID(i))))  Then
'
DataDict(CStr(ArrXiangmuID(i))) = CreateNewOrder(ArrXiangmuID(i),PType)
end If
'
conn.execute("insert into VerifyProDetails (MainID,ProID,ProVol,Types,ProType) values ("&DataDict(CStr(ArrXiangmuID(i)))&","&ProID&","&ProJxVol&",0,0)")
Else
'
If ProJxVol <> RsVerify("ProVol") Then
RsVerify("ProVol") = ProJxVol
RsVerify.update
End If
end if
RsVerify.close
set RsVerify=Nothing
end If
RsXiangmu.close
Set RsXiangmu = Nothing
End If
If ProType = 1 Then
'
Set RsFujia = Server.CreateObject("ADODB.RECORDSET")
RsFujia.open "select fujia.* from fujia inner join yunyong on fujia.jixiang=yunyong.id where yunyong.type=1 and fujia.jixiang="&ProID&" and fujia.xiangmu_id="&ArrXiangmuID(i)&" order by times",conn,1,1
If Not (RsFujia.Eof And RsFujia.Bof) Then
ProHqVol = RsFujia("sl")
Set RsVerify = Server.CreateObject("ADODB.RECORDSET")
RsVerify.Open "select d.* from VerifyProDetails d inner join VerifyProList o on d.mainid=o.id where o.vType="&PType&" and o.Xiangmu_ID="&ArrXiangmuID(i)&" and d.proid="&ProID&" and d.ProType=1",conn,1,3
if RsVerify.eof and RsVerify.bof Then
If DataDict(CStr(ArrXiangmuID(i)))="" Or Not IsNumeric(DataDict(CStr(ArrXiangmuID(i))))  Then
'
DataDict(CStr(ArrXiangmuID(i))) = CreateNewOrder(ArrXiangmuID(i),PType)
end If
'
conn.execute("insert into VerifyProDetails (MainID,ProID,ProVol,Types,ProType) values ("&DataDict(CStr(ArrXiangmuID(i)))&","&ProID&","&ProHqVol&",1,1)")
Else
'
If ProHqVol <> RsVerify("ProVol") Then
RsVerify("ProVol") = ProHqVol
RsVerify.update
End If
end if
RsVerify.close
set RsVerify=Nothing
End If
RsFujia.close
Set RsFujia = Nothing
End If
End If
'
'
Next
end if
end Function
Function CheckEnrolOrders()
dim EnrolBackGroups(8)
dim StrProID,StrProMoney,StrProVol,StrProMemo,StrCompName,StrMemo,StrOther
Dim StrSourcePro,StrSourceMemo,ArrSourcePro,ArrSourceMemo,tmp_memo,tmp_arrmemo,tmp_flag,tmp_counter
dim ArrProID,ArrProMoney,ArrProVol,ArrProMemo
dim rsenrol,rsve
Dim rsorder,cc
cc=0
Set rsorder = server.CreateObject("adodb.recordset")
rsorder.open "select * from ProcessEnrolOrder where isnull(wctime) order by id desc",conn,1,3
Do While Not rsorder.eof
StrProID = ""
StrProMoney = ""
StrProVol = ""
StrProMemo = ""
StrCompName = ""
StrMemo = ""
StrOther = ""
StrSourcePro = ""
StrSourceMemo = ""
Set rsenrol = server.createobject("adodb.recordset")
rsenrol.open "SELECT o.OrderNo, o.CompID, o.LcTime, o.Memo, d.* FROM ProcessEnrolOrder o INNER JOIN ProcessEnrolDetails d ON o.ID = d.OrderID where d.orderid="&rsorder("id"),conn,1,1
do while not rsenrol.eof
tmp_flag = true
tmp_memo = Trim(rsenrol("ProMemo"))
If InStr(tmp_memo,",")>0 Then
tmp_arrmemo = Split(tmp_memo,",")
tmp_memo = ""
For tmp_counter = 0 To UBound(tmp_arrmemo)
tmp_arrmemo(tmp_counter) = Trim(tmp_arrmemo(tmp_counter))
If tmp_arrmemo(tmp_counter)<>"" And IsNumeric(tmp_arrmemo(tmp_counter)) Then
tmp_arrmemo(tmp_counter) = Trim(tmp_arrmemo(tmp_counter))
Set rsve = conn.execute("SELECT D.ProID, D.ProVol FROM VerifyProDetails D INNER JOIN VerifyProList L ON D.MainID = L.ID WHERE D.ProID="&rsenrol("ProID")&" AND L.vType=0 AND D.ProType="&rsenrol("ProType")&" AND L.Xiangmu_ID="&tmp_arrmemo(tmp_counter))
If rsve.eof And rsve.bof Then
tmp_memo = tmp_memo & "," & tmp_arrmemo(tmp_counter)
End If
rsve.Close
Set rsve = Nothing
End If
Next
If tmp_memo <> "" Then
tmp_memo = Mid(tmp_memo,2)
Else
tmp_flag = false
End If
Else
If tmp_memo<>"" And IsNumeric(tmp_memo) then
Set rsve = conn.execute("SELECT D.ProID, D.ProVol FROM VerifyProDetails D INNER JOIN VerifyProList L ON D.MainID = L.ID WHERE D.ProID="&rsenrol("ProID")&" AND L.vType=0 AND D.ProType="&rsenrol("ProType")&" AND L.Xiangmu_ID="&tmp_memo)
If Not (rsve.eof And rsve.bof) Then
tmp_flag = false
End If
rsve.Close
Set rsve = Nothing
End if
End If
If tmp_flag Then
StrProID = StrProID & "|" & rsenrol("ProID")
StrProMoney = StrProMoney & "|" & rsenrol("ProMoney")
StrProVol = StrProVol & "|" & rsenrol("ProVol")
StrProMemo = StrProMemo & "|" & tmp_memo
StrCompName = rsenrol("CompID")
StrMemo = rsenrol("Memo")
StrOther = orderid&", "&rsenrol("OrderNo")&", "&rsenrol("LcTime")&", "
StrSourcePro = StrSourcePro & "|" & rsenrol("ProID")
StrSourceMemo = StrSourceMeMo & "|" & tmp_memo
End If
rsenrol.movenext
loop
rsenrol.close
set rsenrol = Nothing
If StrProID = "" Then
cc=cc+1
'
rsorder("wctime")=Now()
rsorder("wcadmin")=rsorder("adminid")
rsorder.update
End If
rsorder.movenext
Loop
rsorder.close
Set rsorder = Nothing
End Function
Function CreateNewOrder(XmID,PType)
Dim rsyjqj,temp
set rsyjqj = server.CreateObject("adodb.recordset")
sql = "select top 1 * from VerifyProList"
rsyjqj.open sql,conn,1,3
rsyjqj.addnew
rsyjqj("vType")=cint(PType)
rsyjqj("Xiangmu_ID")=cint(XmID)
rsyjqj("AdminID")=session("adminid")
rsyjqj("iDate")=now()
rsyjqj("Memo")=""
rsyjqj.update
temp = rsyjqj.bookmark
rsyjqj.bookmark = temp
CreateNewOrder = rsyjqj("ID")
rsyjqj.close
Set rsyjqj=Nothing
End Function
Function getPrintString(khid)
dim companyname,companytel,tempstr,shopname
dim sid:sid=0
dim rskh
set rskh = conn.execute("select shopid from kehu where id="&khid)
if not (rskh.eof and rskh.bof) then
sid = rskh("shopid")
end if
rskh.close
set rskh = nothing
companyname = conn.execute("select companyname from sysconfig")(0)
companytel = conn.execute("select companytel from sysconfig")(0)
tempstr = companyname
if sid<>0 then
dim rsshop
set rsshop = conn.execute("select shopname from MultipleShopList where id="&sid)
if not (rsshop.eof and rsshop.bof) then
shopname = rsshop(0)
end if
rsshop.close
set rsshop = nothing
tempstr = tempstr & "-"&shopname
end if
if not isnull(companytel) and trim(companytel)<>"" then tempstr = tempstr & "&nbsp;&nbsp;电话"&companytel
getPrintString = tempstr
End Function
Function getWedsuitCost(xmid)
Dim rsxm,rsyy
Dim tmp_i,SumYunyongCost
Dim arr_yunyong,arr_sl
SumYunyongCost = 0
Set rsxm = conn.execute("select jixiang_money,yunyong,sl from shejixiadan where id="&xmid)
If Not (rsxm.eof And rsxm.bof) Then
If rsxm("yunyong")<>"" And Not IsNull(rsxm("yunyong")) Then
arr_yunyong = Split(rsxm("yunyong"),", ")
arr_sl = Split(rsxm("sl"),", ")
For tmp_i = 0 To UBound(arr_yunyong)
Set rsyy = conn.execute("select in_money from yunyong where [type]=1 and in_money<>0 and id="&arr_yunyong(tmp_i))
If Not (rsyy.eof And rsyy.bof) Then
SumYunyongCost = SumYunyongCost + (rsyy("in_money") * CInt(arr_sl(tmp_i)))
End If
rsyy.close
Set rsyy = Nothing
Next
getWedsuitCost = SumYunyongCost
Else
getWedsuitCost = 0
End if
Else
getWedsuitCost = 0
End If
End Function
Function GetAppellation(val, ischild)
dim CompanyType
CompanyType = Conn.Execute("select CompanyType from sysconfig")(0)
if CompanyType=0 then
ischild = false
Select Case val
Case 1
GetAppellation = "客人"
Case 2
GetAppellation = "客人"
Case 3
GetAppellation = "先生"
Case 4
GetAppellation = "女士"
End Select
Else
If IsNull(ischild) Then
GetAppellation = "客人"
Else
If Not ischild Then
GetAppellation = "家长"
Else
GetAppellation = "孩子"
End If
End If
end if
End Function
Function GetDutyName(lv)
dim CompanyType,strDuty
CompanyType = SystemConfig("CompanyType", 0)
if CompanyType=0 then
strDuty = GetFieldDataBySQL("select worktype from worktype where [level]="&lv,"str","未知")
else
Select Case lv
Case 5
strDuty = "引导/爱婴"
Case 14
strDuty = "引导助理"
Case Else
strDuty = GetFieldDataBySQL("select worktype from worktype where [level]="&lv,"str","未知")
End Select
end if
GetDutyName = strDuty
End Function
Function GetWorkName(sign)
dim CompanyType
CompanyType = SystemConfig("CompanyType", 0)
if CompanyType=0 then
GetWorkName = "化妆"
else
GetWorkName = "引导"
end if
End Function
Function SystemConfig(str, defval)
err.clear()
dim val
val = Conn.Execute("select "& str &" from sysconfig")(0)
if err.number<>0 then
val = defval
err.clear()
end if
SystemConfig = val
End Function
'
'
'
Function ShowMultipleShopSelect(types, defval, isshowtext)
If Not IsNull(defval) Then
If CStr(defval)<>"" And IsNumeric(defval) then
defval = CInt(defval)
Else
defval = null
End If
Else
defval = null
End If
dim rsshop
set rsshop=conn.execute("select * from MultipleShopList order by px")
if rsshop.eof and rsshop.bof then response.write "<span style='display:none'>"
If isshowtext Then response.write "连锁店: "
response.write "<select name='shopid' id='shopid'>"&vbcrlf
response.write "<option value=''>"
if types=0 then
response.write "全部"
ElseIf types=1 Or types=3 then
response.write "请选择..."
ElseIf types=2 then
response.write "全部分店"
end if
response.write "</option>"&vbcrlf&"<option value='0'"
If (Not isnull(defval) and defval=0) Or (rsshop.eof and rsshop.bof) then response.write " selected"
response.write ">总店</option>"&vbcrlf
While Not rsshop.eof
response.write "<option value='"&rsshop("id")&"'"
If Not isnull(defval) And defval=rsshop("id") Then response.write " selected"
response.write ">"&rsshop("shopname")&"</option>"&vbcrlf
rsshop.movenext
Wend
response.write "</select>"&vbcrlf
If rsshop.eof And rsshop.bof Then response.write "</span>"
rsshop.close
Set rsshop=Nothing
End Function
Function GetMultipleShopListValue()
if (session("level")<>1 or session("zhuguan")<>1) and session("level")<>6 and session("level")<>7 and session("level")<>10 then
GetMultipleShopListValue = session("UserShopID")
else
GetMultipleShopListValue = null
end if
End Function
Function CheckShareShopToSql()
If session("isshare")=0 Then CheckShareShopToSql = " and k.shopid=" & session("UserShopID")
End Function
Function GetMultipleShopName(sid)
if sid<>"" and isnumeric(sid) then
if sid=0 then
GetMultipleShopName = "总店"
else
dim rsshop
set rsshop=conn.execute("select shopname from MultipleShopList where id="&sid)
if not (rsshop.eof and rsshop.bof) then
GetMultipleShopName = rsshop("shopname")
Else
GetMultipleShopName = "未知分店"
End If
rsshop.close
Set rsshop = Nothing
End If
Else
GetMultipleShopName = ""
End If
End Function
Function GetMultiShopSql(msid,tn,tfn,yfn)
Dim res,arrfield
arrfield = array("userid","username","yuangong_id","id")
If Not isnull(msid) And msid<>"" Then
res = " and "
If tn<>"" Then res = res & tn & "."
res = res & arrfield(tfn) & " in (select "& arrfield(yfn) &" from yuangong where shopid="& msid &")"
Else
res = ""
End If
GetMultiShopSql = res
End Function
Function ShowUserSelect(el_name, userlv, fieldname, nulltext, defval, width, isdisabled)
dim rsshop,rsuser,rsworktype,tmp_sql,tmp_flag,ismultiexist
tmp_flag = false
ismultiexist = false
set rsshop = conn.execute("select * from MultipleShopList order by px")
response.write "<select name='"&el_name&"' id='"&el_name&"'"
if width>0 then response.write " style='width:"&width&"px'"
if isdisabled then response.write " disabled"
response.write "><option value=''>"&nulltext&"</option>"
'
if Not (rsshop.eof and rsshop.bof) then ismultiexist = true
if ismultiexist then response.write "<OPTGROUP LABEL='总店'>"
dim wt_sql
wt_sql = "select * from worktype where 1=1"
If userlv <> "" Then wt_sql = wt_sql & " and [level] in ("&userlv&")"
wt_sql = wt_sql & " order by [level] asc"
if not ismultiexist then
set rsworktype = conn.execute(wt_sql)
do while not rsworktype.eof
response.write "<OPTGROUP LABEL='"& GetDutyName(rsworktype("level")) &"'>"
tmp_sql = "select * from yuangong where shopid=0 and username<>''and isdisabled=0"
tmp_sql = tmp_sql & " and [level]="&rsworktype("level")&" order by id asc"
set rsuser = conn.execute(tmp_sql)
do while not rsuser.eof
If Trim(rsuser("peplename"))<>"" then
response.write "<option value='"&rsuser(fieldname)&"'"
If rsuser(fieldname) = defval Then response.write " selected"
response.write ">"&rsuser("peplename")&"</option>"
End If
rsuser.movenext
loop
rsuser.close
set rsuser = nothing
response.write "</OPTGROUP>"
rsworktype.movenext
loop
rsworktype.close
set rsworktype = nothing
Else
tmp_sql = "select * from yuangong where shopid=0 and username<>''and isdisabled=0"
If userlv <> "" Then tmp_sql = tmp_sql & " and [level] in ("&userlv&")"
tmp_sql = tmp_sql & " order by [level] asc"
set rsuser = conn.execute(tmp_sql)
do while not rsuser.eof
If Trim(rsuser("peplename"))<>"" then
response.write "<option value='"&rsuser(fieldname)&"'"
If rsuser(fieldname) = defval Then
response.write " selected"
tmp_flag = True
End If
response.write ">"&rsuser("peplename")&"</option>"
End If
rsuser.movenext
Loop
If Not tmp_flag And Not IsNull(defval) And defval<>"" Then
response.write "<option value='' selected>"&defval&"</option>"
End If
if Not (rsshop.eof and rsshop.bof) then response.write "</OPTGROUP>"
rsuser.close
end if
'
Do While Not rsshop.eof
response.write "<OPTGROUP LABEL='"&rsshop("shopname")&"'>"
tmp_sql = "select * from yuangong where shopid="&rsshop("id")&" and username<>''"
If userlv <> "" Then
tmp_sql = tmp_sql & " and [level] in ("&userlv&")"
End If
tmp_sql = tmp_sql & " and isdisabled=0 order by [level],id"
set rsuser = conn.execute(tmp_sql)
do while not rsuser.eof
If Trim(rsuser("peplename"))<>"" then
response.write "<option value='"&rsuser(fieldname)&"'"
If rsuser(fieldname) = defval Then response.write " selected"
response.write ">"&rsuser("peplename")&"</option>"
End If
rsuser.movenext
Loop
response.write "</OPTGROUP>"
rsshop.movenext
Loop
response.write "</select>"
rsshop.close
Set rsshop = Nothing
End Function
Function ShowWedSignInput(prefix, xmid, pname, isreadonly)
dim rstype,sqlhs,slhs,stringbuilder
stringbuilder=""
set rstype=server.createobject("adodb.recordset")
sqlhs = "select * from hs_signtype order by px asc"
rstype.open sqlhs,conn,1,1
do while not rstype.eof
stringbuilder = stringbuilder & rstype("title") & "&nbsp;" & "<input type='text' name='"&prefix&rstype("id")&"' size='3'"
if not isnull(pname) and pname<>"" then
slhs=GetFieldDataBySQL("SELECT hs_signhistory.vol FROM hs_signhistory INNER JOIN yuangong ON hs_signhistory.userid = yuangong.ID where yuangong.peplename='"&pname&"' and hs_signhistory.xiangmu_id="&xmid&" and hs_signhistory.typeid="& rstype("id"),"int",0)
stringbuilder = stringbuilder & " value='"&slhs&"'"
if isreadonly then stringbuilder = stringbuilder & " readonly"
end if
stringbuilder = stringbuilder & " />&nbsp;&nbsp;&nbsp;"
rstype.movenext
loop
rstype.close
set rstype = Nothing
ShowWedSignInput = stringbuilder
End Function
Function GetCheckMoneyInfo(starttime, endtime, shopid)
dim rsmoney,sqlmoney,stringbuilder,sqlshop
stringbuilder=""
if shopid<>"" and isnumeric(shopid) then sqlshop=" and k.shopid="&shopid
dim tmp_userid,tmp_money,cur_userid
tmp_userid=0
dim rstest
set rstest=server.CreateObject("adodb.recordset")
rstest.open "select * from (select m.money as smoney,m.id as counts,m.checkuserid from (save_money m INNER JOIN shejixiadan s ON m.xiangmu_id = s.ID) INNER JOIN kehu k ON s.kehu_id = k.ID where k.isdelete=false and s.isdelete=false and m.isdelete=false and not isnull(m.times) and datevalue(m.times)>=#"&starttime&"# and datevalue(m.times)<=#"&endtime&"# and not isnull(m.times)"&sqlshop&" and ischeck=1     union     SELECT smoney,counts,checkuserid FROM (select bm_id,counts,userid,kehu_name,gtype,sum(savemoney) as smoney,ischeck,wzsk,checkuserid,times,beizhu from goumai_jilu where not isnull(times) and datevalue(times)>=#"&starttime&"# and datevalue(times)<=#"&endtime&"# and not isnull(times) group by bm_id,counts,userid,kehu_name,gtype,ischeck,wzsk,checkuserid,times,beizhu union all SELECT goumai_jilu.bm_id, goumai_jilu.counts, goumai_jilu.userid, goumai_jilu.kehu_name, goumai_jilu.gtype, goumai_jilu_rep.money as smoney, goumai_jilu.ischeck,goumai_jilu.wzsk, goumai_jilu.checkuserid, goumai_jilu_rep.dateandtime as times,goumai_jilu_rep.memo as beizhu FROM goumai_jilu INNER JOIN goumai_jilu_rep ON goumai_jilu.counts = goumai_jilu_rep.counts_id where not isnull(goumai_jilu_rep.dateandtime) and datevalue(goumai_jilu_rep.dateandtime)>=#"&starttime&"# and datevalue(goumai_jilu_rep.dateandtime)<=#"&endtime&"# and not isnull(goumai_jilu_rep.dateandtime) and goumai_jilu.ischeck=1) ORDER BY counts   ) order by checkuserid",conn,1,1
do while not rstest.eof
if tmp_userid<>rstest("checkuserid") then
if tmp_userid<>0 then stringbuilder=stringbuilder &"，"& getfielddatabysql("select peplename from yuangong where id="&tmp_userid,"str","N/A") &"："& tmp_money &"元"
tmp_userid=rstest("checkuserid")
tmp_money=rstest("smoney")
else
tmp_money=tmp_money+rstest("smoney")
end if
cur_userid=rstest("checkuserid")
rstest.movenext
if rstest.eof and cur_userid=tmp_userid then
stringbuilder=stringbuilder &"，"& getfielddatabysql("select peplename from yuangong where id="&tmp_userid,"str","N/A") &"："& tmp_money &"元"
end if
loop
rstest.close
if stringbuilder<>"" then stringbuilder=mid(stringbuilder,2)
GetCheckMoneyInfo = stringbuilder
End Function
Function GetFieldDataBySQL(sql,fieldtype,defval)
dim rsnon
set rsnon = server.CreateObject("adodb.recordset")
rsnon.open sql,conn,1,1
if not (rsnon.eof and rsnon.bof) then
GetFieldDataBySQL = rsnon(0)
else
GetFieldDataBySQL = defval
end if
rsnon.close
set rsnon = nothing
End Function
Function GetProjectReportSQL(uid,un)
Dim user_field(1),user_val(1),user_sql(1)
user_field(0) = "userid,userid2,userid3,kj_userid,hz_userid,hz_userid2,hz_userid3,hs_userid"
user_field(1) = "hz_name,cp_name,cp_name2,cp_name3,cp_name4,cp_name5,xp_name,xp2_name,ky_name,sj_name,wc_name,xg_name,zlname,hz_name2,hz_name3,cpzl_name,cpzl_name2,cpzl_name3,cpzl_name4,cpzl_name5,ky_name2"
user_val(0) = uid
user_val(1) = un
Dim arr_uf,ci,fi
For ci = 0 To UBound(user_field)
arr_uf = Split(user_field(ci),",")
For fi = 0 To UBound(arr_uf)
user_sql(ci) = user_sql(ci) & "&','&s." & arr_uf(fi)
Next
If user_sql(ci) <> "" Then user_sql(ci) = Mid(user_sql(ci),6)
Next
Dim tmpsql
tmpsql = "select top 8 b.* from sjs_baobiao b inner join shejixiadan s on b.xiangmu_id=s.id where b.userid='"&session("userid")&"' or b.topeple='"&session("userid")&"' or "& user_sql(0) &" like '%,"& user_val(0) &",%' or "& user_sql(1) &" like '%,"& user_val(1) &",%' and b.eventid=0 order by b.times desc"
GetProjectReportSQL = tmpsql
End Function
Function CheckOldMoneyControl()
dim OldMoneyControl,level2
OldMoneyControl = GetFieldDataBySQL("select OldMoneyControl from sysconfig","int",0)
If session("adminid") = "" Or IsNull(session("adminid")) Then
if OldMoneyControl = 0 then
CheckOldMoneyControl = true
else
CheckOldMoneyControl = False
end if
Else
level2 = GetFieldDataBySQL("select level2 from yuangong where id="&session("adminid"),"str","")
if OldMoneyControl=0 or session("level")=10 then
CheckOldMoneyControl = true
else
if instr(level2,"723")>0 then
CheckOldMoneyControl = true
else
CheckOldMoneyControl = false
end if
end If
End If
End Function

function GetProFirstPhoto(proid)
if cstr(proid)="" or not isnumeric(proid) then GetProFirstPhoto=null
dim rspro,pic
set rspro = conn.execute("select pic from yunyong where id="&proid)
if not (rspro.eof and rspro.bof) then
if trim(rspro("pic"))<>"" and not isnull(rspro("pic"))="" then
if CheckFileIsExist("../upload/",rspro("pic")) then
GetProFirstPhoto = rspro("pic")
exit function
end if
end if
end if
rspro.close
set rspro = nothing
dim rspic
set rspic = conn.execute("select filename from yunyong_pic where proid="&proid&" order by px")
if not (rspic.eof and rspic.bof) then
do while not rspic.eof
if trim(rspic("filename"))<>"" and not isnull(rspic("filename")) then
if CheckFileIsExist("../upload/",rspic("filename")) then
GetProFirstPhoto = rspic("filename")
rspic.close
set rspic = nothing
exit function
end if
end if
rspic.movenext
loop
end if
rspic.close
set rspic = nothing
GetProFirstPhoto = null
end function
function CheckFileIsExist(filepath,filename)
dim FSO,pic
set FSO=server.createobject("scripting.filesystemobject")
pic=server.mappath(filepath&filename)
if FSO.FileExists(pic) then
CheckFileIsExist = true
else
CheckFileIsExist = false
end if
set FSO=nothing
end Function
Function GetProductCosting(yunyong_id, yunyong_sl)
Dim tmp_cost
Dim rscb
Set rscb = Server.CreateObject("ADODB.RECORDSET")
rscb.open "select in_money from yunyong where id=" & yunyong_id, conn, 1, 1
If Not (rscb.eof And rscb.bof) Then
tmp_cost = rscb("in_money") * yunyong_sl
Else
tmp_cost = 0
End If
rscb.close
Set rscb = Nothing
GetProductCosting = tmp_cost
End Function
Function GetCostCalcuation(StartTime, EndTime, UserID, XiangmuNotNullFieldList, XiangmuIdList, ViewType, SplitField)
GetCostCalcuation = GetCostCalcuationForShop(StartTime, EndTime, UserID, XiangmuNotNullFieldList, XiangmuIdList, ViewType, SplitField, "")
End Function
Function GetCostCalcuationForShop(StartTime, EndTime, UserID, XiangmuNotNullFieldList, XiangmuIdList, ViewType, SplitField, ShopID)
Dim rscc,sqlcc,sqlshop
Set rscc = Server.CreateObject("ADODB.RECORDSET")
'
'
Dim m,n
Dim arr_temp(4,1)
For m = 0 To 4
For n = 0 To 1
arr_temp(m,n) = 0
Next
Next
If ShopID<>"" And Not IsNull(ShopID) And IsNumeric(ShopID) Then
sqlshop = " and kehu.shopid=" & ShopID
End If
Dim arr_notnullfield, sql_notnull,q
If XiangmuNotNullFieldList <> "" Then
arr_notnullfield = Split(XiangmuNotNullFieldList,",")
For q = 0 To UBound(arr_notnullfield)
sql_notnull = sql_notnull & " and not isnull(shejixiadan."& arr_notnullfield(q) &")"
Next
End If
If XiangmuIdList <>"" Then
If Left(XiangmuIdList,1)="," Then XiangmuIdList = Mid(XiangmuIdList,2)
If Right(XiangmuIdList,1)="," Then XiangmuIdList = Left(XiangmuIdList,Len(XiangmuIdList)-1)
If XiangmuIdList <>"" Then sql_notnull = sql_notnull & " and shejixiadan.id in ("& XiangmuIdList &")"
End If
Dim arr_splitfield,splitcount,sql_splitfield,r,sql_usercheck,sql_kyusercheck,peplename
If SplitField <> "" Then
If UserID <> "" Then
peplename = GetFieldDataBySQL("select peplename from yuangong where username='"&UserID&"'","str","")
End If
arr_splitfield = Split(SplitField, ",")
For r = 0 To UBound(arr_splitfield)
If arr_splitfield(r)<>"" Then
sql_splitfield = sql_splitfield & ",shejixiadan." & arr_splitfield(r)
If UserID <> "" Then
sql_usercheck = sql_usercheck & " or shejixiadan." & arr_splitfield(r) & "='"& UserID &"'"
End If
If peplename <> "" Then
sql_kyusercheck = sql_kyusercheck & " or shejixiadan." & arr_splitfield(r) & "='"& peplename &"'"
End If
End If
Next
If sql_usercheck<>"" Then
sql_usercheck = Mid(sql_usercheck,5)
sql_usercheck = " and ("&sql_usercheck&")"
End If
If sql_kyusercheck<>"" Then
sql_kyusercheck = Mid(sql_kyusercheck,5)
sql_kyusercheck = " and ("&sql_kyusercheck&")"
End If
End If
'
If ViewType="" Or InStr(ViewType, "0") Then
sqlcc = "select shejixiadan.* from shejixiadan inner join kehu on shejixiadan.kehu_id=kehu.id where not isnull(shejixiadan.times) and datevalue(shejixiadan.times)>=#"&starttime&"# and datevalue(shejixiadan.times)<=#"&endtime&"# and not isnull(shejixiadan.times)" & GetMultiShopSql(defshopvalue,"shejixiadan",0,1)
sqlcc = sqlcc & sql_notnull & sql_usercheck
rscc.open sqlcc,conn,1,1
Do While Not rscc.EOF
splitcount = 0
If IsArray(arr_splitfield) then
For r = 0 To UBound(arr_splitfield)
If Not IsNull(rscc(arr_splitfield(r))) And rscc(arr_splitfield(r))<>"" Then
splitcount = splitcount + 1
End If
Next
Else
splitcount = 1
End If
arr_temp(0,0) = arr_temp(0,0) + rscc("jixiang_money")/splitcount
If Not IsNull(rscc("yunyong")) And rscc("yunyong")<>"" Then
Dim arr_yy,arr_sl,p
arr_yy = Split(rscc("yunyong"),", ")
arr_sl = Split(rscc("sl"),", ")
For p = 0 To UBound(arr_yy)
arr_temp(0,1) = arr_temp(0,1) + GetProductCosting(arr_yy(p), CSng(arr_sl(p))) / splitcount
Next
End If
rscc.movenext
Loop
rscc.close
End If
'
If ViewType="" Or InStr(ViewType, "1") Then
sqlcc = "select fujia.jixiang,fujia.sl,fujia.money" & sql_splitfield & " from (fujia INNER JOIN shejixiadan ON fujia.xiangmu_id = shejixiadan.ID) INNER JOIN kehu ON shejixiadan.kehu_id = kehu.ID where not isnull(fujia.times) and datevalue(fujia.times)>=#"&starttime&"# and datevalue(fujia.times)<=#"&endtime&"# and not isnull(fujia.times)" & GetMultiShopSql(defshopvalue,"fujia",0,1)
sqlcc = sqlcc & sql_notnull & sql_kyusercheck
rscc.open sqlcc,conn,1,1
Do While Not rscc.EOF
splitcount = 0
If IsArray(arr_splitfield) then
For r = 0 To UBound(arr_splitfield)
If Not IsNull(rscc(arr_splitfield(r))) And rscc(arr_splitfield(r))<>"" Then
splitcount = splitcount + 1
End If
Next
Else
splitcount = 1
End If
arr_temp(1,0) = arr_temp(1,0) + rscc("money") / splitcount
arr_temp(1,1) = arr_temp(1,1) + GetProductCosting(rscc("jixiang"), rscc("sl")) / splitcount
rscc.movenext
Loop
rscc.close
End If
'
If ViewType="" Or InStr(ViewType, "2") Then
sqlcc = "select fujia2.jixiang,fujia2.sl,fujia2.money" & sql_splitfield & " from (fujia2 INNER JOIN shejixiadan ON fujia2.xiangmu_id = shejixiadan.ID) INNER JOIN kehu ON shejixiadan.kehu_id = kehu.ID where not isnull(fujia2.times) and datevalue(fujia2.times)>=#"&starttime&"# and datevalue(fujia2.times)<=#"&endtime&"# and not isnull(fujia2.times)" & GetMultiShopSql(defshopvalue,"fujia2",0,1)
sqlcc = sqlcc & sql_notnull
rscc.open sqlcc,conn,1,1
Do While Not rscc.EOF
splitcount = 0
If IsArray(arr_splitfield) then
For r = 0 To UBound(arr_splitfield)
If Not IsNull(rscc(arr_splitfield(r))) And rscc(arr_splitfield(r))<>"" Then
splitcount = splitcount + 1
End If
Next
Else
splitcount = 1
End If
arr_temp(2,0) = arr_temp(2,0) + rscc("money") / splitcount
arr_temp(2,1) = arr_temp(2,1) + GetProductCosting(rscc("jixiang"), rscc("sl")) / splitcount
rscc.movenext
Loop
rscc.close
End If
'
If ViewType="" Or InStr(ViewType, "3") Then
sqlcc = "select goumai.jixiang,goumai.sl,goumai.money" & sql_splitfield & " from (goumai INNER JOIN shejixiadan ON goumai.xiangmu_id = shejixiadan.ID) INNER JOIN kehu ON shejixiadan.kehu_id = kehu.ID where not isnull(goumai.times) and datevalue(goumai.times)>=#"&starttime&"# and datevalue(goumai.times)<=#"&endtime&"# and not isnull(goumai.times)" & GetMultiShopSql(defshopvalue,"goumai",0,1)
sqlcc = sqlcc & sql_notnull
rscc.open sqlcc,conn,1,1
Do While Not rscc.EOF
splitcount = 0
If IsArray(arr_splitfield) then
For r = 0 To UBound(arr_splitfield)
If Not IsNull(rscc(arr_splitfield(r))) And rscc(arr_splitfield(r))<>"" Then
splitcount = splitcount + 1
End If
Next
Else
splitcount = 1
End If
arr_temp(3,0) = arr_temp(3,0) + rscc("money") / splitcount
arr_temp(3,1) = arr_temp(3,1) + GetProductCosting(rscc("jixiang"), rscc("sl")) / splitcount
rscc.movenext
Loop
rscc.close
End If
Set rscc = Nothing
For m = 0 To 3
For n = 0 To 1
arr_temp(4,n) = arr_temp(4,n) + arr_temp(m,n)
Next
Next
GetCostCalcuationForShop = arr_temp
End Function
Function GetDataAnalyzeState()
dim state
state = GetFieldDataBySQL("select DataAnalyzeInvis from sysconfig","int","0")
if session("level")=10 and state=2 then state=0
GetDataAnalyzeState = state
End Function
Function GetScattereArrearage(id)
Dim rsgj
Dim result, gjmoney, djmoney, repmoney
result = 0
If id="" Or Not IsNumeric(id) Then
GetScattereArrearage = 0
Exit Function
End If
'
gjmoney = GetFieldDataBySQL("select sum(money) from goumai_jilu where counts="&id,"int","0")
If IsNull(gjmoney) Then gjmoney = 0
'
djmoney = GetFieldDataBySQL("select top 1 savemoney from goumai_jilu where counts="&id&" order by id","int","0")
If IsNull(djmoney) Then djmoney = 0
'
repmoney = GetFieldDataBySQL("select sum(money) from goumai_jilu_rep where counts_id="&id,"int","0")
If IsNull(repmoney) Then repmoney = 0
'
result = gjmoney - djmoney - repmoney
GetScattereArrearage = result
End Function
Function UpdateScattereArrearage(id)
conn.execute("update goumai_jilu set qkmoney="& GetScattereArrearage(id) &" where counts="& id)
End Function
Function CheckUserPermission(PermissionID)
If InStr(", " & session("level2") & ", ", ", " & PermissionID & ", ") > 0 Or InStr(", " & session("zg_level2") & ", ", ", " & PermissionID & ", ") > 0 Or session("level") = 10 Then
CheckUserPermission = True
Else
CheckUserPermission = False
End If
End Function
function GetDatesInMonth(y,m)
if m=12 then
n_y = y+1
n_m = 1
else
n_y = y
n_m = m+1
end if
GetDatesInMonth = datediff("d",cdate(y&"-"&m&"-1"),cdate(n_y&"-"&n_m&"-1"))
end function
function GetWorkOrders(fieldname,times)
GetWorkOrders = GetWorkOrders2(fieldname,times,times)
end Function
function GetWorkOrders2(fieldname,starttime,endtime)
GetWorkOrders2 = GetWorkOrders3(fieldname,starttime,endtime,"",0)
end Function
function GetWorkOrders3(fieldname,starttime,endtime,colorfield,colorid)
Dim sqlstr,sqlcolor
If colorfield<>"" Then sqlcolor = " and s."&colorfield&"="&colorid
If fieldname="pz_time" then
sqlstr = "select count(0) from (SELECT s.ID FROM shejixiadan s LEFT JOIN kehu k ON s.kehu_id = k.ID where not isnull(pz_time) and s.pz_time>=#"&starttime&"# and s.pz_time<=#"&endtime&"# and not isnull(pz_time)"&sqlcolor&" union all SELECT s.ID FROM shejixiadan s LEFT JOIN kehu k ON s.kehu_id = k.ID where not isnull(pz_time2) and s.pz_time2>=#"&starttime&"# and s.pz_time2<=#"&endtime&"# and not isnull(pz_time2)"&sqlcolor&" union all SELECT s.ID FROM shejixiadan s LEFT JOIN kehu k ON s.kehu_id = k.ID where not isnull(pz_time3) and s.pz_time3>=#"&starttime&"# and s.pz_time3<=#"&endtime&"# and not isnull(pz_time3)"&sqlcolor&")"
Else
sqlstr = "select count(0) from shejixiadan s where not isnull("&fieldname&") and datevalue("&fieldname&")>=#"&starttime&"# and datevalue("&fieldname&")<=#"&endtime&"# and not isnull("&fieldname&")"&sqlcolor
End If
tmp = conn.execute(sqlstr)(0)
if isnull(tmp) then tmp=0
GetWorkOrders3 = tmp
end Function
Sub ShowDatecolorSelector(elname, colorid)
dim color
color = GetFieldDataBySQL("select color from sys_datecolor where id="& colorid,"str","#ffffff")
dim rslist,str
str = "<select name='"& elname &"' id='"& elname &"' style='width:40px'>"&vbcrlf
str = str & "<option value='0' style='background-color:#ffffff'></option>"&vbcrlf
set rslist = server.createobject("adodb.recordset")
rslist.open "select * from sys_datecolor order by px",conn,1,1
do while not rslist.eof
str = str & "<option value='"& rslist("id") &"'"
if rslist("desc")<>"" then
str = str & " title='"& rslist("desc") &"'"
end if
if colorid=rslist("id") then str = str & " selected"
str = str & " style='background-color:"& rslist("color") &"'></option>"&vbcrlf
rslist.movenext
loop
rslist.close
set rslist= nothing
str = str & "</select>"&vbcrlf
response.write str
End Sub
Function GetDatecolor(colorid, defvalue)
if defvalue="" or isnull(defvalue) then defvalue="#ffffff"
if colorid="" or isnull(colorid) then colorid=0
GetDatecolor = GetFieldDataBySQL("select color from sys_datecolor where id="& colorid,"str",defvalue)
End Function
Sub ShowDatecolorList(datafieldname,colorfieldname,starttime,endtime)
dim rslist,str,tmpsl
str = "<b>"&GetWorkOrders3(datafieldname,starttime,endtime,"",0)&"</b>&nbsp;"
set rslist = server.createobject("adodb.recordset")
rslist.open "select * from sys_datecolor order by px",conn,1,1
do while not rslist.eof
tmpsl=GetWorkOrders3(datafieldname,starttime,endtime,colorfieldname,rslist("id"))
If tmpsl>0 Then str = str & "<span style='width:13px; padding:2px; font-size:10px; line-height:10px;background-color:"&rslist("color")&"'>"&tmpsl&"</span>&nbsp;"
rslist.movenext
Loop
rslist.close
Set rslist = nothing
response.write str
End Sub
Function ShowPaixiuInfo(dates)
dim rspx,result
dim i_menshi,i_huazhuang,i_shuma,i_sheying,i_houqin
set rspx = server.createobject("adodb.recordset")
rspx.open "select top 1 * from sys_paixiu where idate=#"& dates &"#",conn,1,1
if not (rspx.eof and rspx.bof) then
i_menshi = rspx("m_menshi")
i_huazhuang = rspx("m_huazhuang")
i_shuma = rspx("m_shuma")
i_sheying = rspx("m_sheying")
i_houqin = rspx("m_houqin")
end if
rspx.close
set rspx = nothing
if i_menshi & i_huazhuang & i_shuma & i_sheying & i_houqin <> "" then
result = result & "<div id=""cont_px"">"
result = result & "	<span class=""title"">今日员工排休信息</span>"
result = result & "	<ul>"
if i_menshi<>"" 	then result = result & "<li><b>门市部：</b>"&i_menshi&"</li>"
if i_sheying<>"" 	then result = result & "<li><b>摄影部：</b>"&i_sheying&"</li>"
if i_huazhuang<>"" 	then result = result & "<li><b>化妆部：</b>"&i_huazhuang&"</li>"
if i_shuma<>"" 		then result = result & "<li><b>数码部：</b>"&i_shuma&"</li>"
if i_houqin<>"" 	then result = result & "<li><b>后勤部：</b>"&i_houqin&"</li>"
result = result & "	</ul>"
result = result & "</div>"
end if
ShowPaixiuInfo = result
End Function
Sub ShowElementForChildSystem(types, elname, elvalue, defvalue, spanname, isdisabled, ischeckver)
Dim sysver, strhtml
sysver = GetFieldDataBySQL("select CompanyType from sysconfig","int",0)
If sysver = 1 Or Not ischeckver Then
Select Case types
Case "CheckIsChild"
strhtml = strhtml & "<input type=""checkbox"" name="""& elname &""" id="""& elname &""""
If Not IsNull(elvalue) Then strhtml = strhtml & " value="""& elvalue &""""
If defvalue Then strhtml = strhtml & " checked"
If isdisabled Then strhtml = strhtml & " disabled"
If spanname<>"" then strhtml = strhtml & " onclick=""javascript:changeContText('"& spanname &"', this.checked, '家长', '孩子');"""
strhtml = strhtml & "><label for="""& elname &""">&nbsp;宝贝</label>"
End Select
If isdisabled Then
strhtml = strhtml & "<input type=""hidden"" name="""& elname &""" id="""& elname &""" value="""& elvalue &""">"
End If
Response.Write strhtml
End If
End Sub
%><html><head><meta http-equiv="Content-Type" content="text/html; charset=gb2312"><title>无标题文档</title><script language="javascript" src="inc/func.js" type="text/javascript"></script></head><body><br><br><br><br><br><%
set rs4=server.CreateObject("adodb.recordset")
rs4.open "select * from richeng where peplename='"&session("username")&"' and  times=#"&date&"#",conn,1,1
if rs4.eof then
response.Write "<script>alert('请先添加今日日程,并回复日程!');location='tjrc.asp'</script>"
else
while not rs4.eof
if isnull(rs4("huifu")) then
response.Write "<script>alert('请先回复所有今日日程!');history.go(-1)</script>"
end if
rs4.movenext
wend
rs4.close
set rs4=nothing
end if
set rs2=server.CreateObject("adodb.recordset")
rs2.open "select * from zhichu where yewuyuan='"&session("username")&"' and  times=#"&date&"#",conn,1,3
if rs2.eof then
%><table width="607" border="0" align="center" cellpadding="0" cellspacing="0"><form action="" method="post" name="form1"><tr bordercolor="#CC9999" bgcolor="#99CCFF"><td height="35" colspan="2"><div align="center">日程费用支出录入</div></td></tr><%
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from richeng where peplename='"&session("username")&"' and times=#"&date&"# order by shijian asc",conn,1,1
while not rs.eof
y=y+1
%><tr bordercolor="#CC9999" bgcolor="#99CCFF"><td width="114" height="21"><div align="right">日程<%= y %>:</div></td><td width="493"><%= rs("title") %></td></tr><%
rs.movenext
wend
rs.close
set rs=nothing
%><tr bordercolor="#CC9999" bgcolor="#9999FF"><td height="18">&nbsp;</td><td>&nbsp;</td></tr><tr bordercolor="#CC9999" bgcolor="#99CCFF"><td height="25"><div align="right">支出总金额:</div></td><td><input name="money" type="text" id="money" size="7">
    元</td></tr><tr bordercolor="#CC9999" bgcolor="#99CCFF"><td height="109" valign="top"><div align="right">支出说明:</div></td><td><textarea name="shuoming" cols="65" rows="9" id="shuoming"><%
set rs3=server.CreateObject("adodb.recordset")
rs3.open "select distinct city from kehu where companyname in (select company from richeng where peplename='"&session("username")&"' and times=#"&date&"#)",conn,1,1
while not rs3.eof
if rs3("city")="泉州" then
response.Write "日照市公交费(&nbsp;&nbsp;)元"&CHR(13)
else
response.Write "公司--->"&rs3("city")&"&nbsp;&nbsp;&nbsp;班车费(&nbsp;&nbsp;)元&nbsp;&nbsp;公交费(&nbsp;&nbsp;)元"&CHR(13)
end if
rs3.movenext
wend
rs3.close
set rs3=nothing
%></textarea></td><tr bordercolor="#CC9999" bgcolor="#99CCFF"><td height="47" colspan="2"><div align="center"><input name="tijiao" type="submit" id="tijiao" value="提交">
&nbsp;&nbsp;&nbsp;&nbsp;
  <input type="button" name="Submit" value="返回" onClick="javascript:history.go(-1)"></div></td></form></table><%
else
%><table width="607" height="254" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#CC9999" bgcolor="#99CCFF"><form action="" method="post" name="form1"><tr><td height="29" colspan="2"><div align="center">日程费用支出录入</div></td></tr><%
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from richeng where peplename='"&session("username")&"' and  times=#"&date&"# order by shijian asc",conn,1,1
while not rs.eof
y=y+1
%><tr><td width="109" height="18"><div align="right">日程<%= y %>:</div></td><td width="492"><%= rs("title") %></td></tr><%
rs.movenext
wend
rs.close
set rs=nothing
%><tr bgcolor="#9999FF"><td height="18">&nbsp;</td><td>&nbsp;</td></tr><tr><td width="109" height="31"><div align="right">支出金额:</div></td><td width="492"><input name="money" type="text" id="money" size="7" value="<%= rs2("money") %>">
        元</td></tr><tr><td height="109" valign="top"><div align="right">支出说明:</div></td><td><textarea name="shuoming" cols="65" rows="9" id="shuoming"><%= encode2(rs2("shuoming")) %></textarea></td></tr><tr><td height="38" colspan="2"><div align="center"><input name="tijiao" type="submit" id="tijiao" value="提交">
&nbsp;&nbsp;&nbsp;&nbsp;
          <input type="button" name="Submit" value="返回" onClick="javascript:history.go(-1)"></div></td></tr></form></table><%
end if
if request("tijiao")="提交"then
if not isnumeric(request("money"))then
response.Write "<script>alert('支出费用只能是数字!');history.go(-1)</script>"
On Error GoTo 0
Err.Raise 9999
end if
if  request("shuoming")="" then
response.Write "<script>alert('请填写支出说明!');history.go(-1)</script>"
On Error GoTo 0
Err.Raise 9999
end if
set rs1=server.CreateObject("adodb.recordset")
rs1.open "select * from zhichu where yewuyuan='"&session("username")&"' and times=#"&date&"#",conn,1,3
if rs1.eof then
rs1.addnew
rs1("money")=request("money")
rs1("shuoming")=htmlencode2(request("shuoming"))
rs1("yewuyuan")=session("username")
rs1("times")=date
rs1.update
else
rs1("yewuyuan")=session("username")
rs1("money")=request("money")
rs1("shuoming")=htmlencode2(request("shuoming"))
rs1.update
end if
rs1.close
set rs1=nothing
response.Write "<script>alert('添加成功!');location='zhichu_look.asp'</script>"
end if
%></body></html>