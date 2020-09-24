VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMD5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Feb 21 2002 UMG
'
'   Prototype
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
DefLng A-Z 'we're 32 bit!

Private Const IniPart1  As Currency = 1732584193@
Private Const IniPart2  As Currency = 4023233417@
Private Const IniPart3  As Currency = 2562383102@
Private Const IniPart4  As Currency = 271733878@
Private Const Maxlong   As Long = 2 ^ 31 - 1

'Pass 1
Private Const S101      As Long = 7
Private Const S102      As Long = 12
Private Const S103      As Long = 17
Private Const S104      As Long = 22

Private Const A101      As Long = -680876936
Private Const A102      As Long = -389564586
Private Const A103      As Long = 606105819
Private Const A104      As Long = -1044525330
Private Const A105      As Long = -176418897
Private Const A106      As Long = 1200080426
Private Const A107      As Long = -1473231341
Private Const A108      As Long = -45705983
Private Const A109      As Long = 1770035416
Private Const A110      As Long = -1958414417
Private Const A111      As Long = -42063
Private Const A112      As Long = -1990404162
Private Const A113      As Long = 1804603682
Private Const A114      As Long = -40341101
Private Const A115      As Long = -1502002290
Private Const A116      As Long = 1236535329

'Pass 2
Private Const S201      As Long = 5
Private Const S202      As Long = 9
Private Const S203      As Long = 14
Private Const S204      As Long = 20

Private Const A201      As Long = -165796510
Private Const A202      As Long = -1069501632
Private Const A203      As Long = 643717713
Private Const A204      As Long = -373897302
Private Const A205      As Long = -701558691
Private Const A206      As Long = 38016083
Private Const A207      As Long = -660478335
Private Const A208      As Long = -405537848
Private Const A209      As Long = 568446438
Private Const A210      As Long = -1019803690
Private Const A211      As Long = -187363961
Private Const A212      As Long = 1163531501
Private Const A213      As Long = -1444681467
Private Const A214      As Long = -51403784
Private Const A215      As Long = 1735328473
Private Const A216      As Long = -1926607734

'Pass 3
Private Const S301      As Long = 4
Private Const S302      As Long = 11
Private Const S303      As Long = 16
Private Const S304      As Long = 23

Private Const A301      As Long = -378558
Private Const A302      As Long = -2022574463
Private Const A303      As Long = 1839030562
Private Const A304      As Long = -35309556
Private Const A305      As Long = -1530992060
Private Const A306      As Long = 1272893353
Private Const A307      As Long = -155497632
Private Const A308      As Long = -1094730640
Private Const A309      As Long = 681279174
Private Const A310      As Long = -358537222
Private Const A311      As Long = -722521979
Private Const A312      As Long = 76029189
Private Const A313      As Long = -640364487
Private Const A314      As Long = -421815835
Private Const A315      As Long = 530742520
Private Const A316      As Long = -995338651

'Pass 4
Private Const S401      As Long = 6
Private Const S402      As Long = 10
Private Const S403      As Long = 15
Private Const S404      As Long = 21

Private Const A401      As Long = -198630844
Private Const A402      As Long = 1126891415
Private Const A403      As Long = -1416354905
Private Const A404      As Long = -57434055
Private Const A405      As Long = 1700485571
Private Const A406      As Long = -1894986606
Private Const A407      As Long = -1051523
Private Const A408      As Long = -2054922799
Private Const A409      As Long = 1873313359
Private Const A410      As Long = -30611744
Private Const A411      As Long = -1560198380
Private Const A412      As Long = 1309151649
Private Const A413      As Long = -145523070
Private Const A414      As Long = -1120210379
Private Const A415      As Long = 718787259
Private Const A416      As Long = -343485551

'some bit masks
Private Const Top16Bits As Long = &HFFFF0000
Private Const Bot16Bits As Long = &HFFFF&
Private Const Bot8Bits  As Long = &HFF&
Private Const Bit16Only As Long = &H10000

Private Const BlockSize As Long = 64

'variables, types, and enums
Private Type Segs
    P1                  As Long
    P2                  As Long
    P3                  As Long
    P4                  As Long
End Type
Private Digest          As Segs

Private Sub DigestBlock(Data As String)

  'MD5 mangle, requires a string of length 64

  Dim IxBlock, Packed(0 To 15)
  Dim a, b, c, d

    'pack 4 bytes into 1 long -> 64 bytes into 16 unsigned longs = 512 bits
    For IxBlock = 1 To BlockSize Step 4
        Packed(IxBlock \ 4) = RotateLeft(RotateLeft(RotateLeft(Asc(Mid$(Data, IxBlock + 3, 1)), 8) Or Asc(Mid$(Data, IxBlock + 2, 1)), 8) Or Asc(Mid$(Data, IxBlock + 1, 1)), 8) Or Asc(Mid$(Data, IxBlock, 1))
    Next IxBlock

    With Digest
        'Pass 1 digest
        a = Pass1(.P1, .P2, .P3, .P4, Packed(0), S101, A101)
        d = Pass1(.P4, a, .P2, .P3, Packed(1), S102, A102)
        c = Pass1(.P3, d, a, .P2, Packed(2), S103, A103)
        b = Pass1(.P2, c, d, a, Packed(3), S104, A104)

        a = Pass1(a, b, c, d, Packed(4), S101, A105)
        d = Pass1(d, a, b, c, Packed(5), S102, A106)
        c = Pass1(c, d, a, b, Packed(6), S103, A107)
        b = Pass1(b, c, d, a, Packed(7), S104, A108)

        a = Pass1(a, b, c, d, Packed(8), S101, A109)
        d = Pass1(d, a, b, c, Packed(9), S102, A110)
        c = Pass1(c, d, a, b, Packed(10), S103, A111)
        b = Pass1(b, c, d, a, Packed(11), S104, A112)

        a = Pass1(a, b, c, d, Packed(12), S101, A113)
        d = Pass1(d, a, b, c, Packed(13), S102, A114)
        c = Pass1(c, d, a, b, Packed(14), S103, A115)
        b = Pass1(b, c, d, a, Packed(15), S104, A116)

        'Pass 2 digest
        a = Pass2(a, b, c, d, Packed(1), S201, A201)
        d = Pass2(d, a, b, c, Packed(6), S202, A202)
        c = Pass2(c, d, a, b, Packed(11), S203, A203)
        b = Pass2(b, c, d, a, Packed(0), S204, A204)

        a = Pass2(a, b, c, d, Packed(5), S201, A205)
        d = Pass2(d, a, b, c, Packed(10), S202, A206)
        c = Pass2(c, d, a, b, Packed(15), S203, A207)
        b = Pass2(b, c, d, a, Packed(4), S204, A208)

        a = Pass2(a, b, c, d, Packed(9), S201, A209)
        d = Pass2(d, a, b, c, Packed(14), S202, A210)
        c = Pass2(c, d, a, b, Packed(3), S203, A211)
        b = Pass2(b, c, d, a, Packed(8), S204, A212)

        a = Pass2(a, b, c, d, Packed(13), S201, A213)
        d = Pass2(d, a, b, c, Packed(2), S202, A214)
        c = Pass2(c, d, a, b, Packed(7), S203, A215)
        b = Pass2(b, c, d, a, Packed(12), S204, A216)

        'Pass 3 digest
        a = Pass3(a, b, c, d, Packed(5), S301, A301)
        d = Pass3(d, a, b, c, Packed(8), S302, A302)
        c = Pass3(c, d, a, b, Packed(11), S303, A303)
        b = Pass3(b, c, d, a, Packed(14), S304, A304)

        a = Pass3(a, b, c, d, Packed(1), S301, A305)
        d = Pass3(d, a, b, c, Packed(4), S302, A306)
        c = Pass3(c, d, a, b, Packed(7), S303, A307)
        b = Pass3(b, c, d, a, Packed(10), S304, A308)

        a = Pass3(a, b, c, d, Packed(13), S301, A309)
        d = Pass3(d, a, b, c, Packed(0), S302, A310)
        c = Pass3(c, d, a, b, Packed(3), S303, A311)
        b = Pass3(b, c, d, a, Packed(6), S304, A312)

        a = Pass3(a, b, c, d, Packed(9), S301, A313)
        d = Pass3(d, a, b, c, Packed(12), S302, A314)
        c = Pass3(c, d, a, b, Packed(15), S303, A315)
        b = Pass3(b, c, d, a, Packed(2), S304, A316)

        'Pass 4 digest
        a = Pass4(a, b, c, d, Packed(0), S401, A401)
        d = Pass4(d, a, b, c, Packed(7), S402, A402)
        c = Pass4(c, d, a, b, Packed(14), S403, A403)
        b = Pass4(b, c, d, a, Packed(5), S404, A404)

        a = Pass4(a, b, c, d, Packed(12), S401, A405)
        d = Pass4(d, a, b, c, Packed(3), S402, A406)
        c = Pass4(c, d, a, b, Packed(10), S403, A407)
        b = Pass4(b, c, d, a, Packed(1), S404, A408)

        a = Pass4(a, b, c, d, Packed(8), S401, A409)
        d = Pass4(d, a, b, c, Packed(15), S402, A410)
        c = Pass4(c, d, a, b, Packed(6), S403, A411)
        b = Pass4(b, c, d, a, Packed(13), S404, A412)

        a = Pass4(a, b, c, d, Packed(4), S401, A413)
        d = Pass4(d, a, b, c, Packed(11), S402, A414)
        c = Pass4(c, d, a, b, Packed(2), S403, A415)
        b = Pass4(b, c, d, a, Packed(9), S404, A416)

        .P1 = UnsignedAdd(.P1, a, 0, 0)
        .P2 = UnsignedAdd(.P2, b, 0, 0)
        .P3 = UnsignedAdd(.P3, c, 0, 0)
        .P4 = UnsignedAdd(.P4, d, 0, 0)
    End With 'DIGEST

End Sub

Private Function Pass1(ByVal a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long, ByVal x As Long, ByVal s As Long, ByVal Ac As Long) As Long

    Pass1 = UnsignedAdd(RotateLeft(UnsignedAdd(a, (b And c) Or (d And Not b), x, Ac), s), b, 0, 0)

End Function

Private Function Pass2(ByVal a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long, ByVal x As Long, ByVal s As Long, ByVal Ac As Long) As Long

    Pass2 = UnsignedAdd(RotateLeft(UnsignedAdd(a, (b And d) Or (c And Not d), x, Ac), s), b, 0, 0)

End Function

Private Function Pass3(ByVal a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long, ByVal x As Long, ByVal s As Long, ByVal Ac As Long) As Long

    Pass3 = UnsignedAdd(RotateLeft(UnsignedAdd(a, b Xor c Xor d, x, Ac), s), b, 0, 0)

End Function

Private Function Pass4(ByVal a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long, ByVal x As Long, ByVal s As Long, ByVal Ac As Long) As Long

    Pass4 = UnsignedAdd(RotateLeft(UnsignedAdd(a, c Xor (b Or Not d), x, Ac), s), b, 0, 0)

End Function

Private Function RotateLeft(ByVal Value As Long, ByVal NumPlaces As Long) As Long

  'rotate a long numplaces to the left

  Dim CntPlaces, SpillOver

    For CntPlaces = 1 To NumPlaces
        SpillOver = Value And &HC0000000
        Value = (Value And &H3FFFFFFF) * 2 Or ((SpillOver < 0) And 1) Or (CBool(SpillOver And &H40000000) And &H80000000)
    Next CntPlaces
    RotateLeft = Value

End Function

Public Function Signature(Reset As Boolean, Data As String) As String

  'digest a string and output the result as a string
  'this can be called several times in a row, for very long texts which have to be split
  'first call for must reset, subsequent calls must not

  Dim ip, Padding As String

    If Len(Data) Then
        With Digest
            If Reset Then 'initialize the algorithm
                .P1 = ToUnsignedLong(IniPart1)
                .P2 = ToUnsignedLong(IniPart2)
                .P3 = ToUnsignedLong(IniPart3)
                .P4 = ToUnsignedLong(IniPart4)
            End If

            'create pseudo random padding, depending on data
            Padding = Mid$(Data, (Len(Data) + 1) \ 2, 1)
            Do
                Padding = Mid$(Data, Asc(Padding) Mod Len(Data) + 1, 1) & Padding
            Loop Until Len(Padding) = BlockSize - 1

            'digest the stuff in chunks
            For ip = 1 To Len(Data) Step BlockSize
                DigestBlock Mid$(Data & Padding, ip, BlockSize)
            Next ip

            Signature = ""
            For ip = 0 To 3
                Signature = Signature & Chr$(.P1 \ 256 ^ ip And Bot8Bits) & Chr$(.P2 \ 256 ^ ip And Bot8Bits) & Chr$(.P3 \ 256 ^ ip And Bot8Bits) & Chr$(.P4 \ 256 ^ ip And Bot8Bits)
            Next ip
        End With 'DIGEST
    End If

End Function

Private Function ToUnsignedLong(Value As Currency) As Long

  'convert a currency datatype to an unsigned long

  Const Over As Currency = 2 ^ 32

    If Value <= Maxlong Then
        ToUnsignedLong = Value
      Else 'NOT VALUE...
        ToUnsignedLong = Value - Over
    End If

End Function

Private Function UnsignedAdd(ByVal Summand1 As Long, ByVal Summand2 As Long, ByVal Summand3 As Long, ByVal Summand4 As Long) As Long

  'adds four unsigned numbers together, overflows are ignored

  Dim Low, High

    Low = (Summand1 And Bot16Bits) + (Summand2 And Bot16Bits) + (Summand3 And Bot16Bits) + (Summand4 And Bot16Bits)
    High = ((Summand1 And Top16Bits) \ Bit16Only + (Summand2 And Top16Bits) \ Bit16Only + (Summand3 And Top16Bits) \ Bit16Only + (Summand4 And Top16Bits) \ Bit16Only + Low \ Bit16Only) And Bot16Bits
    UnsignedAdd = ToUnsignedLong(CCur(High) * Bit16Only + (Low And Bot16Bits))

End Function

':) Ulli's VB Code Formatter V2.10.7 (24.02.2002 22:14:17) 123 + 208 = 331 Lines
