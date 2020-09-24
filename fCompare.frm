VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form fCompare 
   BackColor       =   &H00E0E0E0&
   Caption         =   "File Compare"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8160
   Icon            =   "fCompare.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   543
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   544
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.OptionButton opDummy 
      Caption         =   "Option1"
      Height          =   195
      Left            =   -300
      TabIndex        =   0
      Top             =   165
      Width           =   210
   End
   Begin RichTextLib.RichTextBox rtBox 
      Height          =   2760
      Left            =   105
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   750
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   4868
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      MousePointer    =   1
      DisableNoScroll =   -1  'True
      RightMargin     =   9999
      TextRTF         =   $"fCompare.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton btAbout 
      Caption         =   "&About"
      Height          =   285
      Left            =   4425
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Show About Box"
      Top             =   345
      Width           =   600
   End
   Begin VB.CommandButton btMail 
      Caption         =   "&Mail"
      Height          =   285
      Left            =   4425
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Send mail to author"
      Top             =   60
      Width           =   600
   End
   Begin VB.Image img 
      BorderStyle     =   1  'Fest Einfach
      Height          =   765
      Left            =   105
      Picture         =   "fCompare.frx":052A
      ToolTipText     =   "© 2002 UMGEDV"
      Top             =   0
      Width           =   825
   End
   Begin VB.Label lbEqual 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   1065
      TabIndex        =   4
      Top             =   240
      Width           =   225
   End
End
Attribute VB_Name = "fCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefLng A-Z 'we're 32 bit!

Private Digest              As cMD5
Private hFileOld            As Long
Private InOld               As Long
Private hFileNew            As Long
Private InNew               As Long
Private i                   As Long
Private j                   As Long
Private LenFiles            As Long
Private SoFar               As Long
Private PrevStep            As Long
Private FirstDiff           As Long
Private PgStep              As Single
Private Comparing           As Boolean
Private DrawSep             As Boolean

Private Textline            As String
Private SigsOld             As String
Private SigsNew             As String

Private Const StepChar      As String = ""
Private Const Red           As Long = &HD0
Private Const Blue          As Long = &HD00000

Private Sub btAbout_Click()
    
    With App
        ShellAbout Me.hWnd, "About " & .ProductName & "#Operating System:", AppDetails & vbCrLf & .LegalCopyright, Me.Icon.Handle
    End With 'APP

End Sub

Private Sub btMail_Click()

    WindowState = vbMinimized
    SendMeMail hWnd, AppDetails

End Sub

Public Sub Compare(OldFileName As String, NewFileName As String)
  
    FirstDiff = -1
    Comparing = True
    Show
    SoFar = 0
    PrevStep = 0
    Form_Resize
    DoEvents
    Screen.MousePointer = vbHourglass
    With rtBox
        .Text = ""
        .ToolTipText = ""
        .Enabled = False
        hFileOld = FreeFile
        Open OldFileName For Input Access Read Shared As hFileOld
        LenFiles = LOF(hFileOld)
        SigsOld = ""
        hFileNew = FreeFile
        Open NewFileName For Input Access Read Shared As hFileNew
        LenFiles = LenFiles + LOF(hFileNew)
        LenFiles = LenFiles + LenFiles
        SigsNew = ""
        Do Until EOF(hFileOld)
            Line Input #hFileOld, Textline
            SoFar = SoFar + Len(Textline) + 2 'for crlf
            If Len(Trim$(Textline)) Then
                SigsOld = SigsOld & Digest.Signature(True, Trim$(Textline))
            End If
            UpdateProgress SoFar / LenFiles
        Loop
        Do Until EOF(hFileNew)
            Line Input #hFileNew, Textline
            SoFar = SoFar + Len(Textline) + 2 'for crlf
            If Len(Trim$(Textline)) Then
                SigsNew = SigsNew & Digest.Signature(True, Trim$(Textline))
            End If
            UpdateProgress SoFar / LenFiles
        Loop
        Close hFileOld, hFileNew
   
        hFileOld = FreeFile
        Open OldFileName For Input Access Read Shared As hFileOld
        hFileNew = FreeFile
        Open NewFileName For Input Access Read Shared As hFileNew
        Do
            If Len(SigsOld) Then
                InOld = 0
                Do
                    InOld = InStr(InOld + 1, SigsOld, Left$(SigsNew, 16))
                Loop Until (InOld Mod 16) < 2
              Else 'LEN(SigsOld) = FALSE
                InOld = -1
            End If
            If Len(SigsNew) Then
                InNew = 0
                Do
                    InNew = InStr(InNew + 1, SigsNew, Left$(SigsOld, 16))
                Loop Until (InNew Mod 16) < 2
              Else 'LEN(SigsNew) = FALSE
                InNew = -1
            End If
            If InOld = -1 And InNew = -1 Then
                Exit Do '>---> Loop
            End If
            Select Case True
              Case InOld = 1 And InNew = 1
                '1st line of Old = 1st line of New
                'output new line and skip old line
                OutNewLine vbBlack
                GetLine (hFileOld)
                SigsOld = Mid$(SigsOld, 17)
                SigsNew = Mid$(SigsNew, 17)
              Case InOld = 0 And InNew = 0
                '1st line of Old is not in New and viceversa
                'output lines as deleted and new
                OutDelLine GetLine(hFileOld) & vbCrLf, Blue
                OutNewLine Red
                SigsOld = Mid$(SigsOld, 17)
                SigsNew = Mid$(SigsNew, 17)
              Case InOld <= 0
                '1st line of New is not in Old
                'there are new lines in New - output them
                OutNewLine Red
                SigsNew = Mid$(SigsNew, 17)
              Case InNew <= 0
                '1st line of Old is not in New
                'lines have been deleted
                OutDelLine GetLine(hFileOld) & vbCrLf, Blue
                SigsOld = Mid$(SigsOld, 17)
              Case Else
                'line is in both files at different positions
                If InNew < InOld Then
                    OutNewLine Red
                    SigsNew = Mid$(SigsNew, 17)
                  Else 'NOT InNew...
                    OutDelLine GetLine(hFileOld) & vbCrLf, Blue
                    SigsOld = Mid$(SigsOld, 17)
                End If
            End Select
            UpdateProgress SoFar / LenFiles
        Loop
        If FirstDiff > 0 Then
            .SelStart = FirstDiff
        End If
        Close hFileOld, hFileNew
        Comparing = False
        With lbEqual
            .AutoSize = True
            .ForeColor = vbBlack
            .Caption = " The files are " & IIf(FirstDiff < 0, "equal. ", "different. ")
            .Visible = True
        End With 'LBEQUAL
        Screen.MousePointer = vbDefault
        .Enabled = True
        .ToolTipText = "Compare results; right click to save"
    End With 'RTBOX

End Sub

Private Sub Form_Load()

  Dim Fontname As String
  Dim Fontsize As Long
  
    Set Digest = New cMD5
    If RegOpenKeyEx(HKEY_CURRENT_USER, VBSettings, REG_OPTION_RESERVED, KEY_QUERY_VALUE, i) <> ERROR_NONE Then
        Fontname = "Fixedsys"
        Fontsize = 9
      Else 'NOT REGOPENKEYEX(HKEY_CURRENT_USER,...
        Fontname = String$(128, 0)
        j = Len(Fontname)
        If RegQueryValueEx(i, Fontface, REG_OPTION_RESERVED, j, ByVal Fontname, j) <> ERROR_NONE Then
            Fontname = "Fixedsys"
          Else 'NOT REGQUERYVALUEEX(I,...
            Fontname = Left$(Fontname, j + (Asc(Mid$(Fontname, j, 1)) = 0))
        End If
        j = Len(Fontsize)
        If RegQueryValueEx(i, Fontheight, REG_OPTION_RESERVED, j, Fontsize, j) <> ERROR_NONE Then
            Fontsize = 9
        End If
        RegCloseKey i
    End If
    rtBox.Font.Name = Fontname
    rtBox.Font.Size = Fontsize
    
End Sub

Private Sub Form_Resize()

  Dim sx As String

    If WindowState <> vbMinimized Then
        rtBox.Move 7, 50, ScaleWidth - 14, ScaleHeight - 57
        btMail.Left = ScaleWidth - 8 - btMail.Width
        btAbout.Left = btMail.Left
        If Comparing Then
            With lbEqual
                .Width = btMail.Left - lbEqual.Left - 10
                .AutoSize = False
                .ForeColor = Red
                Set Font = .Font
                PgStep = (TextWidth(StepChar & StepChar) - TextWidth(StepChar)) / (.Width - 2)
            End With 'LBEQUAL
        End If
    End If

End Sub

Private Function GetLine(FromFile As Long) As String

    Do
        GetLine = ""
        If Not EOF(FromFile) Then
            Line Input #FromFile, GetLine
            SoFar = SoFar + Len(GetLine) + 2 'for crlf
        End If
    Loop Until Len(Trim$(GetLine)) Or EOF(FromFile)
    DrawSep = (GetLine = "End Sub" Or GetLine = "End Function" Or GetLine = "End Property")

End Function

Private Sub OutDelLine(Textline As String, Color As Long)

    SavePos
    With rtBox
        .SelStrikeThru = False
        For i = 1 To Len(Textline)
            If Mid$(Textline, i, 1) = " " Then
                .SelText = " "
              Else 'NOT MID$(TEXTLINE,...
                Exit For '>---> Next
            End If
        Next i
        .SelColor = Color
        .SelStrikeThru = True
        .SelText = Mid$(Textline, i)
    End With 'RTBOX

End Sub

Private Sub OutNewLine(Color As Long)
    
    If Color <> vbBlack Then
        SavePos
    End If
    With rtBox
        .SelColor = Color
        .SelStrikeThru = False
        .SelText = GetLine(hFileNew) & vbCrLf
    End With 'RTBOX
    SepOut
    
End Sub

Private Sub rtBox_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        rtBox.SaveFile "CompareResult.RTF"
        rtBox.ToolTipText = "Compare results saved"
    End If

End Sub

Private Sub SavePos()

    If FirstDiff < 0 Then
        FirstDiff = rtBox.SelStart
    End If
    
End Sub

Private Sub SepOut()

    If DrawSep Then
        With rtBox
            .SelColor = &HC0C0C0
            .SelStrikeThru = True
            .SelText = Space$(1023) & Chr$(160) & vbCrLf
        End With 'RTBOX
    End If

End Sub

Private Sub UpdateProgress(Percent As Single)

  Dim CurrStep
    
    CurrStep = Percent / PgStep
    If CurrStep <> PrevStep Then
        lbEqual = String$(CurrStep, StepChar)
        PrevStep = CurrStep
        DoEvents
    End If
    
End Sub

':) Ulli's VB Code Formatter V2.10.7 (24.02.2002 22:14:21) 25 + 274 = 299 Lines
