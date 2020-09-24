VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} dCompare 
   ClientHeight    =   10005
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   11805
   _ExtentX        =   20823
   _ExtentY        =   17648
   _Version        =   393216
   Description     =   "Ulli's VB Source Code Comparator"
   DisplayName     =   "Ulli's VB Source Code Comparator"
   AppName         =   "Visual Basic"
   AppVer          =   "Visual Basic 6.0"
   LoadName        =   "Startup"
   LoadBehavior    =   1
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Visual Basic\6.0"
   CmdLineSupport  =   -1  'True
End
Attribute VB_Name = "dCompare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'© 2002     UMGEDV GmbH  (umgedv@aol.com)
'
'Author     UMG (Ulli K. Muehlenweg)
'
'Title      VB6 Source File Comparator Add-In
'
'Purpose    This AddIn lets compare the current state of your source in the IDE
'           with the source as stored in the corresponding source-file.
'
'           Compile the DLL into your VB directory and then use the Add-Ins
'           Manager to load the Comparator Add-In into VB.
'
'*********************************************************************************
'Development History
'*********************************************************************************
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'24 Feb 2002 Version 1.1.4      UMG
'
'Tidy source
'
'Changed scan thru Sigs to make sure that a signature is found and not something
'in between. This could happen if the tail of one sig and the head of the next
'by coincidence combine to be legal.
'
'Changed 2nd and 3rd param for GetTempFileName
'
'Added check for new file (no old file to compare it with)
'
'Added progress indication
'
'Added tool tips and save option
'
'Removed MsgHook, is not needed any more.
'
'Removed position sensitivity; a line which is only shifted (indented) is considered
'equal to the non-shifted line. Changes to the source by the Code Formatter are thus
'suppressed.
'
'Now outputs new line for unaltered lines to reflect the most recent state for
'shifted lines.
'
'Fixed bug with saving and restoring (.frx files for example)
'
'Known quirk: For no evident reason VB seems to alter .frx files occasionally, so
'             the AddIn may show changes in connection with this where in fact there
'             are none.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'22 Feb 2002 Version 1.0.7      Prototype - UMG
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
DefLng A-Z 'we're 32 bit

Private Const MenuName          As String = "Add-Ins" 'you may need to localize "Add-Ins"
Private Const Backslash         As String = "\"
Attribute Backslash.VB_VarDescription = "A single backslash"
Private Const SleepTime         As Long = 999
Attribute SleepTime.VB_VarDescription = "Time to show the Splash Form"
Private VBInstance              As VBIDE.VBE
Private CommandBarMenu          As CommandBar
Private MenuItem                As CommandBarControl
Private WithEvents MenuEvents   As CommandBarEvents
Attribute MenuEvents.VB_VarHelpID = -1
Private OrigFilenames()         As String
Private TempFilenames()         As String
Private IsDirty                 As Boolean
Private i                       As Long
Private CaptText                As String

Private Sub AddinInstance_OnBeginShutdown(custom() As Variant)

    If Not fCompare Is Nothing Then
        fCompare.WindowState = vbMinimized
        DoEvents
        Unload fCompare
    End If

End Sub

Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)

  Dim ClipboardText As String

    Set VBInstance = Application
    If ConnectMode = ext_cm_External Then
        CompareSources
      Else 'NOT CONNECTMODE...
        On Error Resume Next
          Set CommandBarMenu = VBInstance.CommandBars(MenuName)
        On Error GoTo 0
        If CommandBarMenu Is Nothing Then
            MsgBox "VB Source Comaparator was loaded but could not be connected to the " & MenuName & " menu.", vbCritical
          Else 'NOT COMMANDBARMENU...
            fSplash.Show
            DoEvents
            With CommandBarMenu
                Set MenuItem = .Controls.Add(msoControlButton)
                i = .Controls.Count - 1
                If .Controls(i).BeginGroup And Not .Controls(i - 1).BeginGroup Then
                    'menu separator required
                    MenuItem.BeginGroup = True
                End If
            End With 'COMMANDBARMENU
            'set menu caption
            With App
                MenuItem.Caption = "&" & .ProductName & " V" & .Major & "." & .Minor & "." & .Revision & "..."
            End With 'APP
            With Clipboard
                ClipboardText = .GetText
                'set menu picture
                .SetData fSplash.picMenu.Image
                MenuItem.PasteFace
                .Clear
                .SetText ClipboardText
            End With 'CLIPBOARD
            'set event handler
            Set MenuEvents = VBInstance.Events.CommandBarEvents(MenuItem)
            'done connecting
            Sleep SleepTime
            Unload fSplash
        End If
    End If

End Sub

Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)

    On Error Resume Next
      MenuItem.Delete
    On Error GoTo 0

End Sub

Public Sub CompareSources()
Attribute CompareSources.VB_Description = "The main Sub"
Attribute CompareSources.VB_UserMemId = 0

  Dim fc As Long
  
    If VBInstance.CodePanes.Count Then
        CaptText = App.ProductName & " - " & VBInstance.ActiveVBProject.Name & " [" & VBInstance.SelectedVBComponent.Name & "]"
        With VBInstance.SelectedVBComponent
            If Len(.FileNames(1)) = 0 Then
                MsgBox "This file is new; there is no old file to compare it with.", vbInformation, CaptText
              Else 'NOT LEN(.FILENAMES(1))...
                If MsgBox("Do you want to compare this Source File?", vbQuestion + vbOKCancel, CaptText) = vbOK Then
                    IsDirty = .IsDirty
                              
                    fCompare.Caption = CaptText
                    
                    fc = .FileCount
                    ReDim OrigFilenames(1 To fc), TempFilenames(1 To fc)
                    For i = 1 To fc
                        OrigFilenames(i) = .FileNames(i)
                        TempFilenames(i) = String$(255, 0)
                        GetTempFileName Left$(OrigFilenames(i), InStrRev(OrigFilenames(i), Backslash)), "UMG", 0, TempFilenames(i)
                        TempFilenames(i) = Left$(TempFilenames(i), InStr(TempFilenames(i), Chr$(0)) - 1)
                        FileCopy OrigFilenames(i), TempFilenames(i)
                    Next i
                    
                    .SaveAs OrigFilenames(1) 'may alter the .frx file
                    fCompare.Compare TempFilenames(1), OrigFilenames(1)
                    .IsDirty = IsDirty
                    
                    For i = 1 To fc
                        FileCopy TempFilenames(i), OrigFilenames(i)
                        Kill TempFilenames(i)
                    Next i
                                    
                End If
            End If
        End With 'VBINSTANCE.SELECTEDVBCOMPONENT
      Else 'VBINSTANCE.CODEPANES.COUNT = FALSE
        MsgBox "Cannot see any Code - you must open a Code Panel first.", vbExclamation, "Ulli's VB Code Comparator"
    End If
        
End Sub

Private Sub MenuEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)

    CompareSources

End Sub

':) Ulli's VB Code Formatter V2.10.7 (24.02.2002 22:14:11) 68 + 115 = 183 Lines
