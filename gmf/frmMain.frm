VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get All Info of supported Media files!"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10530
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   10530
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   360
      Left            =   9510
      TabIndex        =   12
      ToolTipText     =   "Close Application"
      Top             =   5205
      Width           =   945
   End
   Begin VB.Frame Frame2 
      Caption         =   "Option Export"
      Height          =   705
      Left            =   5865
      TabIndex        =   9
      Top             =   4995
      Width           =   2160
      Begin VB.PictureBox PicBack02 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   45
         ScaleHeight     =   465
         ScaleWidth      =   2070
         TabIndex        =   10
         Top             =   195
         Width           =   2070
         Begin VB.CheckBox ChckSavetoFile 
            Caption         =   "Save result Media to File"
            Height          =   345
            Left            =   45
            TabIndex        =   11
            ToolTipText     =   "Save the result into File TXT"
            Top             =   75
            Width           =   1980
         End
      End
   End
   Begin VB.CommandButton cmdBrowser 
      Caption         =   "..."
      Height          =   255
      Left            =   9990
      TabIndex        =   8
      ToolTipText     =   "Browse All Supported Media file"
      Top             =   4710
      Width           =   450
   End
   Begin VB.Frame Frame1 
      Caption         =   "Option Info"
      Height          =   720
      Left            =   90
      TabIndex        =   3
      Top             =   4995
      Width           =   5760
      Begin VB.PictureBox PicBack 
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   30
         ScaleHeight     =   480
         ScaleWidth      =   5685
         TabIndex        =   4
         Top             =   195
         Width           =   5685
         Begin VB.OptionButton optGetInfo 
            Caption         =   "Full and Short Info"
            Height          =   285
            Index           =   2
            Left            =   2805
            TabIndex        =   7
            ToolTipText     =   "Get Full and Short Info of Media file"
            Top             =   105
            Width           =   2355
         End
         Begin VB.OptionButton optGetInfo 
            Caption         =   "Short Info"
            Height          =   285
            Index           =   1
            Left            =   1395
            TabIndex        =   6
            ToolTipText     =   "Get Short Info of Media file"
            Top             =   105
            Width           =   1440
         End
         Begin VB.OptionButton optGetInfo 
            Caption         =   "Full Info"
            Height          =   285
            Index           =   0
            Left            =   60
            TabIndex        =   5
            ToolTipText     =   "Get Full Info of Media file"
            Top             =   105
            Value           =   -1  'True
            Width           =   1320
         End
      End
   End
   Begin VB.CommandButton cmdGetInfo 
      Caption         =   "Get Info"
      Enabled         =   0   'False
      Height          =   360
      Left            =   8115
      TabIndex        =   2
      ToolTipText     =   "get Info of Media selected"
      Top             =   5205
      Width           =   1245
   End
   Begin VB.TextBox txtMedia_Path 
      Height          =   270
      Left            =   90
      TabIndex        =   1
      ToolTipText     =   "Path of Media file"
      Top             =   4710
      Width           =   9840
   End
   Begin VB.TextBox txtBody 
      Height          =   4545
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      ToolTipText     =   "Body text"
      Top             =   75
      Width           =   10290
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private InfoMediadll As String
Private defPath As String
Private Sub cmdBrowser_Click()
    Dim mediaFileName As String
    If defPath = Empty Then defPath = App.Path + "\"
    mediaFileName = DialogOpenFile("All Media Files" + Chr$(0) + "*.mpeg;*.wmv;*.divx;*.dat;*.mpx;*.asf;*.avi;*.mov;" _
    & "*.mpg;*.mp3;*wma;*wav;*.mid;*.mp4;*.mp2;*.mkv;*.flac;*.ogg;*.vob;*.mka;*.mks;*.ogm;*.mpgv;*.mpv;*.m1v;" _
    & "*.m2v;*.qt;*.rm;*.rmvb;*.ra;*.ifo;*.ac3;*.dts;*.aac;*.ape;*.mac;*.aiff;*.aifc;*.au;*.iff;*.svx8;*.sv16;" _
    & "*.paf;*.sd2;*.irca;*.w64;*.mat;*.pvf;*.xi;*.sds;*.avr" + Chr$(0), defPath)
    If mediaFileName = Empty Then Exit Sub
    defPath = GetFilePath(mediaFileName, Only_Path)
    txtMedia_Path.Text = mediaFileName
    cmdGetInfo.Enabled = True
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGetInfo_Click()
    Dim display As String
    Dim Handle As Long
    
    '/// Cleat Text Body
    '/// *****************************************************************
    txtBody.Text = Empty
    
    Handle = MediaInfo_Open(StrPtr(txtMedia_Path.Text))
    
    display = InfoMediadll
    
    If optGetInfo(1).Value = True Then
        '/// Get short Info of Media file
        '/// *****************************************************************
        display = display + "Get (Short Info)..." + vbCrLf
        display = display + "------------------------------------------------" + vbCrLf
        Call MediaInfo_Option(Handle, StrPtr("Complete"), StrPtr(""))
        display = display + StripStrinCtoVB(MediaInfo_Inform(Handle, InformOption_Nothing))
    ElseIf optGetInfo(0).Value = True Then
        '/// Get full Info of Media file
        '/// *****************************************************************
        display = display + "Get (Full Info)..." + vbCrLf
        display = display + "------------------------------------------------" + vbCrLf
        Call MediaInfo_Option(Handle, StrPtr("Complete"), StrPtr("1"))
        display = display + StripStrinCtoVB(MediaInfo_Inform(Handle, InformOption_Nothing))
    ElseIf optGetInfo(2).Value = True Then
        '/// Get short and full Info of Media file
        '/// *****************************************************************
        display = display + "Get (Full and Short Info)..." + vbCrLf
        display = display + "------------------------------------------------" + vbCrLf
        Call MediaInfo_Option(Handle, StrPtr("Complete"), StrPtr(""))
        display = display + StripStrinCtoVB(MediaInfo_Inform(Handle, InformOption_Nothing))
        Call MediaInfo_Option(Handle, StrPtr("Complete"), StrPtr("1"))
        display = display + StripStrinCtoVB(MediaInfo_Inform(Handle, InformOption_Nothing))
    End If
  
    '/// Close MediaInfo.dll
    '/// *****************************************************************
    display = display + "------------------------------------------------" + vbCrLf
    display = display + vbCrLf + "Close MediaInfo.dll..."
    Call MediaInfo_Close(Handle)

    '/// Displaying the text Result of Media file
    '/// *****************************************************************
    txtBody.Text = display
    
    '/// If ChckSavetoFile value = 1 then Save the Info File
    '/// *****************************************************************
    Dim sFileName As String: Dim strFileName As String
    If ChckSavetoFile.Value = 1 Then
        strFileName = Mid$(GetFilePath(txtMedia_Path.Text, Only_FileName_and_Extension), _
        1, Len(GetFilePath(txtMedia_Path.Text, Only_FileName_and_Extension)) - 4)
        'sFileName = DialogSaveAs(App.Path & "\", Mid$(strFileName, 1, Len(strFileName) - 4))
        'If sFileName = Empty Then Exit Sub
        Open App.Path + "\" + strFileName + ".txt" For Output As #1
            Print #1, display
        Close #1
        MsgBox "Media Info file created in _/" & App.Path & vbCr & "File _/" & strFileName & ".txt", vbInformation, App.Title
    End If
    display = Empty
Exit Sub
End Sub

Private Sub Form_Initialize()
    '/// Init Controls XP/Vista Manifest
    '/// *****************************************************************
    Call InitCommonControlsVB
End Sub

Private Sub Form_Load()
    
    '/// Verify if the DLL is in the current Path
    '/// *****************************************************************
    If FileExists(App.Path + "\MediaInfo.dll") = False Then
            MsgBox "Sorry, the MediaInfo.dll not found in the current path!" & vbCr _
            & "Put the {MediaInfo.dll} into current path before runnig this Application!", vbExclamation, App.Title
        Unload Me
    End If
    
    '/// Get the MediaInfo.dll Info version
    '/// *****************************************************************
    InfoMediadll = "Init MediaInfo.dll" + vbCrLf + vbCrLf
    InfoMediadll = InfoMediadll + StripStrinCtoVB(MediaInfo_Option(0, StrPtr("Info_Version"), StrPtr(""))) + vbCrLf
    InfoMediadll = InfoMediadll + "------------------------------------------------" + vbCrLf + vbCrLf
    '/// Displaying the text Result of Media file
    '/// *****************************************************************
    txtBody.Text = InfoMediadll
    
    Exit Sub
End Sub


Private Sub Form_Terminate()
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub


