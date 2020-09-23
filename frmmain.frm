VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{19B7F2A2-1610-11D3-BF30-1AF820524153}#1.2#0"; "ccrpftv6.ocx"
Begin VB.Form frmmain 
   Caption         =   "Mass Renamer II"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   11745
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   11745
   StartUpPosition =   3  'Windows Default
   Begin MassRenamerII.XPWin XPWmain 
      Height          =   4215
      Index           =   4
      Left            =   2880
      TabIndex        =   27
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7435
      Caption         =   "Auto Rename Rules"
      Begin VB.TextBox txtindex 
         Height          =   285
         Index           =   10
         Left            =   4920
         TabIndex        =   52
         Text            =   "-"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox txtindex 
         Height          =   285
         Index           =   6
         Left            =   120
         TabIndex        =   32
         Text            =   "episode ep anime amv ova oav divx xvid volume ep tv dvd vol sd"
         Top             =   2040
         Width           =   8535
      End
      Begin VB.CheckBox chkmain 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Force lower case extentions"
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   44
         Top             =   1440
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.TextBox txtindex 
         Height          =   285
         Index           =   9
         Left            =   120
         TabIndex        =   40
         Text            =   "DNA^2"
         Top             =   3840
         Width           =   8535
      End
      Begin VB.TextBox txtindex 
         Height          =   285
         Index           =   8
         Left            =   120
         TabIndex        =   36
         Text            =   "the an no and a or in with on of over at"
         Top             =   3240
         Width           =   8535
      End
      Begin VB.TextBox txtindex 
         Height          =   285
         Index           =   7
         Left            =   120
         TabIndex        =   34
         Text            =   "gantz"
         Top             =   2640
         Width           =   8535
      End
      Begin VB.CheckBox chkmain 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Clean up the excess punctuation"
         Height          =   255
         Index           =   11
         Left            =   3240
         TabIndex        =   43
         Top             =   1200
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox chkmain 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Capitalize Each Word That Isn't"
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   42
         Top             =   1200
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox chkmain 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Replace all commas with spaces"
         Height          =   255
         Index           =   9
         Left            =   3240
         TabIndex        =   41
         Top             =   720
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox chkmain 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Dont seperate the numbers from the following words"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   39
         Top             =   3600
         Value           =   1  'Checked
         Width           =   4215
      End
      Begin VB.CheckBox chkmain 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Force all numbers to be 2 digits or more"
         Height          =   255
         Index           =   7
         Left            =   3240
         TabIndex        =   38
         Top             =   960
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox chkmain 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Correct redundant extentions"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   37
         Top             =   960
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox chkmain 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Lower case the following words"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   35
         Top             =   3000
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox chkmain 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Capitalize the following words"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   33
         Top             =   2400
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox chkmain 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Remove the following words"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   31
         Top             =   1800
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox chkmain 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Replace all underscores with spaces"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.CheckBox chkmain 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Replace all but 1 period with spaces"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   29
         Top             =   480
         Value           =   1  'Checked
         Width           =   3735
      End
      Begin VB.CheckBox chkmain 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Remove text inside brackets"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   28
         Top             =   480
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.Label lstsec 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Prepend numbers with:"
         Height          =   255
         Index           =   6
         Left            =   3240
         TabIndex        =   51
         Top             =   1470
         Width           =   1815
      End
      Begin VB.Label lblhelp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C65D21&
         Height          =   255
         Index           =   4
         Left            =   8160
         TabIndex        =   50
         Top             =   60
         Width           =   255
      End
      Begin VB.Image imgmain 
         Height          =   15
         Index           =   1
         Left            =   120
         Picture         =   "frmmain.frx":0442
         Stretch         =   -1  'True
         Top             =   1740
         Width           =   8535
      End
      Begin VB.Label lblmain 
         Alignment       =   2  'Center
         BackColor       =   &H00F7DFD6&
         Caption         =   "&Add Auto Rename to the shell"
         ForeColor       =   &H00C65D21&
         Height          =   615
         Index           =   9
         Left            =   7560
         TabIndex        =   45
         Top             =   480
         Width           =   1095
      End
   End
   Begin MSComctlLib.ListView lstview 
      Height          =   7695
      Left            =   2880
      TabIndex        =   0
      Top             =   480
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   13573
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      OLEDropMode     =   1
      _Version        =   393217
      Icons           =   "imlmain"
      SmallIcons      =   "imlmain"
      ColHdrIcons     =   "imlmain"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Custom Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Original Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date/Time"
         Object.Width           =   2540
      EndProperty
   End
   Begin MassRenamerII.XPWin XPWmain 
      Height          =   1335
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   6840
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   2355
      Caption         =   "Automatic methods"
      Begin VB.Label lblhelp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C65D21&
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   47
         Top             =   60
         Width           =   255
      End
      Begin VB.Label lblmain 
         BackColor       =   &H00F7DFD6&
         Caption         =   "&Auto Rename"
         ForeColor       =   &H00C65D21&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label lblmain 
         BackColor       =   &H00F7DFD6&
         Caption         =   "&Set to legacy 8.3 format"
         ForeColor       =   &H00C65D21&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lblmain 
         BackColor       =   &H00F7DFD6&
         Caption         =   "&Move files to another folder"
         ForeColor       =   &H00C65D21&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   2415
      End
   End
   Begin MassRenamerII.XPWin XPWmain 
      Height          =   2415
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   4260
      Caption         =   "Numeric indexing"
      Begin VB.TextBox txtindex 
         Height          =   285
         Index           =   4
         Left            =   120
         TabIndex        =   20
         Text            =   "#.[EXT]"
         ToolTipText     =   "Use a single # to represent where you want the numbers to be placed, and [EXT] to represent the old extention"
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox txtindex 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   13
         Text            =   "3"
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox txtindex 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   12
         Text            =   "1"
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblhelp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C65D21&
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   48
         Top             =   60
         Width           =   255
      End
      Begin VB.Label lblmain 
         BackColor       =   &H00F7DFD6&
         Caption         =   "&Use filename as index"
         ForeColor       =   &H00C65D21&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   24
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label lstsec 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Pattern:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblmain 
         BackColor       =   &H00F7DFD6&
         Caption         =   "&Change filenames"
         ForeColor       =   &H00C65D21&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label lstsec 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Digits:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label lstsec 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Starting value:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1695
      End
   End
   Begin MassRenamerII.XPWin XPWmain 
      Height          =   2055
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Top             =   4680
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   3625
      Caption         =   "Replace Text"
      Begin VB.TextBox txtindex 
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   2415
      End
      Begin VB.TextBox txtindex 
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lblhelp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C65D21&
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   49
         Top             =   60
         Width           =   255
      End
      Begin VB.Label lblmain 
         BackColor       =   &H00F7DFD6&
         Caption         =   "&Change filenames"
         ForeColor       =   &H00C65D21&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label lstsec 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Text to replace it with:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lstsec 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Text to search for:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   1695
      End
   End
   Begin MassRenamerII.XPWin XPWmain 
      Height          =   5415
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   9551
      Caption         =   "Operations"
      Begin CCRPFolderTV6.FolderTreeview dirmain 
         Height          =   3660
         Left            =   120
         TabIndex        =   53
         Top             =   840
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   6456
      End
      Begin VB.TextBox txtindex 
         Height          =   285
         Index           =   5
         Left            =   1080
         TabIndex        =   22
         Text            =   "*.*"
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblmain 
         BackColor       =   &H00F7DFD6&
         Caption         =   "&Refresh"
         ForeColor       =   &H00C65D21&
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   54
         Top             =   4560
         Width           =   2415
      End
      Begin VB.Label lblhelp 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C65D21&
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   46
         Top             =   60
         Width           =   255
      End
      Begin VB.Label lstsec 
         BackColor       =   &H00F7DFD6&
         Caption         =   "Pattern:"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblmain 
         BackColor       =   &H00F7DFD6&
         Caption         =   "&Apply all changes"
         ForeColor       =   &H00C65D21&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   5040
         Width           =   2415
      End
      Begin VB.Label lblmain 
         BackColor       =   &H00F7DFD6&
         Caption         =   "&Undo all changes"
         ForeColor       =   &H00C65D21&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   4800
         Width           =   2415
      End
   End
   Begin VB.FileListBox Filmain 
      Height          =   480
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picmain 
      AutoRedraw      =   -1  'True
      BackColor       =   &H008000FF&
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2400
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.ImageList imlmain 
      Left            =   1680
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   8388863
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":07D0
            Key             =   ".folder"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&file"
      Visible         =   0   'False
      Begin VB.Menu mnumain 
         Caption         =   "&Index"
         Index           =   0
      End
      Begin VB.Menu mnumain 
         Caption         =   "&Extract Index"
         Index           =   1
      End
      Begin VB.Menu mnumain 
         Caption         =   "&Replace Text"
         Index           =   2
      End
      Begin VB.Menu mnumain 
         Caption         =   "&Move to a folder"
         Index           =   3
      End
      Begin VB.Menu mnumain 
         Caption         =   "&DOS 8.3"
         Index           =   4
      End
      Begin VB.Menu mnumain 
         Caption         =   "&Auto Rename"
         Index           =   5
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const vbdarkblue = &HC65D21
Const vblightblue = 16748098
Dim oldone As Integer, mustbe As Boolean

Private Sub chkmain_Click(Index As Integer)
    SaveSetting "MassRenamer II", "Auto Rename Check", CStr(Index), chkmain(Index).Value = vbChecked
End Sub

Private Sub dirmain_FolderClick(Folder As CCRPFolderTV6.Folder, Location As CCRPFolderTV6.ftvHitTestConstants)
If InStr(Folder.FullPath, "\") > 0 Then
    lstview.ListItems.Clear
    Filmain.Path = Folder.FullPath
    Filmain.Pattern = txtindex(5)
    SeedListView Filmain, lstview, imlmain, picmain
    ResetList lstview
End If
End Sub

Private Sub Form_Load()
    Dim temp As Long, tempstr As String
    For temp = 1 To XPWmain.UBound
        XPWmain(temp).State = 1
    Next
    For temp = 0 To chkmain.UBound
        chkmain(temp).Tag = GetSetting("MassRenamer II", "Auto Rename Check", CStr(temp), "True")
        chkmain(temp).Value = IIf(StrComp(chkmain(temp).Tag, "True", vbTextCompare) = 0, vbChecked, vbUnchecked)
    Next
    For temp = 0 To txtindex.UBound
        txtindex(temp).text = GetSetting("MassRenamer II", "Auto Rename Text", CStr(temp), txtindex(temp).text)
    Next
    oldone = -1
    HandleCommand Command
    MoveIndexes 0, 120, 0, 2, 3, 1
End Sub

Public Sub HandleCommand(Command As String)
    Dim tempstr As String, temp As Long, tempstr2() As String
    If Len(Command) > 0 Then
        'if the first and last digits are quotes, and contain no other quotes, OR there are no spaces in the command
        If InStr(2, Command, """") = Len(Command) Or InStr(1, Command, " ") = 0 Then
            tempstr = getfromquotes(Command) ', AutoRenameFile
            If Not RenameFile(tempstr, chkdir(GetPath(tempstr), AutoRenameFile(GetFilename(tempstr)))) Then MsgBox tempstr & vbNewLine & "was not renamed!", vbCritical, "An error occurred"
        Else
            If InStr(Command, """") = 0 Then
                tempstr2 = Split(Command, " ")
            Else
                tempstr2 = Split(Command, """ """)
            End If
            For temp = 0 To UBound(tempstr2)
                tempstr = getfromquotes(tempstr2(temp))
                If Not RenameFile(tempstr, chkdir(GetPath(tempstr), AutoRenameFile(GetFilename(tempstr)))) Then MsgBox tempstr & vbNewLine & "was not renamed!", vbCritical, "An error occurred"
            Next
        End If
        Unload Me
        End
    End If
End Sub

Private Sub lblhelp_Click(Index As Integer)
    Dim tempstr As String
    Select Case Index
        Case 0 'Operations
            tempstr = "Pattern: Use this to change which types of files appear when you load a new folder" & vbNewLine & _
                      "A * will represent multiple digits. ie: *.txt will show all files ending in .txt" & vbNewLine & _
                      "A ? will represent only 1 digits. ie: test?.txt will show files test1.txt, test2.txt, test3.txt, etc" & vbNewLine & vbNewLine & _
                      "Undo all changes: Click this to change all the filenames you've edited back to their originals" & vbNewLine & vbNewLine & _
                      "Apply all changes: Files aren't actually renamed until you click this"
        Case 2 'Numeric Indexing
            MsgBox "Pattern: Use this to change the filename to a pattern" & vbNewLine & _
                      "# will be replaced with an index number. ie: test#.txt becomes test001.txt, test002.txt, etc" & vbNewLine & _
                      "[EXT] will be replaced with the old extention. ie: test#.[EXT] becomes test001.txt, test002.gif, test003.jpg, etc" & vbNewLine & _
                      "[OLD] will be replaced with the old filename (everything except the extention)" & vbNewLine & vbNewLine & _
                      "Starting value: # will be replaced with this number, unless you click 'Use filename as index'" & vbNewLine & vbNewLine & _
                      "Digits: Numbers will be padded with zero's on the left side until they have this many digits" & vbNewLine & vbNewLine & _
                      "Change filenames: Changes all filenames to match the pattern using the number in 'Starting Value'" & vbNewLine & vbNewLine & _
                      "Use filename as index: Changes all filenames to match the pattern using numbers inside the filename itself", vbInformation, "Numeric Indexing"
            tempstr = Replace("The following use the last modified date of the file: " & vbNewLine & vbNewLine & _
                        "[DAY] the day/n[MONTH] Numeric month/n[SMONTH] 3 digit name of the month/n[LMONTH] full name of the month/n[YEAR] 4 digit year/n[SYEAR] 2 digit year" & vbNewLine & vbNewLine & _
                        "[12H] the 12 hour format hour/n[24H] the 24 hour format hour/n[MIN] the minute/n[AMPM] AM or PM", "/n", vbNewLine)
        Case 3 'Replace Text
            tempstr = "Text to search for: The text you want to get rid of goes here" & vbNewLine & vbNewLine & _
                      "Text to replace it with: The text you want put in it's place goes here" & vbNewLine & vbNewLine & _
                      "Change filenames: Replaces the text to search for in filenames"
        Case 1 'Automatic methods
            tempstr = "Move files to another folder: Makes the filenames unique so you can move them to another folder" & vbNewLine & vbNewLine & _
                      "Set to legacy 8.3 format: Rename files to their DOS/8.3 digit legacy filenames." & vbNewLine & vbNewLine & _
                      "Auto Rename: Automatically rename all files using the rules you've selected" & vbNewLine & vbNewLine & _
                      "To select individual files to rename, you can hold CTRL and click one at a time, or hold SHIFT and click 2 files to select all of the ones between them. Then Right click a selected file" & vbNewLine & vbNewLine & _
                      "When one file is selected, you can click F2 to type a new name manually"
        Case 4 'Auto rename rules
            MsgBox "Auto rename will do the following things IF the checkboxes are checked" & vbNewLine & vbNewLine & _
                   "Remove text inside brackets: All text inside (), {}, [] are removed" & vbNewLine & vbNewLine & _
                   "Replace all underscores with spaces: All _'s will be replaced with spaces" & vbNewLine & vbNewLine & _
                   "Correct redundant extentions: OGM will be replaced with AVI, DIZ and NFO with TXT" & vbNewLine & vbNewLine & _
                   "Capitalize Each Word That Isn't: If a word contains ALL or NO capitals, the first letter will be capitalized, and the rest lower cased" & vbNewLine & vbNewLine & _
                   "Force lower case extentions: Extentions will always be lower cased" & vbNewLine & vbNewLine & _
                   "Replace all but 1 period with spaces: Every period except for the last one will be replaced with a space" & vbNewLine & vbNewLine & _
                   "Replace all commas with spaces: All commas will be replaced with spaces", vbInformation, "Auto Rename Rules"
            MsgBox "Force all numbers to be 2 digits or more: Any time numbers are found, they'll be padded with 0's to make them 2 digits or more" & vbNewLine & vbNewLine & _
                   "Clean up the excess punctuation: Spaces and dashes at the beginning of a filename, before and after periods will be removed" & vbNewLine & vbNewLine & _
                   "Remove the following words: Remove all words in the filename that you specify" & vbNewLine & vbNewLine & _
                   "Capitalize the following words: FULLY CAPITALIZE all words in the filename that you specify" & vbNewLine & vbNewLine & _
                   "Lower case the following words: fully lowercase all words in the filename that you specify" & vbNewLine & vbNewLine & _
                   "Dont seperate the numbers from the following words: Certain words are meant to contain numbers that you don't want seperated and padded. Place them here." & vbNewLine & vbNewLine & _
                   "Prepend numbers with: When numbers are found, this text will be added before them.", vbInformation, "Auto Rename Rules"
           tempstr = "Add Auto Rename to the shell: Click this if you want to be able to right click a file or folder, and click Auto Rename." & vbNewLine & vbNewLine & _
                     "You can also drag files directly onto this EXE file, or the list box itself to auto rename them en masse"
        Case Else: tempstr = "Help on this function has not yet been made"
    End Select
    MsgBox tempstr, vbInformation, XPWmain(Index).Caption
End Sub

Private Sub lblmain_Click(Index As Integer)
    Dim tempstr As String
    Select Case Index
        Case 1, 2: If Not ResetList(lstview, CLng(Index)) Then MsgBox "One or more files failed to be renamed", vbCritical, "Error(s) occurred" 'Undo/Apply all changes
        Case 3: txtindex(0) = NumericIndex(lstview, txtindex(4), txtindex(0), txtindex(1), mustbe)  'Numeric indexing
        Case 4: ReplaceText lstview, txtindex(2), txtindex(3), mustbe 'Replace text
        Case 5: MakeAllUnique lstview, mustbe, Me.hWnd  'Move files
        Case 6: NumericIndex lstview, txtindex(4), txtindex(0), -txtindex(1), mustbe  'Numeric indexing using numbers within the filename itself
        Case 7: MakeAllDOS lstview, mustbe 'Old DOS
        Case 8: AutoRenameAll mustbe 'Auto Rename All, d'uh
        Case 9 'Enhance the shell
            SaveString HKEY_CLASSES_ROOT, "*\Shell\Auto Rename\Command", Empty, """" & chkdir(App.Path, App.EXEName) & ".exe"" %1"
            SaveString HKEY_CLASSES_ROOT, "Folder\Shell\Auto Rename\Command", Empty, """" & chkdir(App.Path, App.EXEName) & ".exe"" %1"
            MsgBox "The shell has been enhanced", vbInformation, "Shell enhancement complete"
        Case 10: dirmain_FolderClick dirmain.SelectedFolder, ftvOnFolder
    End Select
    autosizeall lstview
End Sub

Public Sub AutoRenameAll(Optional MustbeSelected As Boolean)
    Dim temp As Long, DOIT As Boolean
    For temp = 1 To lstview.ListItems.count
        With lstview.ListItems.Item(temp)
            DOIT = True
            If MustbeSelected Then DOIT = .Selected
            If DOIT Then .text = AutoRenameFile(.text)
        End With
    Next
End Sub

Private Sub lblmain_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lblmain(Index).ForeColor = vbdarkblue Then
        If oldone > -1 Then
            lblmain(oldone).ForeColor = vbdarkblue
            lblmain(oldone).Font.Underline = False
        End If
        lblmain(Index).ForeColor = vblightblue
        lblmain(Index).Font.Underline = True
        oldone = Index
    End If
End Sub

Private Sub lstview_AfterLabelEdit(Cancel As Integer, NewString As String)
    If StrComp(GetExtention(lstview.SelectedItem.text), GetExtention(NewString)) <> 0 Then
        If MsgBox("Are you sure you want to change the extention from '" & GetExtention(lstview.SelectedItem.text) & "' to '" & GetExtention(NewString) & "'?" & vbNewLine & "This could render the file unusable.", vbYesNo + vbQuestion, "Extention change") = vbNo Then
            NewString = lstview.SelectedItem.text
        End If
    End If
End Sub

Private Sub lstview_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If lstview.Sorted Then
        If lstview.SortOrder = lvwAscending Then
            lstview.SortOrder = lvwDescending
        Else
            lstview.SortOrder = lvwAscending
            lstview.SortKey = 0
            lstview.Sorted = False
        End If
    Else
        lstview.SortOrder = lvwAscending
        lstview.SortKey = ColumnHeader.Index - 1
        lstview.Sorted = True
    End If
End Sub

Private Sub lstview_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 17 Or KeyCode = 113 Then lstview.StartLabelEdit
End Sub

Private Sub lstview_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton And Not lstview.SelectedItem Is Nothing Then PopupMenu mnufile
End Sub

Private Sub mnumain_Click(Index As Integer)
    Dim temp() As String
    temp = Split("3,6,4,5,7,8", ",")
    mustbe = True
    lblmain_Click CInt(temp(Index))
    mustbe = False
End Sub

Private Sub txtindex_Change(Index As Integer)
    SaveSetting "MassRenamer II", "Auto Rename Text", CStr(Index), txtindex(Index).text
    If Index = 0 Then
        If Len(txtindex(0).text) > Val(txtindex(1)) Then txtindex(1) = Len(txtindex(0).text)
    End If
End Sub

Public Function GetCheck(Index As Long) As Boolean
    GetCheck = chkmain(Index).Value = vbChecked
End Function

Public Function AutoRenameFile(Filename As String) As String
    AutoRenameFile = AnimeRename(Filename, GetCheck(0), GetCheck(2), GetCheck(9), GetCheck(1), GetCheck(6), GetCheck(11), GetCheck(8), txtindex(9), GetCheck(10), GetCheck(7), 2, GetCheck(3), txtindex(6), GetCheck(4), txtindex(7), GetCheck(5), txtindex(8), GetCheck(12), txtindex(10))
End Function

Private Sub XPWmain_ChangeOver(Index As Integer, State As Boolean)
    If Not lblhelp(Index).Font.Underline = State Then
        lblhelp(Index).Font.Underline = State
        lblhelp(Index).ForeColor = IIf(State, vblightblue, vbdarkblue)
    End If
End Sub

Private Sub XPWmain_MouseMove(Index As Integer)
    If oldone > -1 Then
        lblmain(oldone).ForeColor = vbdarkblue
        lblmain(oldone).Font.Underline = False
        oldone = -1
    End If
End Sub

Private Sub XPWmain_Resize(Index As Integer)
    MoveIndexes Index, 120, 0, 2, 3, 1
    If Index = 4 Then Form_Resize
End Sub

Public Sub MoveIndexes(Start As Integer, Sep As Long, ParamArray Indexes() As Variant)
    Dim temp As Long, Enabled As Boolean
    For temp = 0 To UBound(Indexes)
        If Indexes(temp) = Start Then
            Enabled = True
        Else
            If Enabled Then
                With XPWmain(CInt(Indexes(temp - 1)))
                    XPWmain(CInt(Indexes(temp))).Top = .Top + .Height + Sep
                End With
            End If
        End If
    Next
End Sub

Private Sub Form_Resize()
    If Me.ScaleWidth > 3000 And Me.ScaleHeight > XPWmain(4).Top + XPWmain(4).Height + 100 Then
        lstview.Move lstview.Left, XPWmain(4).Top + XPWmain(4).Height, Me.ScaleWidth - 2970, Me.ScaleHeight - 225 - XPWmain(4).Height
        XPWmain(4).Width = lstview.Width
        lblhelp(4).Left = XPWmain(4).Width - 615
    End If
End Sub

Private Sub txtindex_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 0, 1
           If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
    End Select
End Sub

Private Sub lstview_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next 'Effect 1 list 7 filenames
    Dim temp As Integer, tempstr As String
    If Data.Files.count > 0 Then
        Filmain.Pattern = txtindex(5)
        For temp = 1 To Data.Files.count
            If (GetAttr(Data.Files(temp)) And vbDirectory) <> vbDirectory Then
                tempstr = LCase(Data.Files.Item(temp))
                lstview.ListItems.Add , tempstr, GetFilename(Data.Files.Item(temp)), , GetIcon(tempstr, imlmain, picmain)
            Else
                Filmain.Path = Data.Files.Item(temp)
                SeedListView Filmain, lstview, imlmain, picmain
            End If
        Next
    End If
    ResetList lstview
    autosizeall lstview
End Sub
