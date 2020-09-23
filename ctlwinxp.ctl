VERSION 5.00
Begin VB.UserControl XPWin 
   BackColor       =   &H00F7DFD6&
   ClientHeight    =   2250
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2625
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   2250
   ScaleWidth      =   2625
   ToolboxBitmap   =   "ctlwinxp.ctx":0000
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2040
      Top             =   1680
   End
   Begin VB.Label lblmain 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      ForeColor       =   &H00C65D21&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   585
   End
   Begin VB.Image imgborder 
      Height          =   45
      Index           =   2
      Left            =   0
      Picture         =   "ctlwinxp.ctx":0312
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   2565
   End
   Begin VB.Image imgborder 
      Height          =   1575
      Index           =   1
      Left            =   2520
      Picture         =   "ctlwinxp.ctx":0349
      Stretch         =   -1  'True
      Top             =   480
      Width           =   30
   End
   Begin VB.Image imgborder 
      Height          =   1575
      Index           =   0
      Left            =   0
      Picture         =   "ctlwinxp.ctx":0380
      Stretch         =   -1  'True
      Top             =   480
      Width           =   45
   End
   Begin VB.Image imgbutton 
      Height          =   285
      Left            =   2040
      Picture         =   "ctlwinxp.ctx":03B7
      Tag             =   "0"
      Top             =   960
      Width           =   285
   End
   Begin VB.Image imgstate 
      Height          =   285
      Index           =   1
      Left            =   2160
      Picture         =   "ctlwinxp.ctx":083B
      ToolTipText     =   "Up"
      Top             =   480
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imgstate 
      Height          =   285
      Index           =   0
      Left            =   1800
      Picture         =   "ctlwinxp.ctx":0CBA
      ToolTipText     =   "Down"
      Top             =   480
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Image imghead 
      Height          =   375
      Index           =   3
      Left            =   2520
      Picture         =   "ctlwinxp.ctx":113E
      Top             =   0
      Width           =   30
   End
   Begin VB.Image imghead 
      Height          =   375
      Index           =   0
      Left            =   0
      Picture         =   "ctlwinxp.ctx":1183
      Top             =   0
      Width           =   45
   End
   Begin VB.Image imghead 
      Height          =   375
      Index           =   1
      Left            =   120
      Picture         =   "ctlwinxp.ctx":11C8
      Stretch         =   -1  'True
      Top             =   0
      Width           =   735
   End
   Begin VB.Image imghead 
      Height          =   375
      Index           =   2
      Left            =   960
      Picture         =   "ctlwinxp.ctx":1209
      Top             =   0
      Width           =   1470
   End
End
Attribute VB_Name = "XPWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Const onepixel As Byte = 15
Dim maxheight As Long
Dim minheight As Long
Const babyblue = "&H00F7DFD6&"
Private speed As Long
Private Smooth As Boolean
Const vbdarkblue = &HC65D21
Const vblightblue = 16748098
Public Event Click()
Public Event Resize()
Public Event MouseMove()
Public Event ChangeOver(State As Boolean)

Public Property Let Caption(text As String)
    lblmain = text
End Property
Public Property Get Caption() As String
    Caption = lblmain
End Property
Public Property Get Backcolor() As String
    Backcolor = babyblue
End Property
Public Property Get State() As Long
    State = imgbutton.Tag
End Property
Public Property Let State(ByVal newstate As Long)
    If newstate <> imgbutton.Tag Then imgbutton_Click
End Property
Public Property Let DragSpeed(ByVal newspeed As Byte)
    speed = newspeed
End Property
Public Property Get DragSpeed() As Byte
    DragSpeed = speed
End Property

Public Property Let SmoothDrag(ByVal dragtype As Boolean)
    Smooth = dragtype
End Property
Public Property Get SmoothDrag() As Boolean
    SmoothDrag = Smooth
End Property

Public Sub imgbutton_Click()
If imgbutton.Tag = 0 Then
    imgbutton.Tag = 1
Else
    imgbutton.Tag = 0
End If
RaiseEvent Click
imgbutton.picture = imgstate(imgbutton.Tag).picture
Timer.Enabled = True
End Sub

Private Sub imgbutton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblmain_MouseMove Button, Shift, X, Y
End Sub
Private Sub imghead_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
unselect
End Sub

Private Sub imgstate_Click(Index As Integer)
    imgbutton_Click
End Sub

Private Sub lblmain_Click()
    imgbutton_Click
End Sub

Private Sub lblmain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not lblmain.Font.Underline Then
        lblmain.Font.Underline = True
        lblmain.ForeColor = vblightblue
        RaiseEvent ChangeOver(True)
    End If
End Sub

Private Sub Timer_Timer()
DoEvents
If Smooth = True Then
If imgbutton.Tag = 0 Then   'down
    If UserControl.Height < maxheight Then
        If UserControl.Height + (onepixel * speed) > maxheight Then
            UserControl.Height = maxheight
            Timer.Enabled = False
        Else
            UserControl.Height = UserControl.Height + (onepixel * speed)
        End If
    End If
Else                        'up
    If UserControl.Height > minheight Then
        If UserControl.Height - (onepixel * speed) < minheight Then
            UserControl.Height = minheight
            Timer.Enabled = False
        Else
            UserControl.Height = UserControl.Height - (onepixel * speed)
        End If
    End If
End If
Else
If imgbutton.Tag = 0 Then   'down
    UserControl.Height = maxheight
    Timer.Enabled = False
Else                        'up
    UserControl.Height = minheight
    Timer.Enabled = False
End If
End If
DoEvents
End Sub

Private Sub UserControl_Initialize()
maxheight = UserControl.Height
UserControl_Resize
End Sub

Private Sub UserControl_InitProperties()
minheight = imghead(0).Height
maxheight = minheight
speed = 15
Smooth = False
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    unselect
    RaiseEvent MouseMove
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    SmoothDrag = PropBag.ReadProperty("SmoothDrag", False)
    speed = PropBag.ReadProperty("Speed", 1)
    State = PropBag.ReadProperty("State", 0)
    Caption = PropBag.ReadProperty("Caption", UserControl.Name)
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
If minheight = 0 Then minheight = imghead(0).Height
If UserControl.Height < minheight Then UserControl.Height = minheight
If UserControl.Height > minheight Then maxheight = UserControl.Height
imghead(0).Move 0, 0
imghead(1).Move imghead(0).Width, 0, UserControl.Width - imghead(0).Width - imghead(2).Width - imghead(3).Width
imghead(2).Move imghead(1).Left + imghead(1).Width, 0
imghead(3).Move imghead(2).Left + imghead(2).Width, 0
imgbutton.Move imghead(2).Left + imghead(2).Width - imgbutton.Width - 21, imghead(0).Height / 2 - imgbutton.Height / 2
imgborder(0).Move 0, imghead(0).Height, onepixel, UserControl.Height - imghead(0).Height
imgborder(1).Move UserControl.Width - onepixel, imghead(0).Height, onepixel, UserControl.Height - imghead(0).Height
'picmain.Move onepixel, imghead(0).Height, UserControl.Width - (onepixel * 2)
'If UserControl.Height > minheight Then picmain.Height = UserControl.Height - onepixel - imghead(0).Height
imgborder(2).Move 0, UserControl.Height - onepixel, UserControl.Width, onepixel
RaiseEvent Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "SmoothDrag", Smooth, False
    PropBag.WriteProperty "State", State, 0
    PropBag.WriteProperty "Caption", lblmain.Caption, UserControl.Name
    PropBag.WriteProperty "Speed", speed, 1
End Sub

Private Sub unselect()
    If lblmain.Font.Underline Then
        lblmain.ForeColor = vbdarkblue
        lblmain.Font.Underline = False
        RaiseEvent ChangeOver(False)
    End If
End Sub
