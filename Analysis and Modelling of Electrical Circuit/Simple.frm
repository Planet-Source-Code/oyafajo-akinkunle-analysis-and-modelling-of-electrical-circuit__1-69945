VERSION 5.00
Begin VB.Form SimpleCircuit 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Circuit"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6300
   Icon            =   "Simple.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   6300
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   13
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton ClearAns 
      Caption         =   "Clear Answers"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   12
      Top             =   5640
      Width           =   2055
   End
   Begin VB.CommandButton AnsRes 
      Caption         =   "Answer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton AnsCur 
      Caption         =   "Answer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton AnsVol 
      Caption         =   "Answer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   6
      Top             =   3240
      Width           =   1455
   End
   Begin VB.TextBox Resistance 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox Current 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   4080
      Width           =   2295
   End
   Begin VB.TextBox Voltage 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   3
      Top             =   3240
      Width           =   2295
   End
   Begin VB.CheckBox CheckRes 
      BackColor       =   &H80000004&
      Caption         =   "Resistance (R)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CheckBox CheckCur 
      BackColor       =   &H80000004&
      Caption         =   "Current (I)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CheckBox CheckVol 
      BackColor       =   &H80000004&
      Caption         =   "Voltage (V)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SIMPLE CIRCUIT"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   120
      Width           =   6375
   End
   Begin VB.Line Line23 
      BorderWidth     =   4
      X1              =   0
      X2              =   10080
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   9135
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   1800
      X2              =   1920
      Y1              =   1920
      Y2              =   2040
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   1680
      X2              =   1800
      Y1              =   2040
      Y2              =   1920
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   1920
      Width           =   255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3720
      TabIndex        =   10
      Top             =   1440
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3360
      TabIndex        =   9
      Top             =   2280
      Width           =   255
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   4560
      X2              =   4560
      Y1              =   1320
      Y2              =   2640
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   3720
      X2              =   4560
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   1800
      X2              =   2640
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   1800
      X2              =   1800
      Y1              =   2640
      Y2              =   1320
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   255
      Left            =   2640
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   3240
      X2              =   3240
      Y1              =   2400
      Y2              =   2880
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   3120
      X2              =   3120
      Y1              =   2280
      Y2              =   3000
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   3240
      X2              =   4560
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   1800
      X2              =   3120
      Y1              =   2640
      Y2              =   2640
   End
End
Attribute VB_Name = "SimpleCircuit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ResAnswer As Long
Dim CurAnswer As Long
Dim VolAnswer As Long
Const vbkeyDecPt = 46

Private Sub AnsCur_Click() 'works out the current
Current.Visible = True
AnsCur.Visible = False
CurAnswer = Val(Voltage.Text) / Val(Resistance.Text)
Current.Text = CurAnswer
End Sub

Private Sub AnsRes_Click() 'works out the resistance
Resistance.Visible = True
AnsRes.Visible = False
ResAnswer = Val(Voltage.Text) / Val(Current.Text)
Resistance.Text = ResAnswer
End Sub

Private Sub AnsVol_Click() 'works out the voltage
Voltage.Visible = True
AnsVol.Visible = False
VolAnswer = Val(Current.Text) * Val(Resistance.Text)
Voltage.Text = VolAnswer
End Sub

Private Sub CheckCur_Click() 'displays the current textbox

If CheckCur.Value = Checked Then
Current.Visible = True
Current.SetFocus
    
    If CheckVol.Value = Checked Then
    AnsRes.Visible = True
    End If
    
    If CheckRes.Value = Checked Then
    AnsVol.Visible = True
    End If
    
End If
End Sub

Private Sub CheckRes_Click() 'displays the resistance textbox

If CheckRes.Value = Checked Then
Resistance.Visible = True
Resistance.SetFocus
    
    If CheckVol.Value = Checked Then
    AnsCur.Visible = True
    End If
    
    If CheckCur.Value = Checked Then
    AnsVol.Visible = True
    End If
    
End If

End Sub

Private Sub CheckVol_Click() 'displays the voltage texbox

If CheckVol.Value = Checked Then
Voltage.Visible = True
Voltage.SetFocus
    
    If CheckCur.Value = Checked Then
    AnsRes.Visible = True
    End If
    
    If CheckRes.Value = Checked Then
    AnsCur.Visible = True
    End If
    
End If


End Sub

Private Sub ClearAns_Click()
clear
End Sub

Private Sub Command1_Click()
Unload Me
Splash.Show
End Sub

Private Sub exit_Click()
End
End Sub
Private Sub Current_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub Form_Load()
clear 'calls the clear subroutine
End Sub

Private Sub clear() ' clears all the textboxes, checkboxes and makes the buttons invisible
Voltage.Visible = False
Current.Visible = False
Resistance.Visible = False

AnsVol.Visible = False
AnsCur.Visible = False
AnsRes.Visible = False

CheckCur.Value = Unchecked
CheckVol.Value = Unchecked
CheckRes.Value = Unchecked

Voltage.Text = ""
Current.Text = ""
Resistance.Text = ""
End Sub

Private Sub MainMenu_Click()
clear
Splash.Show
Unload SimpleCircuit
End Sub

Private Sub ParallelC_Click()
clear
ParallelCircuit.Show
Unload SimpleCircuit
End Sub

Private Sub SeriesC_Click()
clear
SeriesCircuit.Show
Unload SimpleCircuit
End Sub

Private Sub AboutO_Click()
About.Show
End Sub
Private Sub Resistance_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub
Private Sub Voltage_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub
