VERSION 5.00
Begin VB.Form SeriesCircuit 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Series Circuit"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8835
   Icon            =   "SeriesCircuit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   8835
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
      Height          =   495
      Left            =   3240
      TabIndex        =   17
      Top             =   9240
      Width           =   2175
   End
   Begin VB.PictureBox Results 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   840
      ScaleHeight     =   3075
      ScaleWidth      =   7035
      TabIndex        =   11
      Top             =   5880
      Width           =   7095
   End
   Begin VB.CommandButton Repeat 
      Caption         =   "Repeat"
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
      Left            =   840
      TabIndex        =   10
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton ResValEnter 
      Caption         =   "Enter"
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
      Left            =   5640
      TabIndex        =   7
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox CountUp 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox VoltVal 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox HMR 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SERIES CIRCUIT"
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
      Left            =   1320
      TabIndex        =   18
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
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Caption         =   "Max of 10"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   16
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Line Line19 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      X1              =   6000
      X2              =   6480
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line18 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      X1              =   4800
      X2              =   4800
      Y1              =   2040
      Y2              =   1680
   End
   Begin VB.Line Line17 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      X1              =   6480
      X2              =   6480
      Y1              =   2040
      Y2              =   1680
   End
   Begin VB.Line Line16 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      X1              =   4800
      X2              =   5280
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line15 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      X1              =   4080
      X2              =   4080
      Y1              =   2040
      Y2              =   1680
   End
   Begin VB.Line Line14 
      BorderColor     =   &H000000C0&
      BorderWidth     =   2
      X1              =   3600
      X2              =   4080
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line13 
      BorderColor     =   &H000000C0&
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      X1              =   2400
      X2              =   2880
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line12 
      BorderColor     =   &H000000C0&
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      X1              =   2400
      X2              =   2400
      Y1              =   1680
      Y2              =   2040
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   1440
      TabIndex        =   15
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Caption         =   "Vs"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   4680
      TabIndex        =   14
      Top             =   3120
      Width           =   495
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Caption         =   "R2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000006&
      Height          =   255
      Left            =   5400
      TabIndex        =   13
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Caption         =   "R1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   3000
      TabIndex        =   12
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Caption         =   "V2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   5400
      TabIndex        =   9
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000004&
      Caption         =   "V1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   255
      Left            =   3000
      TabIndex        =   8
      Top             =   1905
      Width           =   495
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   1920
      X2              =   2040
      Y1              =   2280
      Y2              =   2400
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   1800
      X2              =   1920
      Y1              =   2400
      Y2              =   2280
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   4560
      X2              =   6960
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   6960
      X2              =   6960
      Y1              =   1560
      Y2              =   3000
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   6240
      X2              =   6960
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   1920
      X2              =   2640
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   1920
      X2              =   1920
      Y1              =   3000
      Y2              =   1560
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   4440
      X2              =   1920
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   4560
      X2              =   4560
      Y1              =   2760
      Y2              =   3240
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   4440
      X2              =   4440
      Y1              =   2640
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   3840
      X2              =   5040
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   255
      Left            =   5040
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   255
      Left            =   2640
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Number 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
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
      TabIndex        =   6
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000004&
      Caption         =   "Value of Resistor  ="
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
      Left            =   720
      TabIndex        =   4
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000004&
      Caption         =   "Supply Voltage (Vs)  ="
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
      Left            =   720
      TabIndex        =   2
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "How Many Resistors  ="
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
      Left            =   720
      TabIndex        =   1
      Top             =   3720
      Width           =   2775
   End
End
Attribute VB_Name = "SeriesCircuit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Counter As Integer 'variables
Dim Max As Integer
Dim Res(10) As Single
Dim Vol(10) As Single
Dim it As Single
Dim VS As Single
Dim rt As Single
Dim num As Integer
Const vbkeyDecPt = 46

Private Sub Command1_Click()
Unload Me
Splash.Show
End Sub

Private Sub exit_Click()
End
End Sub
Private Sub CountUp_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub Form_Load()
Counter = 1
Results.Visible = False
rt = 0
End Sub

Private Sub MainMenu_Click()
Splash.Show
Unload SeriesCircuit
End Sub
Private Sub HMR_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub Repeat_Click() 'clears all the values from picture and text boxes
Unload SeriesCircuit
Load SeriesCircuit
SeriesCircuit.Show
End Sub

Private Sub ResValEnter_Click()

If Val(HMR.Text) = 0 Then
MsgBox ("Cannot Have 0 Resistors")
GoTo 1
End If

If CountUp.Text = "" Then 'if there is no value in the texbox
MsgBox ("PLEASE ENTER A VALUE FOR THIS RESISTOR") 'then display this message
GoTo 1 'skips the rest of the rountine and jumps straight to the end of the sub
End If

CountUp.SetFocus 'resets focus back to the textbox

If Val(Number.Caption) = Val(HMR.Text) Then 'if all the values of the resistors have been entered
ResValEnter.Visible = False 'then the enter button disappears
Number.Visible = False
End If

Max = Val(HMR.Text) + 1 'max equals the total resistors in the circuit plus 1

Res(Counter) = Val(CountUp.Text) 'res(1,2,3 etc..) is equal to the entered value
CountUp.Text = "" 'clears the box that had the value of the last resistor in it

If Counter = Max - 1 Then
For num = 1 To Max - 1 'FOR loop
rt = rt + Res(num) 'works out the resistance total
Next num 'start the loop again

Results.Visible = True 'show the picture box with the results

Results.Print "Resistance Total = "; rt; "Ohms" 'print the total resistance to the picturebox
it = Val(VoltVal.Text) / rt 'work out the total current
Results.Print "Total Current ="; it; "Amps" 'print the total current to the picturebox

For num = 1 To Max - 1
Vol(num) = it * Res(num) 'works out the voltages across each resistor
Results.Print "Voltage Across R"; num; "="; Vol(num); "Volts" 'print the voltages to the picturebox
Next num
End If

Counter = Counter + 1 'adds 1 to the value of counter
Number.Caption = Counter

If Counter = Max Then
Results.Visible = True 'shows the picturebox again
End If
1 End Sub 'end sub with the label 1 in the margin

Private Sub SimpleC_Click()
SimpleCircuit.Show
Unload SeriesCircuit
End Sub

Private Sub ParallelC_Click()
ParallelCircuit.Show
Unload SeriesCircuit
End Sub

Private Sub AboutO_Click()
About.Show
End Sub
Private Sub VoltVal_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub
