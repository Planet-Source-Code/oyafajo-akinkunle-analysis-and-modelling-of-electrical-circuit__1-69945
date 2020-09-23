VERSION 5.00
Begin VB.Form ParallelCircuit 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parallel Circuit"
   ClientHeight    =   9450
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8490
   Icon            =   "ParallelCircuit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   8490
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
      Left            =   2760
      TabIndex        =   18
      Top             =   8760
      Width           =   2175
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
      Left            =   3360
      TabIndex        =   12
      ToolTipText     =   "How many resistors in parallel"
      Top             =   3480
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
      Left            =   3360
      TabIndex        =   11
      ToolTipText     =   "Circuit Voltage"
      Top             =   3960
      Width           =   1215
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
      Left            =   3360
      TabIndex        =   10
      ToolTipText     =   "Enter value of resistors here"
      Top             =   4440
      Width           =   1215
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
      Left            =   4920
      TabIndex        =   9
      ToolTipText     =   "Click to save the value of the resistor"
      Top             =   4440
      Width           =   1695
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
      Left            =   600
      TabIndex        =   8
      ToolTipText     =   "Start a new circuit calculation"
      Top             =   5040
      Width           =   1695
   End
   Begin VB.PictureBox Results 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
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
      Height          =   3015
      Left            =   600
      ScaleHeight     =   2955
      ScaleWidth      =   7395
      TabIndex        =   7
      ToolTipText     =   "Results"
      Top             =   5640
      Width           =   7455
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
      Left            =   2880
      TabIndex        =   13
      Top             =   4440
      Width           =   495
   End
   Begin VB.Line Line23 
      BorderWidth     =   4
      X1              =   0
      X2              =   10080
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PARALLEL CIRCUIT"
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
      TabIndex        =   19
      Top             =   120
      Width           =   6375
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
      Left            =   4680
      TabIndex        =   17
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000004&
      Caption         =   "How Many Resistors ="
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
      TabIndex        =   16
      Top             =   3480
      Width           =   2655
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000004&
      Caption         =   "Supply Voltage (Vs) ="
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
      TabIndex        =   15
      Top             =   3960
      Width           =   2655
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000004&
      Caption         =   "Value of Resistor ="
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
      TabIndex        =   14
      Top             =   4440
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000004&
      Caption         =   "R3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   6
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label7 
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
      Height          =   255
      Left            =   5280
      TabIndex        =   5
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label6 
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
      Height          =   255
      Left            =   3720
      TabIndex        =   4
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label5 
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
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000004&
      Caption         =   "I3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6120
      TabIndex        =   2
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000004&
      Caption         =   "I2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4560
      TabIndex        =   1
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000004&
      Caption         =   "I1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   0
      Top             =   2640
      Width           =   255
   End
   Begin VB.Line Line22 
      BorderWidth     =   2
      X1              =   6720
      X2              =   6600
      Y1              =   2760
      Y2              =   2880
   End
   Begin VB.Line Line21 
      BorderWidth     =   2
      X1              =   6480
      X2              =   6600
      Y1              =   2760
      Y2              =   2880
   End
   Begin VB.Line Line20 
      BorderWidth     =   2
      X1              =   5160
      X2              =   5040
      Y1              =   2760
      Y2              =   2880
   End
   Begin VB.Line Line19 
      BorderWidth     =   2
      X1              =   4920
      X2              =   5040
      Y1              =   2760
      Y2              =   2880
   End
   Begin VB.Line Line18 
      BorderWidth     =   2
      X1              =   3480
      X2              =   3600
      Y1              =   2880
      Y2              =   2760
   End
   Begin VB.Line Line17 
      BorderWidth     =   2
      X1              =   3360
      X2              =   3480
      Y1              =   2760
      Y2              =   2880
   End
   Begin VB.Line Line16 
      BorderWidth     =   2
      X1              =   6600
      X2              =   6600
      Y1              =   2520
      Y2              =   3120
   End
   Begin VB.Line Line15 
      BorderWidth     =   2
      X1              =   5040
      X2              =   6600
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line14 
      BorderWidth     =   2
      X1              =   6600
      X2              =   6600
      Y1              =   1680
      Y2              =   1200
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   5040
      X2              =   6600
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   855
      Left            =   6480
      Top             =   1680
      Width           =   255
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   5040
      X2              =   5040
      Y1              =   2520
      Y2              =   3120
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   3480
      X2              =   5040
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line10 
      BorderWidth     =   2
      X1              =   5040
      X2              =   5040
      Y1              =   1680
      Y2              =   1200
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   3480
      X2              =   5040
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   855
      Left            =   4920
      Top             =   1680
      Width           =   255
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   3480
      X2              =   3480
      Y1              =   2520
      Y2              =   3120
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   1800
      X2              =   3480
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line6 
      BorderWidth     =   2
      X1              =   1800
      X2              =   1800
      Y1              =   2280
      Y2              =   3120
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   3480
      X2              =   3480
      Y1              =   1680
      Y2              =   1200
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   1800
      X2              =   3480
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   855
      Left            =   3360
      Top             =   1680
      Width           =   255
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   1800
      X2              =   1800
      Y1              =   2160
      Y2              =   1200
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   1560
      X2              =   2040
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1320
      X2              =   2280
      Y1              =   2160
      Y2              =   2160
   End
End
Attribute VB_Name = "ParallelCircuit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Counter As Integer 'Variables
Dim Max As Integer
Dim Res(10) As Single
Dim Cur(10) As Single
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
it = 0
End Sub


Private Sub MainMenu_Click()
Splash.Show
Unload ParallelCircuit
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
Unload ParallelCircuit
Load ParallelCircuit
ParallelCircuit.Show
End Sub

Private Sub ResValEnter_Click()

If Val(HMR.Text) = 0 Then
MsgBox ("Cannot Have 0 Resistors")
GoTo 1
End If

If CountUp.Text = "" Then 'if there is no value in the textbox
MsgBox ("PLEASE ENTER A VALUE FOR THIS RESISTOR") 'then display this message
GoTo 1 'skips the rest of the rountine and jumps straight to the end of the sub
End If

CountUp.SetFocus 'resets focus back to the textbox

If Val(Number.Caption) = Val(HMR.Text) Then 'if all the values of the resistors have been entered
ResValEnter.Visible = False 'then the enter button disappears
Number.Visible = False 'and the resistor number label disappears
End If

Max = Val(HMR.Text) + 1 'max equals the nnumber of resistors in the circuit plus 1

Res(Counter) = Val(CountUp.Text) 'res(1,2,3 etc..) is equal to the entered value
CountUp.Text = "" 'clears the textbox

If Counter = Max - 1 Then
rt = (Res(1) * Res(2)) / (Res(1) + Res(2)) 'works out rt for the for the first 2 resistors
For num = 1 To Max - 3
rt = (rt * Res(num)) / (rt + Res(num)) 'works out rt for more than 2 resistors
Next num

Results.Visible = True

Results.Print "Resistance Total = "; rt; "Ohms"

For num = 1 To Max - 1
Cur(num) = Val(VoltVal.Text) / Res(num) 'works out the current in each branch
Results.Print "Current"; num; "="; Cur(num); "Amps" 'show each branch current in the picture box
it = it + Cur(num) 'works out the current total
Next num
Results.Print "Current Total ="; it; "Amps" 'total current in the picture box
End If

Counter = Counter + 1 'adds 1 to the value of counter
Number.Caption = Counter

If Counter = Max Then
Results.Visible = True
End If

1 End Sub 'end sub with the label 1 in the margin

Private Sub SeriesC_Click()
SeriesCircuit.Show
Unload ParallelCircuit
End Sub

Private Sub SimpleC_Click()
SimpleCircuit.Show
Unload ParallelCircuit
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
