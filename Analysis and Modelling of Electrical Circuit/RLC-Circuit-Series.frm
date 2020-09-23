VERSION 5.00
Begin VB.Form ParallelCircuit 
   BackColor       =   &H00FF7D04&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parallel Circuit"
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8655
   Icon            =   "RLC-Circuit-Series.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   8655
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3000
      TabIndex        =   13
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
      Left            =   3000
      TabIndex        =   12
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
      Left            =   3000
      TabIndex        =   11
      ToolTipText     =   "Enter value of resistors here"
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton ResValEnter 
      Caption         =   "Enter"
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
      Left            =   4560
      TabIndex        =   10
      ToolTipText     =   "Click to save the value of the resistor"
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton Repeat 
      Caption         =   "Repeat"
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
      Left            =   600
      TabIndex        =   9
      ToolTipText     =   "Start a new circuit calculation"
      Top             =   5040
      Width           =   1695
   End
   Begin VB.PictureBox Results 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   600
      ScaleHeight     =   2955
      ScaleWidth      =   7395
      TabIndex        =   8
      ToolTipText     =   "Results"
      Top             =   5640
      Width           =   7455
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FF7D04&
      Caption         =   "Max of 10"
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
      Left            =   4320
      TabIndex        =   18
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FF7D04&
      Caption         =   "How Many Resistors"
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
      Left            =   600
      TabIndex        =   17
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FF7D04&
      Caption         =   "Supply Voltage (Vs)"
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
      Left            =   600
      TabIndex        =   16
      Top             =   3960
      Width           =   2295
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF7D04&
      Caption         =   "Value of Resistor"
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
      Left            =   600
      TabIndex        =   15
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Number 
      Alignment       =   2  'Center
      BackColor       =   &H00FF7D04&
      Caption         =   "1"
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
      Left            =   2400
      TabIndex        =   14
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FF7D04&
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
      TabIndex        =   7
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FF7D04&
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
      TabIndex        =   6
      Top             =   1800
      Width           =   375
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF7D04&
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
      TabIndex        =   5
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FF7D04&
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
      TabIndex        =   4
      Top             =   2400
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF7D04&
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
      TabIndex        =   3
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF7D04&
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
      TabIndex        =   2
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF7D04&
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
      TabIndex        =   1
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF7D04&
      Caption         =   "Parallel Circuit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8415
   End
   Begin VB.Menu menu 
      Caption         =   "Menu"
      Begin VB.Menu mainmenu 
         Caption         =   "Main Menu"
         Shortcut        =   ^M
      End
      Begin VB.Menu SimpleC 
         Caption         =   "Simple Circuit"
         Shortcut        =   ^S
      End
      Begin VB.Menu SeriesC 
         Caption         =   "Series Circuit"
         Shortcut        =   ^E
      End
      Begin VB.Menu AboutO 
         Caption         =   "About"
         Shortcut        =   ^A
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
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

Private Sub Exit_Click()
End
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
