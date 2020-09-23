VERSION 5.00
Begin VB.Form RLCSeries 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AC Circuit -  R, L, C in Series"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9120
   Icon            =   "RLCSeries.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   9120
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
      Left            =   6480
      TabIndex        =   28
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Frame Frame3 
      Caption         =   "Series RLC Circuit Diagram"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   3375
      Left            =   120
      TabIndex        =   27
      Top             =   5160
      Width           =   8775
      Begin VB.Image Image1 
         Height          =   2835
         Left            =   240
         Picture         =   "RLCSeries.frx":62A82
         Top             =   360
         Width           =   5685
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Calculated Results"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   8775
      Begin VB.Label LblVL 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   6240
         TabIndex        =   21
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label LblVC 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   6240
         TabIndex        =   20
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label LblVR 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   6240
         TabIndex        =   19
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label LblRF40 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2880
         TabIndex        =   18
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label LblI40 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2880
         TabIndex        =   17
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label LblZ40 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2880
         TabIndex        =   16
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "VL:  ="
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
         Index           =   10
         Left            =   5280
         TabIndex        =   12
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "VC:  ="
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
         Index           =   9
         Left            =   5280
         TabIndex        =   11
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "VR:  ="
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
         Index           =   8
         Left            =   5280
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Resonant Frequency:  ="
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
         Index           =   7
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "The Current:  ="
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
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "The Impedance:  ="
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
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parameters"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   8775
      Begin VB.TextBox TextR40 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   6240
         TabIndex        =   26
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox TextL40 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   2160
         TabIndex        =   25
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton calculate 
         Caption         =   "Calculate"
         Default         =   -1  'True
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
         Left            =   4560
         TabIndex        =   24
         Top             =   1560
         Width           =   1455
      End
      Begin VB.CommandButton clear 
         Caption         =   "Clear Values"
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
         Left            =   6480
         TabIndex        =   22
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox TextF40 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   2160
         TabIndex        =   15
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox TextC40 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   6240
         TabIndex        =   14
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox TextV40 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   2160
         TabIndex        =   13
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Frequency:"
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
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Capacitance:"
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
         Index           =   3
         Left            =   4320
         TabIndex        =   4
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Inductance:"
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
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Resistance:"
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
         Index           =   1
         Left            =   4320
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Voltage:"
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
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Line Line2 
      BorderWidth     =   4
      X1              =   0
      X2              =   10080
      Y1              =   8640
      Y2              =   8640
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   0
      X2              =   10080
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SERIES RLC CIRCUIT"
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
      TabIndex        =   23
      Top             =   120
      Width           =   6375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   9135
   End
End
Attribute VB_Name = "RLCSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim V40 As Single
Dim R40 As Single
Dim L40 As Single
Dim C40 As Single
Dim F40 As Single
Dim XL40 As Single
Dim RF40 As Single
Dim XC40 As Single
Dim I40 As Single
Dim Z40 As Single
Dim VR As Single
Dim VC As Single
Dim VL As Single
Const vbkeyDecPt = 46
Const pi = 22 / 7

Private Sub calculate_Click()
V40 = Val(TextV40.Text)
R40 = Val(TextR40.Text)
L40 = Val(TextL40.Text)
C40 = Val(TextC40.Text)
F40 = Val(TextF40.Text)
If L40 = 0 Then
XL40 = 0
Else
XL40 = (2 * pi * F40 * L40)
End If
If C40 = 0 Then
XC40 = 0
Else
XC40 = 1 / (2 * pi * F40 * C40)
End If
If L40 = 0 Or C40 = 0 Then
LblRF40.Caption = " "
Else
RF40 = 1 / (2 * pi * (L40 * C40) ^ (1 / 2))
End If
Z40 = (R40 ^ 2 + (XL40 - XC40) ^ 2) ^ (1 / 2)
I40 = V40 / Z40
VR = I40 * R40
VC = I40 * XC40
VL = I40 * XL40

LblZ40.Caption = Format$(Z40, "###.00ohms")
LblI40.Caption = Format$(I40, "###.000000A")
LblRF40.Caption = Format$(RF40, "###.00HZ")
LblVR.Caption = Format$(VR, "###.00V")
LblVC.Caption = Format$(VC, "###.00V")
LblVL.Caption = Format$(VL, "###.00V")
End Sub

Private Sub clear_Click()
TextV40.Text = " "
TextR40.Text = " "
TextL40.Text = " "
TextC40.Text = " "
TextF40.Text = " "
LblZ40.Caption = " "
LblI40.Caption = " "
LblRF40.Caption = " "
LblVR.Caption = " "
LblVC.Caption = " "
LblVL.Caption = " "
End Sub

Private Sub Command1_Click()
Unload Me
Splash.Show
End Sub

Private Sub TextC40_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub
Private Sub TextF40_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub TextL40_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub TextR40_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub TextV40_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub
