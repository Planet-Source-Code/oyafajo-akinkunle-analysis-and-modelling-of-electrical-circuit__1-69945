VERSION 5.00
Begin VB.Form RLCTwoLoop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AC Circuit -  two loop Circuit"
   ClientHeight    =   12015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9720
   Icon            =   "RLCTwoLoop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   12015
   ScaleWidth      =   9720
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      Caption         =   "Circuit Analysis"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3975
      Left            =   120
      TabIndex        =   45
      Top             =   7920
      Width           =   9375
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Nodal Circuit"
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
         Left            =   4920
         TabIndex        =   47
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Mesh Circuit"
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
         Left            =   720
         TabIndex        =   46
         Top             =   240
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   2955
         Left            =   120
         Picture         =   "RLCTwoLoop.frx":62A82
         Top             =   480
         Width           =   7035
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Analysis | Calculated"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3375
      Left            =   120
      TabIndex        =   25
      Top             =   4440
      Width           =   9375
      Begin VB.Frame Frame8 
         Caption         =   "Nodal Voltage Calculated"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   4200
         TabIndex        =   27
         Top             =   1920
         Width           =   4695
         Begin VB.Label LblVA 
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
            Height          =   255
            Left            =   1080
            TabIndex        =   33
            Top             =   600
            Width           =   2175
         End
         Begin VB.Label Label17 
            Caption         =   "VA:  ="
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
            Height          =   255
            Left            =   240
            TabIndex        =   30
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Mesh Current Calculated"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1335
         Left            =   4200
         TabIndex        =   26
         Top             =   480
         Width           =   4695
         Begin VB.Label LblI52 
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
            Height          =   255
            Left            =   1080
            TabIndex        =   32
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label LblI51 
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
            Height          =   255
            Left            =   1080
            TabIndex        =   31
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label Label16 
            Caption         =   "I2:  ="
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
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   840
            Width           =   735
         End
         Begin VB.Label Label15 
            Caption         =   "I1:  ="
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
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   480
            Width           =   735
         End
      End
      Begin VB.Image Image2 
         Height          =   2850
         Left            =   600
         Picture         =   "RLCTwoLoop.frx":664B0
         Top             =   360
         Width           =   2925
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Analysis"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   9375
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
         Left            =   7080
         TabIndex        =   48
         Top             =   3120
         Width           =   1815
      End
      Begin VB.CommandButton calculate 
         Caption         =   "Calculate"
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
         TabIndex        =   44
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Frame Frame4 
         Caption         =   "Arm 3"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   6360
         TabIndex        =   35
         Top             =   1320
         Width           =   2535
         Begin VB.TextBox TextR52 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   600
            TabIndex        =   39
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox TextC52 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   600
            TabIndex        =   38
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox TextL52 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   600
            TabIndex        =   37
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox TextV52 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   600
            TabIndex        =   36
            Top             =   1320
            Width           =   1695
         End
         Begin VB.Label Label11 
            Caption         =   "R3:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label12 
            Caption         =   "C3:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label13 
            Caption         =   "L3:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label14 
            Caption         =   "V3:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   1320
            Width           =   735
         End
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
         Height          =   375
         Left            =   5040
         TabIndex        =   34
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Frame Frame5 
         Caption         =   "Select Your Method of Analysis"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   8655
         Begin VB.OptionButton Option2 
            Caption         =   "Nodal Analysis"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4320
            TabIndex        =   15
            Top             =   360
            Width           =   2895
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Mesh Analysis"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   600
            TabIndex        =   14
            Top             =   360
            Width           =   2895
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Arm 2"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   3480
         TabIndex        =   3
         Top             =   1320
         Width           =   2535
         Begin VB.TextBox TextV51 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   600
            TabIndex        =   24
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox TextL51 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   600
            TabIndex        =   23
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox TextC51 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   600
            TabIndex        =   22
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox TextR51 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   600
            TabIndex        =   21
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label10 
            Caption         =   "V2:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "L2:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label8 
            Caption         =   "C2:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "R2:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Arm 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   240
         TabIndex        =   2
         Top             =   1320
         Width           =   2535
         Begin VB.TextBox Textf5 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   600
            TabIndex        =   20
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox TextV50 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   600
            TabIndex        =   19
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox TextL50 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   600
            TabIndex        =   18
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox TextC50 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   600
            TabIndex        =   17
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox TextR50 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   600
            TabIndex        =   16
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label6 
            Caption         =   "f:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "V1:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "L1:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "C1:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "R1:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   735
         End
      End
   End
   Begin VB.Line Line2 
      BorderWidth     =   4
      X1              =   0
      X2              =   10080
      Y1              =   12000
      Y2              =   12000
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
      Caption         =   "TWO LOOP R, L, C CIRCUIT"
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
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "RLCTwoLoop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim R5(3) As Single
Dim V5(3) As Single
Dim L5(3) As Single
Dim C5(3) As Single
Dim f5 As Single
Dim n As Integer
Dim Z1, Z2, Z3 As Single
Dim D, D1, D2 As Single
Dim Z12, Z23, Z21, V12, V32 As Single
Dim I51, I52, VA As Single
Dim XL5(3) As Single
Dim XC5(3) As Single
Dim X, Y As Single
Const vbkeyDecPt = 46
Const pi = 22 / 7

Private Sub calculate_Click()
If TextR50.Text = " " Then
MsgBox "Please Enter The Resistance 1"
ElseIf TextR51.Text = " " Then
MsgBox "Please Enter The Resistance 2"
ElseIf TextR52.Text = " " Then
MsgBox "Please Enter The Resistance 3"
End If
R5(0) = Val(TextR50.Text)
R5(1) = Val(TextR51.Text)
R5(2) = Val(TextR52.Text)
L5(0) = Val(TextL50.Text)
L5(1) = Val(TextL51.Text)
L5(2) = Val(TextL52.Text)
C5(0) = Val(TextC50.Text)
C5(1) = Val(TextC51.Text)
C5(2) = Val(TextC52.Text)
V5(0) = Val(TextV50.Text)
V5(1) = Val(TextV51.Text)
V5(2) = Val(TextV52.Text)
f5 = Val(Textf5.Text)
For n = 0 To 2
If L5(n) = 0 Then
XL5(n) = 0
Else
XL5(n) = 2 * pi * f5 * L5(n)
End If
Next n
For n = 0 To 2
If C5(n) = 0 Or f5 = 0 Then
XC5(n) = 0
Else
XC5(n) = 1 / (2 * pi * f5 * C5(n))
End If
Next n
Z1 = (R5(0) ^ 2 + (XL5(0) - XC5(0)) ^ 2) ^ (1 / 2)
Z2 = (R5(1) ^ 2 + (XL5(1) - XC5(1)) ^ 2) ^ (1 / 2)
Z3 = (R5(2) ^ 2 + (XL5(2) - XC5(2)) ^ 2) ^ (1 / 2)
Z12 = Z1 + Z2
Z23 = -(Z2 + Z3)
Z21 = -Z2
V12 = V5(0) - V5(1)
V32 = V5(2) - V5(1)

If Option1.Value = True Then
D = ((Z12 * Z23) + (Z21) ^ 2)
D1 = ((Z23 * V12) - (Z21 * V32))
D2 = (Z12 * V32) + (Z21 * V12)
I51 = D1 / D
I52 = D2 / D
LblI51.Caption = Format$(I51, "###.000000A")
LblI52.Caption = Format$(I52, "###.000000A")
LblVA.Caption = " "
ElseIf Option2.Value = True Then
Y = ((V5(0) * Z2 * Z3) + (V5(2) * Z1 * Z2) + (V5(1) * Z1 * Z3))
X = ((Z2 * Z3) + (Z1 * Z3) + (Z1 * Z2))
VA = Y / X
LblVA.Caption = Format$(VA, "###.000000V")
LblI51.Caption = " "
LblI52.Caption = " "
End If
End Sub

Private Sub clear_Click()
TextR50.Text = " "
TextR51.Text = " "
TextR52.Text = " "
TextL50.Text = " "
TextL51.Text = " "
TextL52.Text = " "
TextC50.Text = " "
TextC51.Text = " "
TextC52.Text = " "
TextV50.Text = " "
TextV51.Text = " "
TextV52.Text = " "
LblVA.Caption = " "
LblI51.Caption = " "
LblI52.Caption = " "
End Sub

Private Sub Command1_Click()
Unload Me
Splash.Show
End Sub



Private Sub TextC50_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub
Private Sub TextC51_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub


Private Sub TextC52_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub


Private Sub Textf5_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub


Private Sub TextL50_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub


Private Sub TextL51_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub


Private Sub TextL52_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub


Private Sub TextR50_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub


Private Sub TextR51_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub


Private Sub TextR52_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub


Private Sub TextV50_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub


Private Sub TextV51_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub


Private Sub TextV52_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub
