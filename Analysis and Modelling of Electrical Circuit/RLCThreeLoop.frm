VERSION 5.00
Begin VB.Form RLCThreeLoop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AC Circuit -  three loop Circuit"
   ClientHeight    =   10290
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10320
   Icon            =   "RLCThreeLoop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10290
   ScaleWidth      =   10320
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
      Left            =   8760
      TabIndex        =   69
      Top             =   8040
      Width           =   1455
   End
   Begin VB.Frame Frame11 
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
      Height          =   3735
      Left            =   120
      TabIndex        =   54
      Top             =   6240
      Width           =   10095
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
         Left            =   5880
         TabIndex        =   56
         Top             =   360
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
         Left            =   840
         TabIndex        =   55
         Top             =   360
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   2910
         Left            =   120
         Picture         =   "RLCThreeLoop.frx":62A82
         Top             =   600
         Width           =   8265
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
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   10095
      Begin VB.Frame Frame8 
         Caption         =   "Nodal Voltages Calculated"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   6840
         TabIndex        =   64
         Top             =   3360
         Width           =   2895
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
            Left            =   960
            TabIndex        =   68
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label Label17 
            Caption         =   "VA:  ="
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
            TabIndex        =   67
            Top             =   360
            Width           =   735
         End
         Begin VB.Label LblVB 
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
            Left            =   960
            TabIndex        =   66
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label RLCThreeLoop 
            Caption         =   "VB:  ="
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
            TabIndex        =   65
            Top             =   720
            Width           =   735
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Mesh Currents"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1575
         Left            =   4080
         TabIndex        =   57
         Top             =   3360
         Width           =   2535
         Begin VB.Label LblI2 
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
            Left            =   960
            TabIndex        =   63
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label LblI1 
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
            Left            =   960
            TabIndex        =   62
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label16 
            Caption         =   "I2:  ="
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
            TabIndex        =   61
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label15 
            Caption         =   "I1:  ="
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
            TabIndex        =   60
            Top             =   360
            Width           =   735
         End
         Begin VB.Label LblI3 
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
            Left            =   960
            TabIndex        =   59
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label20 
            Caption         =   "I3:  ="
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
            TabIndex        =   58
            Top             =   1080
            Width           =   735
         End
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
         Left            =   2400
         TabIndex        =   53
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Frame Frame10 
         Caption         =   "Arm 5"
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
         Left            =   4680
         TabIndex        =   44
         Top             =   1320
         Width           =   2055
         Begin VB.TextBox TextR5 
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
            TabIndex        =   48
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox TextC5 
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
            TabIndex        =   47
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox TextL5 
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
            TabIndex        =   46
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox TextV5 
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
            TabIndex        =   45
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label28 
            Caption         =   "R5:"
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
            TabIndex        =   52
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label27 
            Caption         =   "C5:"
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
            TabIndex        =   51
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label26 
            Caption         =   "L5:"
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
            TabIndex        =   50
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label25 
            Caption         =   "V5:"
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
            TabIndex        =   49
            Top             =   1320
            Width           =   735
         End
      End
      Begin VB.Frame Frame9 
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
         Left            =   6840
         TabIndex        =   35
         Top             =   1320
         Width           =   2895
         Begin VB.TextBox TextV2 
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
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox TextL2 
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
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox TextC2 
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
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox TextR2 
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
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label24 
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
            TabIndex        =   43
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label23 
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
            TabIndex        =   42
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label22 
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
            TabIndex        =   41
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label21 
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
            TabIndex        =   40
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Arm 4"
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
         TabIndex        =   23
         Top             =   3120
         Width           =   2055
         Begin VB.TextBox TextR4 
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
            TabIndex        =   28
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox TextC4 
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
            TabIndex        =   27
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox TextL4 
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
            TabIndex        =   26
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox TextV4 
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
            TabIndex        =   25
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox Textf 
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
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "R4:"
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
            TabIndex        =   33
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "C4:"
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
            TabIndex        =   32
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "L4:"
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
            TabIndex        =   31
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "V4:"
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
            TabIndex        =   30
            Top             =   1320
            Width           =   735
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
            TabIndex        =   29
            Top             =   1680
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
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
         Height          =   1695
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   2055
         Begin VB.TextBox TextR1 
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
            Left            =   720
            TabIndex        =   18
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox TextC1 
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
            Left            =   720
            TabIndex        =   17
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox TextL1 
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
            Left            =   720
            TabIndex        =   16
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox TextV1 
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
            Left            =   720
            TabIndex        =   15
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label7 
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
            Left            =   240
            TabIndex        =   22
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label8 
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
            Left            =   240
            TabIndex        =   21
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label9 
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
            Left            =   240
            TabIndex        =   20
            Top             =   960
            Width           =   735
         End
         Begin VB.Label Label10 
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
            Left            =   240
            TabIndex        =   19
            Top             =   1320
            Width           =   735
         End
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
         TabIndex        =   11
         Top             =   360
         Width           =   9495
         Begin VB.OptionButton Option1 
            Caption         =   "Mesh Analysis"
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
            Top             =   360
            Width           =   2895
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Nodal Analysis"
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
            Left            =   4320
            TabIndex        =   12
            Top             =   360
            Width           =   2895
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
         Height          =   495
         Left            =   2400
         TabIndex        =   10
         Top             =   4320
         Width           =   1575
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
         Left            =   2520
         TabIndex        =   1
         Top             =   1320
         Width           =   2055
         Begin VB.TextBox TextV3 
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
            TabIndex        =   5
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox TextL3 
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
            TabIndex        =   4
            Top             =   960
            Width           =   1215
         End
         Begin VB.TextBox TextC3 
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
            TabIndex        =   3
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox TextR3 
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
            TabIndex        =   2
            Top             =   240
            Width           =   1215
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
            TabIndex        =   9
            Top             =   1320
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
            TabIndex        =   8
            Top             =   960
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
            TabIndex        =   7
            Top             =   600
            Width           =   735
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
            TabIndex        =   6
            Top             =   240
            Width           =   735
         End
      End
   End
   Begin VB.Line Line2 
      BorderWidth     =   4
      X1              =   0
      X2              =   10320
      Y1              =   10200
      Y2              =   10200
   End
   Begin VB.Line Line1 
      BorderWidth     =   4
      X1              =   0
      X2              =   10320
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "THREE LOOP R, L, C CIRCUIT"
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
      Left            =   1680
      TabIndex        =   34
      Top             =   120
      Width           =   6375
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FF8080&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   -120
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "RLCThreeLoop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim R(5) As Single
Dim V(5) As Single
Dim L(5) As Single
Dim C(5) As Single
Dim Z(5) As Single
Dim f As Single
Dim n As Integer
Dim XC(5), XL(5) As Single
Dim Z1, Z2, Z3, Z4, Z5 As Single
Dim D, D1, D2, D3, D4, D5, D6 As Single
Dim V12, V23, V34 As Single
Dim Z12, Z21, Z23, Z31, Z34 As Single
Dim VA, VB, I1, I2, I3 As Single
Dim VC, VD, VE, VF, VG, VH As Single
Const vbkeyDecPt = 46
Const pi = 22 / 7

Private Sub calculate_Click()
R(0) = Val(TextR1.Text)
R(1) = Val(TextR2.Text)
R(2) = Val(TextR3.Text)
R(3) = Val(TextR4.Text)
R(4) = Val(TextR5.Text)
L(0) = Val(TextL1.Text)
L(1) = Val(TextL2.Text)
L(2) = Val(TextL3.Text)
L(3) = Val(TextL4.Text)
L(4) = Val(TextL5.Text)
C(0) = Val(TextC1.Text)
C(1) = Val(TextC2.Text)
C(2) = Val(TextC3.Text)
C(3) = Val(TextC4.Text)
C(4) = Val(TextC5.Text)
f = Val(Textf.Text)
V(0) = Val(TextV1.Text)
V(1) = Val(TextV2.Text)
V(2) = Val(TextV3.Text)
V(3) = Val(TextV4.Text)
V(4) = Val(TextV5.Text)
For n = 0 To 4
If L(n) = 0 Then
XL(n) = 0
Else
XL(n) = 2 * pi * f * L(n)
End If
Next n
For n = 0 To 4
If C(n) = 0 Or f = 0 Then
XC(n) = 0
Else
XC(n) = 1 / (2 * pi * f * C(n))
End If
Next n
Z1 = (R(0) ^ 2 + (XL(0) - XC(0)) ^ 2) ^ (1 / 2)
Z2 = (R(1) ^ 2 + (XL(1) - XC(1)) ^ 2) ^ (1 / 2)
Z3 = (R(2) ^ 2 + (XL(2) - XC(2)) ^ 2) ^ (1 / 2)
Z4 = (R(3) ^ 2 + (XL(3) - XC(3)) ^ 2) ^ (1 / 2)
Z5 = (R(4) ^ 2 + (XL(4) - XC(4)) ^ 2) ^ (1 / 2)

If Option1.Value = True Then
V12 = V(0) - V(1)
Z12 = Z1 + Z2
Z21 = -Z2
V23 = V(1) - V(2) - V(4)
Z23 = -Z3
V34 = V(2) - V(3)
Z34 = Z3 + Z4
D = (Z12 * Z23 * Z34) - (Z12 * (Z31) ^ 2) - ((Z21) ^ 2 * Z34)
D1 = (V12 * Z23 * Z34) - (V12 * (Z31) ^ 2) - (Z21 * V23 * Z34) + (Z21 * Z31 * V34)
D2 = (V23 * Z12 * Z34) - (V34 * Z12 * Z31) - (V12 * Z21 * Z34)
D3 = (V34 * Z12 * Z23) - (V23 * Z12 * Z31) - (V34 * (Z21) ^ 2) + (V12 * Z21 * Z31)
I1 = D1 / D
I2 = D2 / D
I3 = D3 / D
LblI1.Caption = Format$(I1, "###.000000A")
LblI2.Caption = Format$(I2, "###.000000A")
LblI3.Caption = Format$(I3, "###.000000A")
LblVA.Caption = " "
LblVB.Caption = " "
ElseIf Option2.Value = True Then
VC = (Z2 * Z5) + (Z1 * Z5) + (Z1 * Z2)
VD = -(Z1 * Z2)
VE = (V(0) * Z2 * Z5) + (V(1) * Z1 * Z5) + (V(4) * Z1 * Z2)
VF = -(Z3 * Z4)
VG = (V(2) * Z4 * Z5) + (V(3) * Z3 * Z5) - (V(4) * Z3 * Z4)
VH = (Z4 * Z5) + (Z3 * Z5) + (Z3 * Z4)
D4 = (VC * VH) - (VD * VF)
D5 = (VE * VH) - (VD * VG)
D6 = (VC * VG) - (VE * VF)
VA = D5 / D4
VB = D6 / D4
LblVA.Caption = Format$(VA, "###.000000V")
LblVB.Caption = Format$(VB, "###.000000V")
LblI1.Caption = " "
LblI2.Caption = " "
LblI3.Caption = " "
End If
End Sub

Private Sub clear_Click()
TextR1.Text = " "
TextR2.Text = " "
TextR3.Text = " "
TextR4.Text = " "
TextR5.Text = " "
TextL1.Text = " "
TextL2.Text = " "
TextL3.Text = " "
TextL4.Text = " "
TextL5.Text = " "
TextC1.Text = " "
TextC2.Text = " "
TextC3.Text = " "
TextC4.Text = " "
TextC5.Text = " "
Textf.Text = " "
TextV1.Text = " "
TextV2.Text = " "
TextV3.Text = " "
TextV4.Text = " "
TextV5.Text = " "
LblVA.Caption = " "
LblVB.Caption = " "
LblI1.Caption = " "
LblI2.Caption = " "
LblI3.Caption = " "
End Sub

Private Sub Command1_Click()
Unload Me
Splash.Show
End Sub


Private Sub LblVA_Click()

End Sub

Private Sub TextC1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub


Private Sub TextC2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub


Private Sub TextC3_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub


Private Sub TextC4_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub


Private Sub TextC5_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub


Private Sub Textf_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub


Private Sub TextL1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub


Private Sub TextL2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub


Private Sub TextL3_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub


Private Sub TextL4_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub


Private Sub TextL5_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub


Private Sub TextR1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub TextR2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub TextR3_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub TextR4_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub


Private Sub TextR5_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub


Private Sub TextV1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub TextV2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub TextV3_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub TextV4_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub

Private Sub TextV5_KeyPress(KeyAscii As Integer)
If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or _
KeyAscii = vbKeyBack Or KeyAscii = vbkeyDecPt Then
Exit Sub
Else
KeyAscii = 0
Beep
End If
End Sub
