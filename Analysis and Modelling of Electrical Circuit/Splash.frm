VERSION 5.00
Object = "{E63844D0-CCAA-43B9-BBC0-78BB7CCD6AC2}#1.0#0"; "AUFOBUTTON.OCX"
Begin VB.Form Splash 
   BackColor       =   &H00E33D2B&
   BorderStyle     =   0  'None
   Caption         =   "Main Menu"
   ClientHeight    =   8925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7380
   Icon            =   "Splash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin aUfoButton.Button Simple 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   4560
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImageCellHeight =   0
      ImageCellWidth  =   0
      BorderColorTL   =   255
      BorderColorTLover=   16744576
      BorderColorBR   =   255
      BorderColorBRover=   16744576
      BackColor       =   12640511
      BackColorOver   =   16777215
      ForeColor       =   0
      ForeColorOver   =   255
      Caption         =   "Simple Circuit"
      ShowFocus       =   0   'False
      ShowFocusColor  =   0
   End
   Begin aUfoButton.Button Series 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   5040
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImageCellHeight =   0
      ImageCellWidth  =   0
      BorderColorTL   =   255
      BorderColorTLover=   16744576
      BorderColorBR   =   255
      BorderColorBRover=   16744576
      BackColor       =   12640511
      BackColorOver   =   16777215
      ForeColor       =   0
      ForeColorOver   =   255
      Caption         =   "Series Circuit"
      ShowFocus       =   0   'False
      ShowFocusColor  =   0
   End
   Begin aUfoButton.Button Exit 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   7920
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImageCellHeight =   0
      ImageCellWidth  =   0
      BorderColorTL   =   255
      BorderColorTLover=   16744576
      BorderColorBR   =   255
      BorderColorBRover=   16744576
      BackColor       =   12640511
      BackColorOver   =   16777215
      ForeColor       =   0
      ForeColorOver   =   255
      Caption         =   "Exit Application"
      ShowFocus       =   0   'False
      ShowFocusColor  =   0
   End
   Begin aUfoButton.Button AboutFRM 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   7440
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImageCellHeight =   0
      ImageCellWidth  =   0
      BorderColorTL   =   255
      BorderColorTLover=   16744576
      BorderColorBR   =   255
      BorderColorBRover=   16744576
      BackColor       =   12640511
      BackColorOver   =   16777215
      ForeColor       =   0
      ForeColorOver   =   255
      Caption         =   "About"
      ShowFocus       =   0   'False
      ShowFocusColor  =   0
   End
   Begin aUfoButton.Button Parallel 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   5520
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImageCellHeight =   0
      ImageCellWidth  =   0
      BorderColorTL   =   255
      BorderColorTLover=   16744576
      BorderColorBR   =   255
      BorderColorBRover=   16744576
      BackColor       =   12640511
      BackColorOver   =   16777215
      ForeColor       =   0
      ForeColorOver   =   255
      Caption         =   "Parallel Circuit"
      ShowFocus       =   0   'False
      ShowFocusColor  =   0
   End
   Begin aUfoButton.Button RLC 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   6000
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImageCellHeight =   0
      ImageCellWidth  =   0
      BorderColorTL   =   255
      BorderColorTLover=   16744576
      BorderColorBR   =   255
      BorderColorBRover=   16744576
      BackColor       =   12640511
      BackColorOver   =   16777215
      ForeColor       =   0
      ForeColorOver   =   255
      Caption         =   "AC Circuit -  R, L, C in Series"
      ShowFocus       =   0   'False
      ShowFocusColor  =   0
   End
   Begin aUfoButton.Button TWOLOOP 
      Height          =   375
      Left            =   1440
      TabIndex        =   6
      Top             =   6480
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImageCellHeight =   0
      ImageCellWidth  =   0
      BorderColorTL   =   255
      BorderColorTLover=   16744576
      BorderColorBR   =   255
      BorderColorBRover=   16744576
      BackColor       =   12640511
      BackColorOver   =   16777215
      ForeColor       =   0
      ForeColorOver   =   255
      Caption         =   "AC Circuit -  two loop Circuit"
      ShowFocus       =   0   'False
      ShowFocusColor  =   0
   End
   Begin aUfoButton.Button THREELOOP 
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   6960
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ImageCellHeight =   0
      ImageCellWidth  =   0
      BorderColorTL   =   255
      BorderColorTLover=   16744576
      BorderColorBR   =   255
      BorderColorBRover=   16744576
      BackColor       =   12640511
      BackColorOver   =   16777215
      ForeColor       =   0
      ForeColorOver   =   255
      Caption         =   "AC Circuit -  three loop Circuit"
      ShowFocus       =   0   'False
      ShowFocusColor  =   0
   End
   Begin VB.Image Image1 
      Height          =   8925
      Left            =   0
      Picture         =   "Splash.frx":038A
      Top             =   0
      Width           =   7380
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub AboutFRM_Click()
About.Show
End Sub
Private Sub exit_Click()
End
End Sub
Private Sub Parallel_Click()
ParallelCircuit.Show
Splash.Hide
End Sub
Private Sub RLCSeries_Click()
RLCSeries.Show
Splash.Hide
End Sub
Private Sub RLC_Click()
RLCSeries.Show
Splash.Hide
End Sub
Private Sub Series_Click()
SeriesCircuit.Show
Splash.Hide
End Sub
Private Sub Simple_Click()
SimpleCircuit.Show
Splash.Hide
End Sub
Private Sub THREELOOP_Click()
RLCThreeLoop.Show
Splash.Hide
End Sub
Private Sub TWOLOOP_Click()
RLCTwoLoop.Show
Splash.Hide
End Sub
