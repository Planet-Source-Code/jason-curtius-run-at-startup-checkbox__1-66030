VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TestForm"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin RunAtStartUp.Run_At_StartUp Run_At_StartUp1 
      Height          =   195
      Left            =   2280
      TabIndex        =   1
      Top             =   1440
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   344
      prbCaption      =   "Run at Startup"
      prbBackColor    =   -2147483633
      prbForeColor    =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      prbAlignment    =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Query"
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   2640
      Width           =   1455
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox Run_At_StartUp1.qurey
End Sub
