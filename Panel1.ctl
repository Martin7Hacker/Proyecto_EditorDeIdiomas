VERSION 5.00
Begin VB.UserControl Panel1 
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3645
   Picture         =   "Panel1.ctx":0000
   ScaleHeight     =   735
   ScaleWidth      =   3645
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Panel1.ctx":B5DF
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Idioma Actual"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   2715
   End
End
Attribute VB_Name = "Panel1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
