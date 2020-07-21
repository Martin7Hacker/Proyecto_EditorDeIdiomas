VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmidioma 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Editor de Idiomas: Virtual Martin Temporize v2017"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11040
   Icon            =   "frmidioma.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   11040
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cobT 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   7440
      Width           =   3615
   End
   Begin VB.Timer timLista 
      Interval        =   1
      Left            =   8160
      Top             =   7200
   End
   Begin MSComDlg.CommonDialog cdgAbrir 
      Left            =   8640
      Top             =   7200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdDuplicar 
      BackColor       =   &H00FFFFC0&
      Caption         =   "X----->| &Duplicar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7200
      Width           =   1815
   End
   Begin VB.ListBox lblnumero 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   6075
      ItemData        =   "frmidioma.frx":0CCA
      Left            =   120
      List            =   "frmidioma.frx":0CCC
      TabIndex        =   10
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton cmdAbrir 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Abrir .txt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdGuardar 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Grabar .txt"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Enabled         =   0   'False
      Height          =   255
      Left            =   9840
      TabIndex        =   6
      Top             =   777
      Width           =   1095
   End
   Begin VB.TextBox txtEditar 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   7320
      TabIndex        =   5
      Top             =   750
      Width           =   2415
   End
   Begin VB.ListBox lblhidiomaNuevo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   6075
      ItemData        =   "frmidioma.frx":0CCE
      Left            =   5880
      List            =   "frmidioma.frx":0CD0
      TabIndex        =   2
      Top             =   1080
      Width           =   5055
   End
   Begin VB.ListBox lblhidiomaBase 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   6075
      ItemData        =   "frmidioma.frx":0CD2
      Left            =   960
      List            =   "frmidioma.frx":0CD4
      TabIndex        =   1
      Top             =   1080
      Width           =   4935
   End
   Begin Editor_de_Idiomas.Panel1 Panel11 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   1296
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Abrir en Panel :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1440
      TabIndex        =   13
      Top             =   7200
      Width           =   3735
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N-Ficha"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   630
   End
   Begin VB.Image imgIdioma 
      Height          =   480
      Left            =   5160
      Picture         =   "frmidioma.frx":0CD6
      Top             =   7200
      Width           =   480
   End
   Begin VB.Label lblidioma 
      BackStyle       =   0  'Transparent
      Caption         =   "Idioma .txt"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   3600
      TabIndex        =   7
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ficha Idioma Editable*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Top             =   840
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ficha Idioma / Base*"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   840
      Width           =   5055
   End
End
Attribute VB_Name = "frmidioma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************
'                                      *
' Editor de Idiomas                    *
' Autor: Martin Grasso Castrillo.      *
' Para Virtual Martin Temporize v1.7   *
'                                      *
'***************************************
Private dato, archivo, guardar As String
Private Idice As Integer

Private Sub cmdAbrir_Click()
On Error GoTo no_se
Idice = 0
lblhidiomaBase.Clear
lblhidiomaNuevo.Clear
lblnumero.Clear
With cdgAbrir
 .DialogTitle = "Abrir Archivo..."
 .Filter = "Virtual Martin temporize(*.vmt)|*.vmt|todos los Archivos (*.*)|*.*|Archivos de Idioma (*.txt*)|*.txt*|"
 .FilterIndex = 3
 .ShowOpen
If Not (.FileName = "") Then
  If .FileName <> "" Then
   If .CancelError = False Then
   archivo = .FileName
   lblidioma.Caption = .FileTitle
   .FileName = ""
 End If
  End If
   End If
End With

Open archivo For Input As 1
 Do While Not EOF(1)
 Line Input #1, dato
  Select Case cobT.ListIndex
  Case (0)
  lblhidiomaBase.AddItem dato
  controlActivo False
  Case (1)
  lblhidiomaNuevo.AddItem dato
  controlActivo True
  Case (2)
  lblhidiomaBase.AddItem dato
  lblhidiomaNuevo.AddItem dato
  controlActivo True
 End Select
  Idice = Idice + 1
  lblnumero.AddItem Idice
  Loop
Close #1
no_se:
End Sub

Private Sub cmdDuplicar_Click()
lblhidiomaNuevo.List(lblhidiomaNuevo.ListIndex) = lblhidiomaBase.Text
End Sub

Private Sub cmdEditar_Click()
If Not (txtEditar.Text = "") Then
   lblhidiomaNuevo.List(lblhidiomaNuevo.ListIndex) = txtEditar.Text
   txtEditar.Text = ""
End If
End Sub

Private Sub cmdGuardar_Click()
On Error GoTo nose
 With cdgAbrir
 .DialogTitle = "Virtual Martin temporize v1.0: Guardar Archivo"
 .Filter = "Virtual Martin temporize(*.vmt)|*.vmt|todos los Archivos (*.*)|*.*|Archivos de Idioma (*.txt*)|*.txt*|"
 .FilterIndex = 3
 .FileName = "Nuevo"
 .ShowSave
 If .FileName = "" Then
 MsgBox "No le asignaste un nombre de archivo", vbInformation
 End If
  guardar = .FileName
 If .FileName <> "" Then
 If .CancelError = False Then
 Dim x As Integer
  Open guardar & ".txt" For Output As 1
 For x = 0 To lblhidiomaNuevo.ListCount - 1
 Print #1, lblhidiomaNuevo.List(x)
 Next x
 Close #1
Else
 End If
 End If
 End With
nose:
End Sub

Private Sub Form_Load()
crearRegistros
cobT.ListIndex = 2
End Sub

Private Sub Form_Resize()
With Panel11
     .Width = Me.Width
End With
End Sub
 
Private Sub lblhidiomaBase_Click()
lblnumero.ListIndex = lblhidiomaBase.ListIndex
End Sub

Private Sub lblnumero_Click()
On Error GoTo nose
  Select Case cobT.ListIndex
Case (0)
  lblhidiomaBase.ListIndex = lblnumero.ListIndex
Case (1)
  lblhidiomaNuevo.ListIndex = lblnumero.ListIndex
Case (2)
  lblhidiomaNuevo.ListIndex = lblnumero.ListIndex
  lblhidiomaBase.ListIndex = lblnumero.ListIndex
  End Select
nose:
End Sub

Private Sub crearRegistros()
With cobT
    .Clear
    .AddItem "Fichas en Panel Estático."
    .AddItem "Fichas en Panel Editable."
    .AddItem "Fichas en Ambos Casos."
End With
End Sub

Private Sub lblnumero_Scroll()
lblnumero_Click
End Sub

Private Sub timLista_Timer()
lblnumero_Click
End Sub

Private Sub controlActivo(ByVal control As Boolean)
txtEditar.Enabled = control
cmdEditar.Enabled = control
cmdGuardar.Enabled = control
cmdDuplicar.Enabled = control
End Sub
