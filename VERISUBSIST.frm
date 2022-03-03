VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form VERISUBSIST 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Verificación de actualización y baja disco boletines de subsistemas"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "VERISUBSIST.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      Begin VB.Frame Frame3 
         Caption         =   "Fechas de baja discos boletines Subsistemas"
         Height          =   2535
         Left            =   3840
         TabIndex        =   3
         Top             =   120
         Width           =   3495
         Begin MSFlexGridLib.MSFlexGrid MTBAJ 
            Height          =   2175
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   3836
            _Version        =   393216
            Rows            =   1
            FixedCols       =   0
            BackColorBkg    =   12632256
            GridColor       =   12582912
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Fechas de Actualizaciones Subsistemas"
         Height          =   2535
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   3495
         Begin MSFlexGridLib.MSFlexGrid MTACT 
            Height          =   2175
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   3836
            _Version        =   393216
            Rows            =   1
            FixedCols       =   0
            BackColorBkg    =   12632256
            GridColor       =   12582912
         End
      End
   End
End
Attribute VB_Name = "VERISUBSIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Verificar las fechas de actualización y baja disco boletines de los subsistemas."
End Sub

Private Sub Form_Load()
MTACT.ColWidth(0) = 1650
MTACT.ColWidth(1) = 1200
MTBAJ.ColWidth(0) = 1650
MTBAJ.ColWidth(1) = 1200
MTACT.Row = 0
MTACT.Col = 0
MTACT.CellForeColor = RGB(255, 255, 255)
MTACT.CellBackColor = RGB(0, 0, 150)
MTACT.Text = "SUBSISTEMA"
MTACT.Col = 1
MTACT.CellForeColor = RGB(255, 255, 255)
MTACT.CellBackColor = RGB(0, 0, 150)
MTACT.Text = "FECHA"
MTBAJ.Row = 0
MTBAJ.Col = 0
MTBAJ.CellForeColor = RGB(255, 255, 255)
MTBAJ.CellBackColor = RGB(0, 0, 150)
MTBAJ.Text = "SUBSISTEMA"
MTBAJ.Col = 1
MTBAJ.CellForeColor = RGB(255, 255, 255)
MTBAJ.CellBackColor = RGB(0, 0, 150)
MTBAJ.Text = "FECHA"
NAR = FreeFile
Open Ruta & "infosub.edu" For Input As #NAR
While Not EOF(NAR)
    Input #NAR, infsub.subsistema, infsub.actualsub, infsub.bajasub
    If RTrim(infsub.bajasub) = "" Then
        MTACT.Rows = MTACT.Rows + 1
        MTACT.TextMatrix((MTACT.Rows - 1), 0) = "SUBSISTEMA No." & infsub.subsistema
        MTACT.TextMatrix((MTACT.Rows - 1), 1) = infsub.actualsub
    Else
        MTBAJ.Rows = MTBAJ.Rows + 1
        MTBAJ.TextMatrix((MTBAJ.Rows - 1), 0) = "SUBSISTEMA No." & infsub.subsistema
        MTBAJ.TextMatrix((MTBAJ.Rows - 1), 1) = infsub.bajasub
    End If
Wend
Close #NAR
End Sub
