VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Alias_Grupo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alias - Grupo"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin MSFlexGridLib.MSFlexGrid MtAlias 
         Height          =   4095
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   7223
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "Alias_Grupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Dir(Ruta & "aliasgrupos.edu") <> "" Then
   Kill Ruta & "aliasgrupos.edu"
End If
NAR = FreeFile
For I = 1 To MtAlias.Rows - 1
    aliasg = MtAlias.TextMatrix(I, 1)
    Open Ruta & "aliasgrupos.edu" For Append As #NAR
    Write #NAR, aliasg
    Close #NAR
Next I
Unload Me
End Sub

Private Sub Form_Load()
If Dir(Ruta & "infcur.edu") <> "" Then
    Command1.Enabled = True
Else
    Command1.Enabled = False
    Exit Sub
End If
MtAlias.Row = 0
MtAlias.Col = 0
MtAlias.ColWidth(0) = 1700
MtAlias.CellFontBold = True
MtAlias.CellForeColor = RGB(255, 255, 255)
MtAlias.CellBackColor = RGB(0, 0, 150)
MtAlias.Text = "      G R U P O  "
MtAlias.Col = 1
MtAlias.ColWidth(1) = 2500
MtAlias.CellFontBold = True
MtAlias.CellForeColor = RGB(255, 255, 255)
MtAlias.CellBackColor = RGB(0, 0, 150)
MtAlias.Text = "            A L I A S    "

CONT = 0
NAR = FreeFile
Open Ruta & "infcur.edu" For Input As #NAR
While Not EOF(NAR)
    Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
    CONT = CONT + 1
    MtAlias.Rows = CONT + 1
    MtAlias.TextMatrix(CONT, 0) = RTrim(icur.nom)
Wend
Close #NAR
CONT = 0
If Dir(Ruta & "aliasgrupos.edu") <> "" Then
    Open Ruta & "aliasgrupos.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, aliasg
        CONT = CONT + 1
        MtAlias.TextMatrix(CONT, 1) = RTrim(aliasg)
    Wend
    Close #NAR
End If
End Sub

Private Sub MtAlias_KeyPress(KeyAscii As Integer)
If MtAlias.Col > 0 Then
    If KeyAscii = 8 Then
        If Trim(MtAlias.Text) <> "" Then
           MtAlias.Text = Left(MtAlias.Text, Len(MtAlias.Text) - 1)
        End If
    Else
        valias = Chr(KeyAscii)
        MtAlias.CellForeColor = RGB(0, 0, 0)
        MtAlias.CellFontBold = False
        MtAlias.Text = MtAlias.Text + valias
    End If
End If
End Sub
