VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Folio_Config 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configurar número de folio"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5055
   Icon            =   "Folio_Config.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "A&ceptar"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin MSFlexGridLib.MSFlexGrid Mxfolio 
         Height          =   1050
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4490
         _ExtentX        =   7911
         _ExtentY        =   1852
         _Version        =   393216
         Rows            =   4
         Cols            =   5
      End
   End
End
Attribute VB_Name = "Folio_Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
NAR = FreeFile
Open Ruta & "folio.edu" For Output As #NAR
Write #NAR, Mxfolio.TextMatrix(1, 1), Mxfolio.TextMatrix(1, 2), Mxfolio.TextMatrix(1, 3), Mxfolio.TextMatrix(1, 4), _
Mxfolio.TextMatrix(2, 1), Mxfolio.TextMatrix(2, 2), Mxfolio.TextMatrix(2, 3), Mxfolio.TextMatrix(2, 4), _
Mxfolio.TextMatrix(3, 1), Mxfolio.TextMatrix(3, 2), Mxfolio.TextMatrix(3, 3), Mxfolio.TextMatrix(3, 4)
Close #NAR
Unload Me
End Sub

Private Sub Form_Load()
Dim InpFolio As String, contfolio As Byte
Dim masrow As Byte, mascol As Byte
Mxfolio.ColWidth(0) = 1200
Mxfolio.TextMatrix(1, 0) = "PREESCOLAR"
Mxfolio.TextMatrix(2, 0) = "PRIMARIA"
Mxfolio.TextMatrix(3, 0) = "SECUNDARIA"
Mxfolio.ColWidth(1) = 800
Mxfolio.TextMatrix(0, 1) = "UNICA"
Mxfolio.ColWidth(2) = 800
Mxfolio.TextMatrix(0, 2) = "MAÑANA"
Mxfolio.ColWidth(3) = 800
Mxfolio.TextMatrix(0, 3) = "TARDE"
Mxfolio.ColWidth(4) = 800
Mxfolio.TextMatrix(0, 4) = "NOCHE"
If Dir(Ruta & "folio.edu") <> "" Then
    contfolio = 1
    masrow = 1
    mascol = 1
    NAR = FreeFile
    Open Ruta & "folio.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, InpFolio
        Mxfolio.TextMatrix(masrow, mascol) = InpFolio
        If contfolio Mod 4 = 0 And contfolio <> 12 Then
            masrow = masrow + 1
            mascol = 1
        Else
            mascol = mascol + 1
        End If
        contfolio = contfolio + 1
    Wend
    Close #NAR
End If
End Sub

Private Sub Mxfolio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Mxfolio.Col = 4 And Mxfolio.Row <> 3 Then
        Mxfolio.Row = Mxfolio.Row + 1
        Mxfolio.Col = 1
        Exit Sub
    End If
    If Mxfolio.Row = 3 And Mxfolio.Col = 4 Then
    Exit Sub
    Else
        Mxfolio.Col = Mxfolio.Col + 1
        Exit Sub
    End If
End If
If KeyAscii = 8 Then
   If Mxfolio.Text <> "" Then
      Mxfolio.Text = Left(Mxfolio.Text, Len(Mxfolio.Text) - 1)
   End If
   Exit Sub
End If
Mxfolio.Text = Mxfolio.Text + Chr(KeyAscii)
End Sub
