VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CONSARESUB 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Grupos-Subsistema"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3495
   Icon            =   "CONSARESUB.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   3495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   3000
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin MSFlexGridLib.MSFlexGrid MTXSUBS 
         Height          =   2535
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   4471
         _Version        =   393216
         Rows            =   0
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   12632256
         GridColor       =   12582912
      End
   End
End
Attribute VB_Name = "CONSARESUB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
RESP = MsgBox("Desea imprimir la lista de grupos?", vbYesNo + vbQuestion + vbDefaultButton2, "Imprimir grupos")
If RESP = vbYes Then
    Screen.MousePointer = 11
    NAR = FreeFile
    Open Ruta & "inicial.edu" For Input As #NAR
    Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector
    Close #NAR
    Printer.ScaleMode = 7
    Printer.CurrentY = 1
    Printer.CurrentX = 8.5
    Printer.Print Format(Frame1.Caption, ">")
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 2.5
    Printer.Print RTrim(ini.nombre)
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 2.5
    Printer.Print ""
    For I = 0 To (MTXSUBS.Rows - 1)
        Printer.CurrentX = 2.5
        Printer.Print MTXSUBS.TextMatrix(I, 0)
    Next I
    Printer.Print ""
    Printer.CurrentX = 2.5
    Printer.Print "TOTAL..." & MTXSUBS.Rows
    Printer.EndDoc
    Screen.MousePointer = 0
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Consulta de grupos por Subsistemas."
End Sub

Private Sub Form_Load()
MTXSUBS.ColWidth(0) = 2700
If Dir(Ruta & "subsis" & (SELECSUBS.List1.ListIndex) + 1 & ".sub") <> "" Then
    NAR = FreeFile
    Open Ruta & "subsis" & (SELECSUBS.List1.ListIndex) + 1 & ".sub" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, TTT
        If RTrim(TTT) <> "" Then
            MTXSUBS.AddItem RTrim(TTT)
        End If
    Wend
    Close #NAR
End If
End Sub
