VERSION 5.00
Begin VB.Form AREASTEC 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   Icon            =   "AREASTEC.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      ToolTipText     =   "Imprimir listado de grupos"
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin VB.Frame Frame3 
         Caption         =   "Grupos del Subsistema"
         Height          =   2535
         Left            =   3720
         TabIndex        =   5
         Top             =   120
         Width           =   3015
         Begin VB.ListBox List2 
            Height          =   2205
            Left            =   120
            MultiSelect     =   1  'Simple
            TabIndex        =   7
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Grupos existentes"
         Height          =   2535
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   3015
         Begin VB.ListBox List1 
            Height          =   2205
            Left            =   120
            MultiSelect     =   1  'Simple
            TabIndex        =   6
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "<<"
         Height          =   255
         Left            =   3240
         TabIndex        =   3
         ToolTipText     =   "Eliminar grupo(s)"
         Top             =   1560
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Caption         =   ">>"
         Height          =   255
         Left            =   3240
         TabIndex        =   2
         ToolTipText     =   "Agregar grupo(s)"
         Top             =   1200
         Width           =   375
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Visible         =   0   'False
      Width           =   45
   End
End
Attribute VB_Name = "AREASTEC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
I = 0
While I < List1.ListCount
   If List1.Selected(I) = True Then
        NAR = FreeFile
        Open Label1.Caption For Append As #NAR
        Write #NAR, RTrim(List1.List(I))
        Close #NAR
        List2.AddItem RTrim(List1.List(I))
        List1.RemoveItem I
   Else
        I = I + 1
   End If
Wend
End Sub

Private Sub Command2_Click()
I = 0
While I < List2.ListCount
   If List2.Selected(I) = True Then
        List1.AddItem RTrim(List2.List(I))
        List2.RemoveItem I
   Else
        I = I + 1
   End If
Wend
If Dir(Label1.Caption) <> "" Then
    Kill Label1.Caption
End If
For J = 0 To (List2.ListCount - 1)
    NAR = FreeFile
    Open Label1.Caption For Append As #NAR
    Write #NAR, RTrim(List2.List(J))
    Close #NAR
Next J
End Sub

Private Sub Command3_Click()
RESP = MsgBox("Desea imprimir la lista de grupos?", vbYesNo + vbQuestion + vbDefaultButton2, "Imprimir grupos")
If RESP = vbYes Then
    Screen.MousePointer = 11
    NAR = FreeFile
    Open Ruta & "inicial.edu" For Input As #NAR
    Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector
    Close #NAR
    Printer.ScaleMode = 7
    Printer.CurrentY = 1
    Printer.CurrentX = 9
    Printer.Print "LISTADO DE GRUPOS"
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 2.5
    Printer.Print RTrim(ini.nombre)
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 2.5
    Printer.FontUnderline = True
    Printer.Print "GRUPOS EXISTENTES";
    Printer.CurrentX = 12
    Printer.Print Format(AREASTEC.Caption, ">")
    Printer.FontUnderline = False
    Printer.Print ""
    CY = Printer.CurrentY
    For I = 0 To (List1.ListCount - 1)
        Printer.CurrentX = 2.5
        Printer.Print List1.List(I)
    Next I
    Printer.Print ""
    Printer.CurrentX = 2.5
    Printer.Print "TOTAL..." & List1.ListCount
    Printer.CurrentY = CY
    For I = 0 To (List2.ListCount - 1)
        Printer.CurrentX = 12
        Printer.Print List2.List(I)
    Next I
    Printer.Print ""
    Printer.CurrentX = 12
    Printer.Print "TOTAL..." & List2.ListCount
    Printer.EndDoc
    Screen.MousePointer = 0
End If
End Sub

Private Sub Form_Activate()
ENTRADA.stb.Panels.Item(1).Text = "Creación de grupos para el Subsistema seleccionado."
End Sub

Private Sub Form_Load()
If Dir(Ruta & "infcur.edu") <> "" Then
    NAR = FreeFile
    Open Ruta & "infcur.edu" For Input As #NAR
    While Not EOF(NAR)
        Input #NAR, icur.nom, icur.jornada, icur.grado, icur.director
        List1.AddItem RTrim(icur.nom)
    Wend
    Close #NAR
    Label1.Caption = Ruta & "subsis" & (SELECSUBS.List1.ListIndex) + 1 & ".sub"
    If Dir(Label1.Caption) <> "" Then
        Open Label1.Caption For Input As #NAR
        While Not EOF(NAR)
            Input #NAR, TTT
            If RTrim(TTT) <> "" Then
                List2.AddItem RTrim(TTT)
            End If
            For I = 0 To (List1.ListCount - 1)
                If List1.List(I) = RTrim(TTT) Then
                    List1.RemoveItem I
                    Exit For
                End If
            Next I
        Wend
        Close #NAR
    End If
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
Else
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
End If
End Sub
