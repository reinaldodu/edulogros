VERSION 5.00
Begin VB.Form IMPARTEC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3015
   Icon            =   "COPIARTEC.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   3015
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      Height          =   320
      Left            =   840
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Campos que se imprimen"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2775
      Begin VB.ListBox List1 
         Height          =   1410
         ItemData        =   "COPIARTEC.frx":0442
         Left            =   120
         List            =   "COPIARTEC.frx":0458
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
   End
End
Attribute VB_Name = "IMPARTEC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If List1.SelCount = 0 Then
    MsgBox "No ha seleccionado ningún campo para imprimir", 64, "Imprimir"
    Exit Sub
End If
RESP = MsgBox("Desea imprimir esta información?", vbYesNo + vbQuestion + vbDefaultButton2, "Imprimir")
    If RESP = vbYes Then
        Screen.MousePointer = 11
        Printer.Orientation = 2
        Printer.PaperSize = 5
        PAG = 1
        NAR = FreeFile
        Open Ruta & "inicial.edu" For Input As #NAR
        Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector
        Close #NAR
        Printer.ScaleMode = 7
        Printer.CurrentY = 1
        Printer.CurrentX = 6
        Printer.Print "CONSULTA DE BOLETIN POR GRADO - " & GRUPOSTEC.Label5.Caption
        Printer.Print ""
        Printer.Print ""
        Printer.CurrentX = 1
        Printer.Print ini.nombre;
        Printer.CurrentX = 31
        Printer.Print "Pág." & PAG
        Printer.Print ""
        Printer.Print ""
        For I = 0 To (GRUPOSTEC.MATXECN.Rows - 1)
            Printer.CurrentX = 1
            Printer.Print GRUPOSTEC.MATXECN.TextMatrix(I, 0);
            CX = 1.5
            If List1.Selected(0) = True Then
                Printer.CurrentX = CX
                Printer.Print GRUPOSTEC.MATXECN.TextMatrix(I, 1);
                CX = CX + 6.5
            End If
            If List1.Selected(1) = True Then
                Printer.CurrentX = CX
                Printer.Print GRUPOSTEC.MATXECN.TextMatrix(I, 14);
                CX = CX + 2
            End If
            If List1.Selected(2) = True Then
                Printer.CurrentX = CX
                Printer.Print GRUPOSTEC.MATXECN.TextMatrix(I, 15);
                CX = CX + 3
            End If
            If List1.Selected(3) = True Then
                Printer.CurrentX = CX
                Printer.Print GRUPOSTEC.MATXECN.TextMatrix(I, 12);
                CX = CX + 1
            End If
            If List1.Selected(4) = True Then
                Printer.CurrentX = CX
                Printer.Print GRUPOSTEC.MATXECN.TextMatrix(I, 13);
                CX = CX + 1
            End If
            If List1.Selected(5) = True Then
                For J = 2 To 11
                    Printer.CurrentX = CX
                    Printer.Print GRUPOSTEC.MATXECN.TextMatrix(I, J);
                    CX = CX + 1
                Next J
            End If
            If I = 0 Then Printer.Print ""
            Printer.Print ""
            If (I <> 0) And ((I Mod 47) = 0) Then
                Printer.NewPage
                PAG = PAG + 1
                Printer.CurrentY = 1
                Printer.CurrentX = 6
                Printer.Print "CONSULTA DE BOLETIN POR GRADO - " & GRUPOSTEC.Label5.Caption
                Printer.Print ""
                Printer.Print ""
                Printer.CurrentX = 1
                Printer.Print ini.nombre
                Printer.CurrentX = 31
                Printer.Print "Pág." & PAG
                Printer.Print ""
                Printer.Print ""
                Printer.Print ""
            End If
        Next I
        Printer.EndDoc
        Unload Me
        Printer.Orientation = 1
        Printer.PaperSize = 1
        Screen.MousePointer = 0
    End If
End Sub
