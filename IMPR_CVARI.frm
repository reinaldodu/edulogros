VERSION 5.00
Begin VB.Form IMPR_CVARI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir consulta"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2895
   Icon            =   "IMPR_CVARI.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   2895
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      ToolTipText     =   "Imprime los campos seleccionados"
      Top             =   3240
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   2760
      ItemData        =   "IMPR_CVARI.frx":0442
      Left            =   240
      List            =   "IMPR_CVARI.frx":046A
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Campos que se imprimen"
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   45
   End
End
Attribute VB_Name = "IMPR_CVARI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim ini As inicio
RESP = MsgBox("DESEA IMPRIMIR ESTA INFORMACION?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR")
If RESP = vbYes Then
        Screen.MousePointer = 11
        PAG = 1
        NAR = FreeFile
        Open Ruta & "inicial.edu" For Input As #NAR
        Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
        Close #NAR
        Printer.ScaleMode = 7
        Printer.CurrentY = 0.5
        Printer.CurrentX = 5.5
        Printer.Print Format(CONSVARIAS.Caption, ">")
        Printer.CurrentY = 1.5
        Printer.CurrentX = 0.8
        Printer.Print CONSVARIAS.Frame1.Caption;
        Printer.CurrentX = 14
        Printer.Print "Fecha: " & Format(Date, "mmm/dd/yyyy");
        Printer.CurrentX = 19
        Printer.Print "Pág." & PAG
        Printer.CurrentX = 0.8
        Printer.Print ini.nombre
        Printer.CurrentY = 2.5
        Printer.CurrentX = 0.8
        Printer.Print CONSVARIAS.MATRICON.TextMatrix(0, 0);
        Printer.CurrentX = 2.2
        Printer.Print CONSVARIAS.MATRICON.TextMatrix(0, 1);
        CX = 8.5
        For I = 0 To 11
                If List1.Selected(I) = True Then
                    Printer.CurrentX = CX
                    Printer.Print CONSVARIAS.MATRICON.TextMatrix(0, (I + 2));
                    If (I = 6) Or (I = 7) Then
                        CX = CX + 5.3
                    Else
                        If (I > -1) And (I < 5) Then
                            CX = CX + 1.6
                        Else
                            If (I = 9) Then
                                CX = CX + 1.3
                            Else
                                CX = CX + 2
                            End If
                        End If
                    End If
                End If
        Next I
        Printer.Print ""
        Printer.Print ""
        For J = 1 To (CONSVARIAS.MATRICON.Rows - 1)
            Printer.CurrentX = 0.8
            Printer.Print CONSVARIAS.MATRICON.TextMatrix(J, 0);
            Printer.CurrentX = 2.2
            Printer.Print CONSVARIAS.MATRICON.TextMatrix(J, 1);
            CX = 8.5
            For I = 0 To 11
                If List1.Selected(I) = True Then
                    Printer.CurrentX = CX
                    Printer.Print CONSVARIAS.MATRICON.TextMatrix(J, (I + 2));
                    If (I = 6) Or (I = 7) Then
                        CX = CX + 5.3
                    Else
                        If (I > -1) And (I < 5) Then
                            CX = CX + 1.6
                        Else
                            If (I = 9) Then
                                CX = CX + 1.3
                            Else
                                CX = CX + 2
                            End If
                        End If
                    End If
                End If
            Next I
            Printer.Print ""
            If (J Mod 67) = 0 Then
                Printer.NewPage
                PAG = PAG + 1
                Printer.CurrentY = 0.5
                Printer.CurrentX = 5.5
                Printer.Print Format(CONSVARIAS.Caption, ">")
                Printer.CurrentY = 1.5
                Printer.CurrentX = 0.8
                Printer.Print CONSVARIAS.Frame1.Caption;
                Printer.CurrentX = 14
                Printer.Print "Fecha: " & Format(Date, "mmm/dd/yyyy");
                Printer.CurrentX = 19
                Printer.Print "Pág." & PAG
                Printer.CurrentX = 0.8
                Printer.Print ini.nombre
                Printer.CurrentY = 2.5
                Printer.CurrentX = 0.8
                Printer.Print CONSVARIAS.MATRICON.TextMatrix(0, 0);
                Printer.CurrentX = 2.2
                Printer.Print CONSVARIAS.MATRICON.TextMatrix(0, 1);
                CX = 8.5
                For I = 0 To 11
                    If List1.Selected(I) = True Then
                        Printer.CurrentX = CX
                        Printer.Print CONSVARIAS.MATRICON.TextMatrix(0, (I + 2));
                        If (I = 6) Or (I = 7) Then
                            CX = CX + 5.3
                        Else
                            If (I > -1) And (I < 5) Then
                                CX = CX + 1.6
                            Else
                                If (I = 9) Then
                                    CX = CX + 1.3
                                Else
                                    CX = CX + 2
                                End If
                            End If
                        End If
                    End If
                Next I
                Printer.Print ""
                Printer.Print ""
            End If
        Next J
        Printer.EndDoc
        Screen.MousePointer = 0
        Unload Me
End If
End Sub

Private Sub Form_Load()
Label1.Caption = 6
End Sub

Private Sub List1_ItemCheck(Item As Integer)
If ((Item < 6) Or (Item > 7)) Then
    I = 1
Else
    I = 3
End If
    If List1.Selected(Item) = False Then
        Label1.Caption = Val(Label1.Caption) + I
    End If
    If (Val(Label1.Caption) < I) Then
        List1.Selected(Item) = False
        If (I = 3) And (Val(Label1.Caption) > 0) Then
            MsgBox "No puede seleccionar este campo"
            Exit Sub
        Else
            MsgBox "No puede seleccionar más campos para imprimir"
            Exit Sub
        End If
    End If
    If List1.Selected(Item) = True Then
        Label1.Caption = Val(Label1.Caption) - I
    End If
End Sub
