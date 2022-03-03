VERSION 5.00
Begin VB.Form FORMATO2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "FORMATO2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1440
      Left            =   120
      TabIndex        =   7
      Top             =   1215
      Width           =   3735
      Begin VB.CheckBox Check6 
         Caption         =   "Si"
         Height          =   255
         Left            =   1800
         TabIndex        =   16
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Si"
         Height          =   195
         Left            =   1800
         TabIndex        =   11
         Top             =   600
         Width           =   615
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Si"
         Height          =   195
         Left            =   1800
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Ver promoción?"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Firma del Secretario?"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Firma del Rector?"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1245
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   2880
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.CommandButton Command3 
         Caption         =   "-"
         Height          =   320
         Left            =   3390
         TabIndex        =   14
         Top             =   150
         Width           =   195
      End
      Begin VB.CommandButton Command2 
         Caption         =   "+"
         Height          =   320
         Left            =   3210
         TabIndex        =   13
         Top             =   150
         Width           =   195
      End
      Begin VB.TextBox Txt_Espa 
         Height          =   320
         Left            =   2865
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "0"
         Top             =   150
         Width           =   345
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Encabezado"
         Height          =   195
         Left            =   1560
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Sin Final"
         Height          =   255
         Left            =   1560
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "4 Periodos"
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Sin Formato"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Con Formato"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FORMATO2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function CortaObs(Observacion As String)
Dim Recorrer As Integer, Cortar() As String, XSuma As Single
Cortar = Split(Observacion, " ")
XSuma = 0
Printer.CurrentX = 1.3
For Recorrer = 0 To UBound(Cortar)
    XSuma = XSuma + Printer.TextWidth(Cortar(Recorrer))
    If XSuma <= 17 Then
        Printer.Print Cortar(Recorrer);
        Printer.Print " ";
    Else
        XSuma = Printer.TextWidth(Cortar(Recorrer))
        Printer.Print ""
        Printer.CurrentX = 1.3
        Printer.Print Cortar(Recorrer);
        Printer.Print " ";
    End If
Next Recorrer
Printer.Print ""
End Function

Private Sub Check3_Click()
If Check3.Value = 0 Then
    Txt_Espa.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
Else
    Txt_Espa.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
End If
End Sub

Private Sub Command1_Click()
Dim DsFinal As String, XMax As Single
RESP = MsgBox("DESEA IMPRIMIR EL INFORME FINAL DEL ESTUDIANTE?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR INFORME")
If RESP = vbYes Then
    Screen.MousePointer = 11
    If Check2.Value = 1 Then
        If Check1.Value = 1 Then
            CX = 0.3
        Else
            CX = 0.5
        End If
    Else
        CX = 0
    End If
    Printer.ScaleMode = 7
    If Option2.Value = True Then
        Printer.Font.Size = 14
        Printer.CurrentY = 1
        Printer.CurrentX = 10.2 - ((Len(ini.nombre) / 3.3) / 2)
        Printer.FontBold = True
        Printer.Print ini.nombre
        Printer.CurrentX = 8.2
        Printer.Print "INFORME FINAL"
        Printer.FontBold = False
    End If
    If Check3.Value = 1 Then
        Printer.CurrentY = Val(Txt_Espa.Text) / 10 + 2.7
    Else
        Printer.CurrentY = 2.7
    End If
    Printer.Font.Size = 10
    'If Option1.Value = True Then
    '    Printer.CurrentX = 3.7
    '    Printer.Print RESUFINA.Frame1.Caption;
    'Else
        Printer.CurrentX = 1.3
        Printer.Print "ESTUDIANTE: " & RESUFINA.Frame1.Caption;
    'End If
    'If Option1.Value = True Then
    '    Printer.CurrentX = 17
    '    Printer.Print Year(Date)
    'Else
        Printer.CurrentX = 18.5
        Printer.Print "AÑO:" & Year(Date)
    'End If
    'If Option1.Value = True Then
    '    Printer.CurrentX = 2.9
    '    Printer.Print RE22;
    'Else
        Printer.CurrentX = 1.3
        Printer.Print "GRADO: " & RE22;
    'End If
    'If Option1.Value = True Then
    '    Printer.CurrentX = 7.6
    '    Printer.Print RESUFINA.Label4.Caption;
    'Else
        Printer.CurrentX = 6
        Printer.Print "GRUPO: " & RESUFINA.Label4.Caption;
    'End If
    'If Option1.Value = True Then
    '    Printer.CurrentX = 18.5
    '    Printer.Print Date
    'Else
        Printer.CurrentX = 17
        Printer.Print "FECHA: " & Date
    'End If
    Printer.Print ""
    Printer.Line (1, 4)-(20.3, 4)
    Printer.Print ""
    Printer.CurrentX = 1.3
    Printer.Print "M A T E R I A S";
    Printer.CurrentX = 7.5
    Printer.Print "I.H.";
    If Check1.Value = 1 Then
        Printer.CurrentX = 8.7
        Printer.Print "I PERIODO";
        Printer.CurrentX = 11.4 + CX
        Printer.Print "II PERIODO";
        Printer.CurrentX = 14 + (2 * CX)
        Printer.Print "III PERIODO";
        Printer.CurrentX = 16.5 + (4 * CX)
        Printer.Print "IV PERIODO";
    Else
        Printer.CurrentX = 9
        Printer.Print "I PERIODO";
        Printer.CurrentX = 12.3 + (2 * CX)
        Printer.Print "II PERIODO";
        Printer.CurrentX = 15.7 + (3 * CX)
        Printer.Print "III PERIODO";
    End If
    If Check2.Value = 0 Then
        Printer.CurrentX = 19
        Printer.Print "FINAL"
    Else
        Printer.Print ""
    End If
    Printer.Print ""
    Printer.Line (1, Printer.CurrentY)-(20.3, Printer.CurrentY)
    Printer.Print ""
    For I = 1 To (RESUFINA.MATI50.Rows - 1)
        RESUFINA.MATI50.Row = I
        RESUFINA.MATI50.Col = 0
        Printer.CurrentX = 1.3
        Printer.Print RESUFINA.MATI50.Text;
        RESUFINA.MATI50.Col = 1
        Printer.CurrentX = 7.7
        Printer.Print RESUFINA.MATI50.Text;
        If Check1.Value = 1 Then
            'Printer.Print RESUFINA.MATI50.Text;
            RESUFINA.MATI50.Col = 2
            Printer.CurrentX = 9.2
            Printer.Print RESUFINA.MATI50.Text;
            RESUFINA.MATI50.Col = 3
            Printer.CurrentX = 12 + CX
            Printer.Print RESUFINA.MATI50.Text;
            RESUFINA.MATI50.Col = 4
            Printer.CurrentX = 14.6 + (2 * CX)
            Printer.Print RESUFINA.MATI50.Text;
            RESUFINA.MATI50.Col = 5
            Printer.CurrentX = 17.2 + (4 * CX)
            Printer.Print RESUFINA.MATI50.Text;
        Else
            'Printer.Print RESUFINA.MATI50.Text;
            RESUFINA.MATI50.Col = 2
            Printer.CurrentX = 9.5
            Printer.Print RESUFINA.MATI50.Text;
            RESUFINA.MATI50.Col = 3
            Printer.CurrentX = 12.8 + (2 * CX)
            Printer.Print RESUFINA.MATI50.Text;
            RESUFINA.MATI50.Col = 4
            Printer.CurrentX = 16.4 + (3 * CX)
            Printer.Print RESUFINA.MATI50.Text;
        End If
        If Check2.Value = 0 Then
            RESUFINA.MATI50.Col = 6
            Printer.CurrentX = 19.2
            Printer.Print RESUFINA.MATI50.Text
        Else
            Printer.Print ""
        End If
    Next I
    Printer.Print ""
    Printer.Line (1, Printer.CurrentY)-(20.3, Printer.CurrentY)
    Printer.Line (1, 4)-(1, Printer.CurrentY)
    Printer.Line (7.3, 4)-(7.3, Printer.CurrentY)
    Printer.Line (8.2, 4)-(8.2, Printer.CurrentY)
    If Check1.Value = 1 Then
        If Check2.Value = 0 Then
            Printer.Line (11, 4)-(11, Printer.CurrentY)
            Printer.Line (13.7, 4)-(13.7, Printer.CurrentY)
            Printer.Line (16.3, 4)-(16.3, Printer.CurrentY)
            Printer.Line (18.8, 4)-(18.8, Printer.CurrentY)
        Else
            Printer.Line (11.2, 4)-(11.2, Printer.CurrentY)
            Printer.Line (14.2, 4)-(14.2, Printer.CurrentY)
            Printer.Line (17.2, 4)-(17.2, Printer.CurrentY)
        End If
    Else
        If Check2.Value = 0 Then
            Printer.Line (11.6, 4)-(11.6, Printer.CurrentY)
            Printer.Line (15, 4)-(15, Printer.CurrentY)
            Printer.Line (18.5, 4)-(18.5, Printer.CurrentY)
        Else
            Printer.Line (12.2, 4)-(12.2, Printer.CurrentY)
            Printer.Line (16.2, 4)-(16.2, Printer.CurrentY)
        End If
    End If
    Printer.Line (20.3, 4)-(20.3, Printer.CurrentY)
    Printer.Print ""
    ' EXPLICACION CONVENCIONES
    Printer.CurrentX = 1.3
    Printer.Font.Size = 7
    Printer.Print Trim(confdesemp.desemp(1)) & "=Desempeño Superior (";
    Printer.Print confdesemp.rango(1) + 1 & "% - 100%),  ";
    Printer.Print Trim(confdesemp.desemp(2)) & "=Desempeño Alto (";
    Printer.Print confdesemp.rango(2) + 1 & "% - " & confdesemp.rango(1) & "%),  ";
    Printer.Print Trim(confdesemp.desemp(3)) & "=Desempeño Básico (";
    Printer.Print confdesemp.rango(3) + 1 & "% - " & confdesemp.rango(2) & "%),  ";
    Printer.Print Trim(confdesemp.desemp(4));
    If Trim(confdesemp.desemp(4)) = "*LEP" Then
        Printer.Print "=Logro en Proceso (";
    Else
        Printer.Print "=Desempeño Bajo (";
    End If
    Printer.Print "0% - " & confdesemp.rango(3) & "%)."
    Printer.Font.Size = 10
    Printer.Print ""
    Printer.Print ""
    'Verificar para mostrar la información de promoción.
    If Check6.Value = 1 Then
        Printer.CurrentX = 1.3
        Printer.Print RESUFINA.Label3.Caption
    Else
        Printer.Print ""
    End If
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.CurrentX = 1.3
    Printer.Print "Observaciones:"
    Printer.Print ""
    J = 0
    If (Dir(Ruta & "lrf" & RESUFINA.Label4.Caption & ".lrf") <> "") And (Dir(Ruta & "orf" & RESUFINA.Label4.Caption & ".orf") <> "") Then
        cona = 0
        Open Ruta & "lrf" & RESUFINA.Label4.Caption & ".lrf" For Random As #NAR Len = Len(leyfin)
        While Not EOF(NAR)
            cona = cona + 1
            Get #NAR, cona, leyfin
            If Val(leyfin.num_carnet) = Val(RESUFINA.Label5.Caption) Then
                NAR = FreeFile
                Open Ruta & "orf" & RESUFINA.Label4.Caption & ".orf" For Random As #NAR Len = Len(obsfin)
                For I = 1 To 5
                    If leyfin.fnob(I) <> 0 Then
                        Get #NAR, leyfin.fnob(I), obsfin
                        Printer.CurrentX = 1.3
                        
                        XMax = Printer.TextWidth(Trim(obsfin))
                        If XMax > 17 Then
                            CortaObs (Trim(obsfin))
                        Else
                            Printer.Print Trim(obsfin)
                        End If
                    Else
                        J = J + 1
                    End If
                Next I
                Close #NAR
                Close #(NAR - 1)
                NAR = NAR - 1
                GoTo saobfi
            End If
        Wend
        Close #NAR
    Else
        J = 5
    End If
saobfi:
    If J <> 0 Then
        For I = 1 To J
            Printer.Print ""
        Next I
    End If
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    Printer.Print ""
    'If Option2.Value = True Then
        'Y = Printer.CurrentY
        If Check4.Value = 1 Then
            'Printer.CurrentY = Y
            Printer.Line (3, Printer.CurrentY)-(8.5, Printer.CurrentY)
            Printer.CurrentY = Printer.CurrentY + 0.2
            Printer.CurrentX = 5
            Printer.Print vini.VRector;
        End If
        If Check5.Value = 1 Then
            'Printer.CurrentY = Y
            Printer.Line (13, Printer.CurrentY)-(18.5, Printer.CurrentY)
            Printer.CurrentY = Printer.CurrentY + 0.2
            Printer.CurrentX = 14.3
            Printer.Print "Firma del Secretario"
        End If
    'End If
    Printer.EndDoc
    'Printer.Font.Size = 8
    Unload Me
End If
Screen.MousePointer = 0
End Sub

Private Sub Command2_Click()
Txt_Espa.Text = Txt_Espa.Text + 1
End Sub

Private Sub Command3_Click()
Txt_Espa.Text = Txt_Espa.Text - 1
End Sub

Private Sub Form_Load()
Check4.Value = 1
Check5.Value = 1
Option2.Value = True
Txt_Espa.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
End Sub
