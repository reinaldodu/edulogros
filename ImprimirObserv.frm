VERSION 5.00
Begin VB.Form ImprimirObserv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir Observaciones y/o Logros"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4590
   Icon            =   "ImprimirObserv.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4590
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccione el tipo de reporte"
      ForeColor       =   &H00800000&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.OptionButton Option3 
         Caption         =   "Logros y Observaciones"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Sólo Observaciones"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sólo Logros"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "ImprimirObserv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function CortaObs(Observacion As String)
Dim Recorrer As Integer, Cortar() As String, XSuma As Single
Cortar = Split(Observacion, " ")
XSuma = 0
Printer.CurrentX = 2
For Recorrer = 0 To UBound(Cortar)
    XSuma = XSuma + Printer.TextWidth(Cortar(Recorrer))
    If XSuma <= 16 Then
        Printer.FontSize = 8
        Printer.Print Cortar(Recorrer);
        Printer.Print " ";
    Else
        XSuma = Printer.TextWidth(Cortar(Recorrer))
        Printer.Print ""
        Call SaltaLinea
        Printer.FontSize = 8
        Printer.CurrentX = 2
        Printer.Print Cortar(Recorrer);
        Printer.Print " ";
    End If
Next Recorrer
Printer.Print ""
End Function

Private Function SaltaLinea()
If (Printer.CurrentY > 26) Then
   Printer.NewPage
   Call Encabezado
End If
End Function

Private Function Encabezado()
    Printer.ScaleMode = 7
    PAG = Printer.Page
    Printer.Font.Size = 10
    Printer.CurrentY = 1
    Printer.CurrentX = 18.5
    Printer.Print "Pág." & PAG
    Printer.CurrentX = 1
    Printer.Print COPEGA.Frame1.Caption
    Printer.CurrentX = 1
    Printer.Print ini.nombre;
    Printer.CurrentX = 16.5
    Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
    Printer.Print ""
    Printer.CurrentX = 1
    Printer.Print "No.";
    Printer.CurrentX = 2
    Printer.Print "OBSERVACIONES Y/O LOGROS"
    Printer.Print ""
    Printer.Font.Size = 8
End Function

Private Sub Command1_Click()
Dim CuentaLinea As Integer, XMax As Single
If Val(COPEGA.Label1.Caption) = 0 Then
    MsgBox "NO EXISTE INFORMACION PARA IMPRIMIR", 16, "IMPRIMIR"
    Exit Sub
End If
RESP = MsgBox("DESEA IMPRIMIR ESTA INFORMACION?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR")
If RESP = vbYes Then
    Screen.MousePointer = 11
    NAR = FreeFile
    Open Ruta & "inicial.edu" For Input As #NAR
    Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
    Close #NAR
    Call Encabezado
    CuentaLinea = 0
    For I = 1 To COPEGA.MATU20.Rows - 1
        If Option1.Value = True Then
            If Trim(COPEGA.MATU20.TextMatrix(I, 1)) = "L" Then
                CuentaLinea = CuentaLinea + 1
                Call SaltaLinea
                Printer.CurrentX = 1
                Printer.Print CuentaLinea;
                Printer.CurrentX = 2
                XMax = Printer.TextWidth(Trim(COPEGA.MATU20.TextMatrix(I, 1)) & " - " & Trim(COPEGA.MATU20.TextMatrix(I, 2)))
                If XMax > 16 Then
                    CortaObs (Trim(COPEGA.MATU20.TextMatrix(I, 1)) & " - " & Trim(COPEGA.MATU20.TextMatrix(I, 2)))
                Else
                    Printer.Print Trim(COPEGA.MATU20.TextMatrix(I, 1)) & " - " & Trim(COPEGA.MATU20.TextMatrix(I, 2))
                End If
            End If
        End If
        If Option2.Value = True Then
            If Trim(COPEGA.MATU20.TextMatrix(I, 1)) <> "L" Then
                CuentaLinea = CuentaLinea + 1
                Call SaltaLinea
                Printer.CurrentX = 1
                Printer.Print CuentaLinea;
                Printer.CurrentX = 2
                XMax = Printer.TextWidth(Trim(COPEGA.MATU20.TextMatrix(I, 1)) & " - " & Trim(COPEGA.MATU20.TextMatrix(I, 2)))
                If XMax > 16 Then
                    CortaObs (Trim(COPEGA.MATU20.TextMatrix(I, 1)) & " - " & Trim(COPEGA.MATU20.TextMatrix(I, 2)))
                Else
                    Printer.Print Trim(COPEGA.MATU20.TextMatrix(I, 1)) & " - " & Trim(COPEGA.MATU20.TextMatrix(I, 2))
                End If
            End If
        End If
        If Option3.Value = True Then
            CuentaLinea = CuentaLinea + 1
                Call SaltaLinea
                Printer.CurrentX = 1
                Printer.Print CuentaLinea;
                Printer.CurrentX = 2
                XMax = Printer.TextWidth(Trim(COPEGA.MATU20.TextMatrix(I, 1)) & " - " & Trim(COPEGA.MATU20.TextMatrix(I, 2)))
                If XMax > 16 Then
                    CortaObs (Trim(COPEGA.MATU20.TextMatrix(I, 1)) & " - " & Trim(COPEGA.MATU20.TextMatrix(I, 2)))
                Else
                    Printer.Print Trim(COPEGA.MATU20.TextMatrix(I, 1)) & " - " & Trim(COPEGA.MATU20.TextMatrix(I, 2))
                End If
        End If
    Next I
    Printer.EndDoc
    Screen.MousePointer = 0
End If
Unload Me
End Sub

Private Sub Form_Load()
Option1.Value = True
End Sub
