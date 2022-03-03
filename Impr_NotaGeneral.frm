VERSION 5.00
Begin VB.Form Impr_NotaGeneral 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir reporte general"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones de impresión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      Begin VB.CheckBox Check2 
         Caption         =   "Imprimir sólo estudiantes con materias perdidas"
         Height          =   435
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Imprimir puesto del estudiante"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Impr_NotaGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
RESP = MsgBox("DESEA IMPRIMIR ESTA INFORMACION?", vbYesNo + vbQuestion + vbDefaultButton2, "IMPRIMIR")
    If RESP = vbYes Then
        Screen.MousePointer = 11
        Printer.Orientation = 2
        Printer.PaperSize = 14
        NAR = FreeFile
        Open Ruta & "inicial.edu" For Input As #NAR
        Input #NAR, ini.ciudad, ini.nombre, ini.modalidad, ini.Telefono, ini.Rector, ini.secretario
        Close #NAR
        Printer.ScaleMode = 7
        Printer.Font.Size = 9
        Printer.CurrentY = 0.5
        Printer.CurrentX = 17.2 - ((Len(ini.nombre) / 3.3) / 2)
        Printer.FontBold = True
        Printer.Print ini.nombre
'        Printer.CurrentX = 16.5 - (Len(ini.Rector) / 5.2) / 2
'        Printer.Print ini.Rector
        Printer.Print ""
        Printer.CurrentX = 17.2 - ((Len(TituloPrint) / 4) / 2)
        Printer.Print TituloPrint
        Printer.FontBold = False
        Printer.Print ""
        Printer.Font.Size = 8
        Printer.CurrentX = 1
        Printer.Print "GRUPO: " & NOTA_GENRL.Frame1.Caption;
        Printer.CurrentX = 8
        Printer.Print "PERIODO: " & NOTA_GENRL.Combo2.Text & AcumulaPrint;
        Printer.CurrentX = 24
        Printer.Print "FECHA: " & Format(Date, "mmm/dd/yyyy")
        Printer.CurrentY = 3
        CX = 8
        For I = 3 To (NOTA_GENRL.MATI126.Cols - 1)
            Printer.CurrentX = CX
            'If I <> (nota_genrl.MATI126.Cols - 1) Then
                Printer.Print Trim(Right(NOTA_GENRL.MATI126.TextMatrix(0, I), 5));
            'Else
            '    Printer.Print "TTL";
            'End If
            CX = CX + 1.17
        Next I
        Printer.Print ""
        
        Printer.CurrentX = 1
        Printer.Print "CD";
        Printer.CurrentX = 1.5
        Printer.Print "APELLIDOS Y NOMBRES";
        
        
        CX = 8
        For I = 3 To (NOTA_GENRL.MATI126.Cols - 1)
            Printer.CurrentX = CX
            If I < (NOTA_GENRL.MATI126.Cols - 2) Then
                Printer.Print Left(NOTA_GENRL.MATI126.TextMatrix(1, I), 3) & Right(NOTA_GENRL.MATI126.TextMatrix(1, I), 4);
            Else
                If I = (NOTA_GENRL.MATI126.Cols - 2) Then
                    Printer.Print "TTL";
                End If
                If I = (NOTA_GENRL.MATI126.Cols - 1) Then
                    If Check1.Value = 1 Then
                        Printer.Print "PTO";
                    End If
                End If
            End If
            CX = CX + 1.17
        Next I
        Printer.Print ""
        Printer.Print ""
        If Check1.Value = 1 Then
            TT = NOTA_GENRL.MATI126.Cols - 1
        Else
            TT = NOTA_GENRL.MATI126.Cols - 2
        End If
        For I = 2 To (NOTA_GENRL.MATI126.Rows - 1)
            If (Check2.Value = 1) And (Val(NOTA_GENRL.MATI126.TextMatrix(I, NOTA_GENRL.MATI126.Cols - 2)) = 0) And (I <> NOTA_GENRL.MATI126.Rows - 1) Then
                Printer.Print "";
            Else
                Printer.CurrentX = 1
                Printer.Print NOTA_GENRL.MATI126.TextMatrix(I, 0);
                Printer.CurrentX = 1.5
                Printer.Print Left(NOTA_GENRL.MATI126.TextMatrix(I, 1), 35);
                CX = 8
            End If
            For J = 3 To (TT)
                If (Check2.Value = 1) And (Val(NOTA_GENRL.MATI126.TextMatrix(I, NOTA_GENRL.MATI126.Cols - 2)) = 0) And (I <> NOTA_GENRL.MATI126.Rows - 1) Then
                    Exit For
                End If
                Printer.CurrentX = CX
                If Val(NOTA_GENRL.MATI126.TextMatrix(I, J)) < 70 Then
                    Printer.FontBold = True
                End If
                Printer.Print NOTA_GENRL.MATI126.TextMatrix(I, J);
                Printer.FontBold = False
                CX = CX + 1.17
            Next J
            If (Check2.Value = 1) And (Val(NOTA_GENRL.MATI126.TextMatrix(I, NOTA_GENRL.MATI126.Cols - 2)) = 0) And (I <> NOTA_GENRL.MATI126.Rows - 1) Then
                Printer.Print "";
            Else
                Printer.Print ""
            End If
        Next I
        Printer.EndDoc
        Printer.Orientation = 1
        Printer.PaperSize = 1
        Screen.MousePointer = 0
    End If
    Unload Me
End Sub
