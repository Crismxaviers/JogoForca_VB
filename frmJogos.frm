VERSION 5.00
Begin VB.Form FrmJogos 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jogos"
   ClientHeight    =   5685
   ClientLeft      =   2445
   ClientTop       =   1965
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   7500
   Begin VB.Timer tmrLetras 
      Interval        =   500
      Left            =   6960
      Top             =   240
   End
   Begin VB.Label lbllet10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1095
      Left            =   5400
      TabIndex        =   10
      Top             =   3120
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lbllet9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1095
      Left            =   4560
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lbllet8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1095
      Left            =   3600
      TabIndex        =   8
      Top             =   3120
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lbllet7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1095
      Left            =   2640
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label lbllet6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1095
      Left            =   1800
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Label lbllet5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1095
      Left            =   3960
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lbllet4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1095
      Left            =   3120
      TabIndex        =   4
      Top             =   1920
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label lbllet3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1095
      Left            =   4920
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label lbllet2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1095
      Left            =   3840
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label lbllet1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1095
      Left            =   2760
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label lbllet0 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1095
      Left            =   2040
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Menu mnuJogos 
      Caption         =   "&Jogo"
      Begin VB.Menu mnuforca 
         Caption         =   "&Forca"
      End
      Begin VB.Menu mnutraco 
         Caption         =   "-"
      End
      Begin VB.Menu mnusair 
         Caption         =   "Sai&r"
      End
   End
   Begin VB.Menu mnuSobre 
      Caption         =   "&Sobre"
   End
End
Attribute VB_Name = "FrmJogos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub tmrLetras_Timer()
Static letras As Integer

Select Case letras
    Case 0
    lbllet0.Visible = True
    letras = 1
    Beep
    
    Case 1
    lbllet1.Visible = True
    letras = 2
    Beep
    
    Case 2
    lbllet2.Visible = True
    letras = 3
    Beep
    
    Case 3
    lbllet3.Visible = True
    letras = 4
    Beep
    
    Case 4
    lbllet4.Visible = True
    letras = 5
    Beep
    
    Case 5
    lbllet5.Visible = True
    letras = 6
    Beep
    
    Case 6
    lbllet6.Visible = True
    letras = 7
    Beep
    
    Case 7
    lbllet7.Visible = True
    letras = 8
    Beep
    
    Case 8
    lbllet8.Visible = True
    letras = 9
    Beep
    
    Case 9
    lbllet9.Visible = True
    letras = 10
    Beep
    
    Case 10
    lbllet10.Visible = True
    letras = 11
    Beep
    
    Case 11
    LetVisivel (False)
    letras = 0
    End Select
        
End Sub
Private Sub LetVisivel(bool As Boolean)

    lbllet0.Visible = bool
    lbllet1.Visible = bool
    lbllet2.Visible = bool
    lbllet3.Visible = bool
    lbllet4.Visible = bool
    lbllet5.Visible = bool
    lbllet6.Visible = bool
    lbllet7.Visible = bool
    lbllet8.Visible = bool
    lbllet9.Visible = bool
    lbllet10.Visible = bool
    
    
    End Sub

Private Sub mnuforca_Click()

frmForca.Show 1

End Sub

Private Sub mnusair_Click()
End
End Sub

Private Sub mnusobre_Click()
 MsgBox "Autoras:" & Chr(10) & Chr(13) & "   Cristina M. X. Silva" & Chr(10) & Chr(13) & "   Fernanda Lins Bandeira" & Chr(10) & Chr(13) & "Professor: Marcos Junior"
End Sub


