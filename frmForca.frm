VERSION 5.00
Begin VB.Form frmForca 
   BackColor       =   &H80000008&
   Caption         =   "JOGO DA FORCA"
   ClientHeight    =   6885
   ClientLeft      =   2325
   ClientTop       =   1305
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   8025
   Visible         =   0   'False
   Begin VB.CommandButton cmdinicio 
      Caption         =   "Inicio"
      Height          =   495
      Left            =   5640
      MouseIcon       =   "frmForca.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   36
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox txtpalavra 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   5160
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   34
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CommandButton cmdletra 
      BackColor       =   &H80000009&
      Caption         =   "A"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   0
      Left            =   720
      MouseIcon       =   "frmForca.frx":030A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5760
      Width           =   400
   End
   Begin VB.CommandButton cmdletra 
      BackColor       =   &H80000009&
      Caption         =   "Z"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   25
      Left            =   6480
      MouseIcon       =   "frmForca.frx":0614
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   6360
      Width           =   400
   End
   Begin VB.CommandButton cmdletra 
      BackColor       =   &H80000009&
      Caption         =   "Y"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   24
      Left            =   6000
      MouseIcon       =   "frmForca.frx":091E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6360
      Width           =   400
   End
   Begin VB.CommandButton cmdletra 
      BackColor       =   &H80000009&
      Caption         =   "X"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   23
      Left            =   5520
      MouseIcon       =   "frmForca.frx":0C28
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6360
      Width           =   400
   End
   Begin VB.CommandButton cmdletra 
      BackColor       =   &H80000009&
      Caption         =   "W"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   22
      Left            =   5040
      MouseIcon       =   "frmForca.frx":0F32
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6360
      Width           =   400
   End
   Begin VB.CommandButton cmdletra 
      BackColor       =   &H80000009&
      Caption         =   "V"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   21
      Left            =   4560
      MouseIcon       =   "frmForca.frx":123C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6360
      Width           =   400
   End
   Begin VB.CommandButton cmdletra 
      BackColor       =   &H80000009&
      Caption         =   "U"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   20
      Left            =   4080
      MouseIcon       =   "frmForca.frx":1546
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6360
      Width           =   400
   End
   Begin VB.CommandButton cmdletra 
      BackColor       =   &H80000009&
      Caption         =   "T"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   19
      Left            =   3600
      MouseIcon       =   "frmForca.frx":1850
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6360
      Width           =   400
   End
   Begin VB.CommandButton cmdletra 
      BackColor       =   &H80000009&
      Caption         =   "S"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   18
      Left            =   3120
      MouseIcon       =   "frmForca.frx":1B5A
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6360
      Width           =   400
   End
   Begin VB.CommandButton cmdletra 
      BackColor       =   &H80000009&
      Caption         =   "R"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   17
      Left            =   2640
      MouseIcon       =   "frmForca.frx":1E64
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6360
      Width           =   400
   End
   Begin VB.CommandButton cmdletra 
      BackColor       =   &H80000009&
      Caption         =   "Q"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   16
      Left            =   2160
      MouseIcon       =   "frmForca.frx":216E
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6360
      Width           =   400
   End
   Begin VB.CommandButton cmdletra 
      BackColor       =   &H80000009&
      Caption         =   "P"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   15
      Left            =   1680
      MouseIcon       =   "frmForca.frx":2478
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6360
      Width           =   400
   End
   Begin VB.CommandButton cmdletra 
      BackColor       =   &H80000009&
      Caption         =   "O"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   14
      Left            =   1200
      MouseIcon       =   "frmForca.frx":2782
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6360
      Width           =   400
   End
   Begin VB.CommandButton cmdletra 
      BackColor       =   &H80000009&
      Caption         =   "N"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   13
      Left            =   720
      MouseIcon       =   "frmForca.frx":2A8C
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6360
      Width           =   400
   End
   Begin VB.CommandButton cmdletra 
      BackColor       =   &H80000009&
      Caption         =   "M"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   12
      Left            =   6480
      MouseIcon       =   "frmForca.frx":2D96
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5760
      Width           =   400
   End
   Begin VB.CommandButton cmdletra 
      BackColor       =   &H80000009&
      Caption         =   "L"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   11
      Left            =   6000
      MouseIcon       =   "frmForca.frx":30A0
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5760
      Width           =   400
   End
   Begin VB.CommandButton cmdletra 
      BackColor       =   &H80000009&
      Caption         =   "K"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   10
      Left            =   5520
      MouseIcon       =   "frmForca.frx":33AA
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5760
      Width           =   400
   End
   Begin VB.CommandButton cmdletra 
      BackColor       =   &H80000009&
      Caption         =   "J"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   9
      Left            =   5040
      MouseIcon       =   "frmForca.frx":36B4
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5760
      Width           =   400
   End
   Begin VB.CommandButton cmdletra 
      BackColor       =   &H80000009&
      Caption         =   "I"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   8
      Left            =   4560
      MouseIcon       =   "frmForca.frx":39BE
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5760
      Width           =   400
   End
   Begin VB.CommandButton cmdletra 
      BackColor       =   &H80000009&
      Caption         =   "H"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   7
      Left            =   4080
      MouseIcon       =   "frmForca.frx":3CC8
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5760
      Width           =   400
   End
   Begin VB.CommandButton cmdletra 
      BackColor       =   &H80000009&
      Caption         =   "G"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   6
      Left            =   3600
      MouseIcon       =   "frmForca.frx":3FD2
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5760
      Width           =   400
   End
   Begin VB.CommandButton cmdletra 
      BackColor       =   &H80000009&
      Caption         =   "F"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   5
      Left            =   3120
      MouseIcon       =   "frmForca.frx":42DC
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5760
      Width           =   400
   End
   Begin VB.CommandButton cmdletra 
      BackColor       =   &H80000009&
      Caption         =   "E"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   4
      Left            =   2640
      MouseIcon       =   "frmForca.frx":45E6
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   400
   End
   Begin VB.CommandButton cmdletra 
      BackColor       =   &H80000009&
      Caption         =   "D"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   3
      Left            =   2160
      MouseIcon       =   "frmForca.frx":48F0
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   400
   End
   Begin VB.CommandButton cmdletra 
      BackColor       =   &H80000009&
      Caption         =   "C"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   2
      Left            =   1680
      MouseIcon       =   "frmForca.frx":4BFA
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5760
      Width           =   400
   End
   Begin VB.CommandButton cmdletra 
      BackColor       =   &H80000009&
      Caption         =   "B"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   1
      Left            =   1200
      MouseIcon       =   "frmForca.frx":4F04
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5760
      Width           =   400
   End
   Begin VB.Image imgvenc 
      Height          =   4530
      Left            =   120
      Picture         =   "frmForca.frx":520E
      Top             =   600
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.Image imgBoneco 
      Height          =   4530
      Index           =   6
      Left            =   120
      Picture         =   "frmForca.frx":7B5F
      Top             =   600
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.Image imgBoneco 
      Height          =   4530
      Index           =   5
      Left            =   120
      Picture         =   "frmForca.frx":C68A
      Top             =   600
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.Image imgBoneco 
      Height          =   4530
      Index           =   4
      Left            =   120
      Picture         =   "frmForca.frx":EDE5
      Top             =   600
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.Image imgBoneco 
      Height          =   4530
      Index           =   3
      Left            =   120
      Picture         =   "frmForca.frx":112D9
      Top             =   600
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.Image imgBoneco 
      Height          =   4530
      Index           =   2
      Left            =   240
      Picture         =   "frmForca.frx":1345D
      Top             =   600
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.Image imgBoneco 
      Height          =   4530
      Index           =   1
      Left            =   120
      Picture         =   "frmForca.frx":151CF
      Top             =   600
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.Image imgBoneco 
      Height          =   4530
      Index           =   0
      Left            =   240
      Picture         =   "frmForca.frx":16AC0
      Top             =   600
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jogo da Forca"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   435
      Left            =   3960
      TabIndex        =   37
      Top             =   240
      Width           =   2445
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Palavra: "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   225
      Left            =   5160
      TabIndex        =   35
      Top             =   1560
      Width           =   720
   End
   Begin VB.Label LBLLETRA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Index           =   7
      Left            =   7080
      TabIndex        =   32
      Top             =   5040
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label LBLLETRA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Index           =   6
      Left            =   6720
      TabIndex        =   31
      Top             =   5040
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label LBLLETRA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Index           =   5
      Left            =   6360
      TabIndex        =   30
      Top             =   5040
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label LBLLETRA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Index           =   4
      Left            =   6000
      TabIndex        =   29
      Top             =   5040
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label LBLLETRA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Index           =   3
      Left            =   5640
      TabIndex        =   28
      Top             =   5040
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label LBLLETRA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Index           =   2
      Left            =   5280
      TabIndex        =   27
      Top             =   5040
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label LBLLETRA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Index           =   1
      Left            =   4920
      TabIndex        =   26
      Top             =   5040
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label LBLLETRA 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Index           =   0
      Left            =   4560
      TabIndex        =   25
      Top             =   5040
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Menu mnujogo 
      Caption         =   "&Jogo"
      Begin VB.Menu mnunovo 
         Caption         =   "&novo"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuinicio 
         Caption         =   "&inicio"
      End
      Begin VB.Menu mnutraco 
         Caption         =   "-"
      End
      Begin VB.Menu mnusair 
         Caption         =   "Sai&r"
      End
   End
   Begin VB.Menu mnuOpcoes 
      Caption         =   "&Opções"
      Begin VB.Menu mnudupla 
         Caption         =   "&Dupla"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnusolitario 
         Caption         =   "&Solitario"
      End
   End
   Begin VB.Menu mnusobre 
      Caption         =   "&Sobre"
   End
End
Attribute VB_Name = "frmForca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Erros As Byte
Dim carros(20) As String
Private Sub Form_Load()
    carros(0) = "VECTRA"
    carros(1) = "OMEGA"
    carros(2) = "SAVEIRO"
    carros(3) = "FIESTA"
    carros(4) = "PALIO"
    carros(5) = "CORSA"
    carros(6) = "SANTANA"
    carros(7) = "SIENA"
    carros(8) = "GOLF"
    carros(9) = "CIVIC"
    carros(10) = "FERNANDA"
    carros(11) = "CRISTINA"
    carros(12) = "JULIANA"
    carros(13) = "ALINE"
    carros(14) = "SILVANA"
    carros(15) = "SANDRA"
    carros(16) = "ADRIANA"
    carros(17) = "MARCOS"
    carros(18) = "FABIO"
    carros(19) = "DANIEL"
End Sub

Private Sub txtpalavra_KeyPress(KeyAscii As Integer)
    'abaixo está sendo verificado se a tecla digitada é:
    'A até Z = 65 até 90
    'a ate´z = 97 até 122
    ' Backspace = 8
    If Not ((KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or (KeyAscii = 8)) Then
        KeyAscii = 0
    End If
End Sub
Private Sub mnuinicio_Click()
    'verifica se algo foi digitado
    If mnusolitario.Checked = True Then
        Randomize
        txtpalavra.Text = carros(Int(Rnd * 19))
        inicia
    ElseIf Len(txtpalavra.Text) < 1 Then
        MsgBox ("digite a palavra")
        txtpalavra.SetFocus
    Else
        inicia
        txtpalavra.Locked = True
    End If
End Sub
Private Sub inicia()
    Dim i As Byte
    Habilita (True) 'Habilita os botões
    'Deixa visivel um label para cada letra da palavra digitada
    For i = 0 To Len(txtpalavra.Text) - 1
        LBLLETRA(i).Visible = True
    Next
    cmdinicio.Enabled = False
End Sub
Private Sub Habilita(logica As Boolean)
    Dim i As Byte
    For i = 0 To 25
        cmdletra(i).Enabled = logica
    Next
End Sub
Private Sub cmdinicio_Click()
    mnuinicio_Click
End Sub

Private Sub LigaOpc(logica As Boolean)
    mnudupla.Checked = Not logica
    mnusolitario.Checked = logica
    txtpalavra.Locked = logica
    mnuNovo_click
End Sub
Private Sub mnuDupla_Click()
    LigaOpc (False)
End Sub

Private Sub mnuSolitario_click()
    LigaOpc (True)
End Sub

Private Sub cmdletra_Click(Index As Integer)
    Dim Letr As String, palavra As String
    Dim i As Byte
    Dim Acerto As Boolean, Final As Boolean
    
    'Transforma letra em maiuscula
        Letr = cmdletra(Index).Caption
        palavra = UCase(txtpalavra.Text)
    
    'Procura a letra digitada dentro da palavra
    'caso o encontre apresenta o mesmo dentro da label
    For i = 1 To Len(palavra)
        If Mid(palavra, i, 1) = Letr Then
            LBLLETRA(i - 1).Caption = Letr
            Acerto = True
        End If
    Next
    
    'desabilita letra ja usada
    cmdletra(Index).Enabled = False
    
    'caso o jogador tenha errado apresenta parte do boneco
    If Not Acerto Then
        imgBoneco(Erros).Visible = True
        Erros = Erros + 1
    End If
    
    If Erros = 6 Then
       Finaliza ("Infelizmente vc perdeu! Tente de novo")
        imgBoneco(6).Visible = True
    End If
    
    Final = True
    
    For i = 0 To Len(palavra) - 1
        If LBLLETRA(i).Caption = "_" Then
            Final = False
        End If
    Next
    
    'Caso afirmativo não apresenta a morte....
    'apresente o boneco de corpo inteiro
    If Final Then
        imgBoneco(0).Visible = False
        For i = 1 To 6
            imgBoneco(i).Visible = True
            imgvenc.Visible = True
        Next
        Finaliza ("Parabens! Vc descobriu a palavra e ganhou")
        imgvenc.Visible = True
        
    End If
End Sub
Private Sub Finaliza(mens As String)
    'caso mens seja nulo indica q foi iniciado novo jogo
    If Not (mens = "") Then MsgBox mens, vbInformation, "jogo da forca"
    Habilita (False)
    Erros = 0
    mnuinicio.Enabled = False
    cmdinicio.Enabled = False
End Sub

Private Sub mnuNovo_click()
    Dim i As Byte
    
    Finaliza ("")
    Habilita (False)
    
    For i = 0 To 7
        LBLLETRA(i).Caption = "_"
        LBLLETRA(i).Visible = False
    Next
        
    For i = 0 To 6
        imgBoneco(i).Visible = False
         imgvenc.Visible = False
    Next
    imgvenc.Visible = False
    mnuinicio.Enabled = True
    cmdinicio.Enabled = True
    
  With txtpalavra
  .Text = ""
        If mnusolitario.Checked = False Then
            .SetFocus
            .Locked = False
        End If
    End With
End Sub

Private Sub mnusair_Click()
    Unload Me
End Sub

Private Sub mnusobre_Click()
    MsgBox "Autoras:" & Chr(10) & Chr(13) & "   Cristina M. X. Silva" & Chr(10) & Chr(13) & "   Fernanda Lins Bandeira" & Chr(10) & Chr(13) & "Professor: Marcos Junior"
End Sub





















