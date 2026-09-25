VERSION 5.00
Begin VB.Form Pisteikkuna 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pistetilanne"
   ClientHeight    =   2730
   ClientLeft      =   4155
   ClientTop       =   4860
   ClientWidth     =   5355
   ControlBox      =   0   'False
   Icon            =   "Pisteikkuna.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2730
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   340
      Left            =   4080
      TabIndex        =   0
      Top             =   330
      Width           =   900
   End
   Begin VB.Label Pisteet4 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   12
      Left            =   3315
      TabIndex        =   56
      Top             =   2475
      Width           =   225
   End
   Begin VB.Label Pisteet4 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   11
      Left            =   3315
      TabIndex        =   55
      Top             =   2310
      Width           =   225
   End
   Begin VB.Label Pisteet4 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   10
      Left            =   3315
      TabIndex        =   54
      Top             =   2145
      Width           =   225
   End
   Begin VB.Label Pisteet4 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   9
      Left            =   3315
      TabIndex        =   53
      Top             =   1980
      Width           =   225
   End
   Begin VB.Label Pisteet4 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   8
      Left            =   3315
      TabIndex        =   52
      Top             =   1815
      Width           =   225
   End
   Begin VB.Label Pisteet4 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   7
      Left            =   3315
      TabIndex        =   51
      Top             =   1650
      Width           =   225
   End
   Begin VB.Label Pisteet4 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   6
      Left            =   3315
      TabIndex        =   50
      Top             =   1470
      Width           =   225
   End
   Begin VB.Label Pisteet4 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   5
      Left            =   3315
      TabIndex        =   49
      Top             =   1305
      Width           =   225
   End
   Begin VB.Label Pisteet4 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   4
      Left            =   3315
      TabIndex        =   48
      Top             =   1140
      Width           =   225
   End
   Begin VB.Label Pisteet4 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   3
      Left            =   3315
      TabIndex        =   47
      Top             =   975
      Width           =   225
   End
   Begin VB.Label Pisteet4 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   2
      Left            =   3315
      TabIndex        =   46
      Top             =   810
      Width           =   225
   End
   Begin VB.Label Pisteet4 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   1
      Left            =   3315
      TabIndex        =   45
      Top             =   645
      Width           =   225
   End
   Begin VB.Label Pisteet4 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   165
      Index           =   0
      Left            =   3315
      TabIndex        =   44
      Top             =   480
      Width           =   225
   End
   Begin VB.Label Pisteet3 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   165
      Index           =   0
      Left            =   2325
      TabIndex        =   31
      Top             =   480
      Width           =   225
   End
   Begin VB.Label Pisteet2 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   12
      Left            =   1360
      TabIndex        =   30
      Top             =   2475
      Width           =   225
   End
   Begin VB.Label Pisteet2 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   11
      Left            =   1360
      TabIndex        =   29
      Top             =   2310
      Width           =   225
   End
   Begin VB.Label Pisteet2 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   10
      Left            =   1360
      TabIndex        =   28
      Top             =   2145
      Width           =   225
   End
   Begin VB.Label Pisteet2 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   9
      Left            =   1360
      TabIndex        =   27
      Top             =   1980
      Width           =   225
   End
   Begin VB.Label Pisteet2 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   8
      Left            =   1360
      TabIndex        =   26
      Top             =   1815
      Width           =   225
   End
   Begin VB.Label Pisteet2 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   7
      Left            =   1360
      TabIndex        =   25
      Top             =   1650
      Width           =   225
   End
   Begin VB.Label Pisteet2 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   6
      Left            =   1360
      TabIndex        =   24
      Top             =   1470
      Width           =   225
   End
   Begin VB.Label Pisteet2 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   5
      Left            =   1360
      TabIndex        =   23
      Top             =   1305
      Width           =   225
   End
   Begin VB.Label Pisteet2 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   4
      Left            =   1360
      TabIndex        =   22
      Top             =   1140
      Width           =   225
   End
   Begin VB.Label Pisteet2 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   3
      Left            =   1360
      TabIndex        =   21
      Top             =   975
      Width           =   225
   End
   Begin VB.Label Pisteet2 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   2
      Left            =   1360
      TabIndex        =   20
      Top             =   810
      Width           =   225
   End
   Begin VB.Label Pisteet2 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   1
      Left            =   1360
      TabIndex        =   19
      Top             =   645
      Width           =   225
   End
   Begin VB.Label Pisteet2 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   165
      Index           =   0
      Left            =   1360
      TabIndex        =   18
      Top             =   480
      Width           =   225
   End
   Begin VB.Label Pisteet1 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   12
      Left            =   360
      TabIndex        =   17
      Top             =   2475
      Width           =   225
   End
   Begin VB.Label Pisteet1 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   11
      Left            =   360
      TabIndex        =   16
      Top             =   2310
      Width           =   225
   End
   Begin VB.Label Pisteet1 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   10
      Left            =   360
      TabIndex        =   15
      Top             =   2145
      Width           =   225
   End
   Begin VB.Label Pisteet1 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   9
      Left            =   360
      TabIndex        =   14
      Top             =   1980
      Width           =   225
   End
   Begin VB.Label Pisteet1 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   8
      Left            =   360
      TabIndex        =   13
      Top             =   1815
      Width           =   225
   End
   Begin VB.Label Pisteet1 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   7
      Left            =   360
      TabIndex        =   12
      Top             =   1650
      Width           =   225
   End
   Begin VB.Label Pisteet1 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   6
      Left            =   360
      TabIndex        =   11
      Top             =   1470
      Width           =   225
   End
   Begin VB.Label Pisteet1 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   5
      Left            =   360
      TabIndex        =   10
      Top             =   1305
      Width           =   225
   End
   Begin VB.Label Pisteet1 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   4
      Left            =   360
      TabIndex        =   9
      Top             =   1140
      Width           =   225
   End
   Begin VB.Label Pisteet1 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   3
      Left            =   360
      TabIndex        =   8
      Top             =   975
      Width           =   225
   End
   Begin VB.Label Pisteet1 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   2
      Left            =   360
      TabIndex        =   7
      Top             =   810
      Width           =   225
   End
   Begin VB.Label Pisteet1 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   645
      Width           =   225
   End
   Begin VB.Label Pisteet1 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   165
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   480
      Width           =   225
   End
   Begin VB.Label lblPelaaja4 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   885
   End
   Begin VB.Label lblPelaaja3 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2025
      TabIndex        =   3
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   885
   End
   Begin VB.Image imgKuva 
      Height          =   480
      Left            =   4200
      Picture         =   "Pisteikkuna.frx":000C
      Top             =   1080
      Width           =   480
      Visible         =   0   'False
   End
   Begin VB.Label lblPelaaja2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1035
      TabIndex        =   2
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   885
   End
   Begin VB.Label lblPelaaja1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   885
   End
   Begin VB.Label Pisteet3 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   12
      Left            =   2325
      TabIndex        =   43
      Top             =   2475
      Width           =   225
   End
   Begin VB.Label Pisteet3 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   11
      Left            =   2325
      TabIndex        =   42
      Top             =   2310
      Width           =   225
   End
   Begin VB.Label Pisteet3 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   10
      Left            =   2325
      TabIndex        =   41
      Top             =   2145
      Width           =   225
   End
   Begin VB.Label Pisteet3 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   9
      Left            =   2325
      TabIndex        =   40
      Top             =   1980
      Width           =   225
   End
   Begin VB.Label Pisteet3 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   8
      Left            =   2325
      TabIndex        =   39
      Top             =   1815
      Width           =   225
   End
   Begin VB.Label Pisteet3 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   7
      Left            =   2325
      TabIndex        =   38
      Top             =   1650
      Width           =   225
   End
   Begin VB.Label Pisteet3 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   6
      Left            =   2325
      TabIndex        =   37
      Top             =   1470
      Width           =   225
   End
   Begin VB.Label Pisteet3 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   5
      Left            =   2325
      TabIndex        =   36
      Top             =   1305
      Width           =   225
   End
   Begin VB.Label Pisteet3 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   4
      Left            =   2325
      TabIndex        =   35
      Top             =   1140
      Width           =   225
   End
   Begin VB.Label Pisteet3 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   3
      Left            =   2325
      TabIndex        =   34
      Top             =   975
      Width           =   225
   End
   Begin VB.Label Pisteet3 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   2
      Left            =   2325
      TabIndex        =   33
      Top             =   810
      Width           =   225
   End
   Begin VB.Label Pisteet3 
      Alignment       =   1  'Right Justify
      Height          =   165
      Index           =   1
      Left            =   2325
      TabIndex        =   32
      Top             =   645
      Width           =   225
   End
End
Attribute VB_Name = "Pisteikkuna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()

If Kurkkuikkuna.Pelikaynnissa <> 1 Then
    Kurkkuikkuna.Nollaus
    Kurkkuikkuna.UusiJako
    Kurkkuikkuna.Enabled = True
    Kurkkuikkuna.Monesko = Kurkkuikkuna.Monesko + 1
    Pisteikkuna.Visible = False
    If Aloitusvuoro <> 1 Then
        Kurkkuikkuna.Aloitettu = 0
        If Kurkkuikkuna.Aloitusvuoro <> 1 Then
            Kurkkuikkuna.pcP1k_Click (7)
        End If
    End If
Else
    Kurkkuikkuna.Enabled = True
    Pisteikkuna.Visible = False
End If

End Sub


Private Sub Form_Load()


Kurkkuikkuna.Enabled = False

lblPelaaja1.Caption = Kurkkuikkuna.lblPelaaja1.Caption
lblPelaaja2.Caption = Kurkkuikkuna.lblPelaaja2.Caption
lblPelaaja3.Caption = Kurkkuikkuna.lblPelaaja3.Caption
lblPelaaja4.Caption = Kurkkuikkuna.lblPelaaja4.Caption

Pisteikkuna.Caption = "Pistetilanne"

End Sub



