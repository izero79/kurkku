VERSION 5.00
Begin VB.Form Asetuksetikkuna 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Kurkun asetukset"
   ClientHeight    =   3675
   ClientLeft      =   3420
   ClientTop       =   3315
   ClientWidth     =   3825
   ClipControls    =   0   'False
   Icon            =   "Asetuksetikkuna.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3675
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmTietokoneet 
      Caption         =   "&Tietokoneen pelaajien nimet"
      Height          =   1695
      Left            =   100
      TabIndex        =   7
      Top             =   1680
      Width           =   2250
      Begin VB.TextBox txtPelaaja3 
         Height          =   360
         Left            =   120
         MaxLength       =   14
         TabIndex        =   2
         Text            =   "Keijo"
         Top             =   780
         Width           =   1120
      End
      Begin VB.TextBox txtPelaaja4 
         Height          =   360
         Left            =   120
         MaxLength       =   14
         TabIndex        =   3
         Text            =   "Juuso"
         Top             =   1200
         Width           =   1120
      End
      Begin VB.TextBox txtPelaaja2 
         Height          =   360
         Left            =   120
         MaxLength       =   14
         TabIndex        =   1
         Text            =   "Anne"
         Top             =   360
         Width           =   1120
      End
   End
   Begin VB.Frame frmAnimaatio 
      Caption         =   "Pelin asetukset"
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   105
      Width           =   2250
      Begin VB.TextBox txtPisteraja 
         Height          =   360
         Left            =   120
         MaxLength       =   3
         TabIndex        =   0
         Top             =   720
         Width           =   1120
      End
      Begin VB.Label Label1 
         Caption         =   "Pistem‰‰r‰ johon saakka pelataan. (30-99)"
         Height          =   405
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdPeruuta 
      Caption         =   "Peruuta"
      Height          =   375
      Left            =   2550
      TabIndex        =   5
      Top             =   640
      Width           =   1130
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2550
      TabIndex        =   4
      Top             =   180
      Width           =   1130
   End
End
Attribute VB_Name = "Asetuksetikkuna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOK_Click()

On Error GoTo Virhe
pistetext = Int(txtPisteraja.Text)

If pistetext < 30 Or pistetext > 99 Then
    Viesti = MsgBox("Pisteraja oltava v‰lill‰ 30-99", 48, "Virhe asetuksissa")
    Exit Sub
End If

If pistetext <= Kurkkuikkuna.Kokonaispisteet1 Or pistetext <= Kurkkuikkuna.Kokonaispisteet2 Or pistetext <= Kurkkuikkuna.Kokonaispisteet3 Or pistetext <= Kurkkuikkuna.Kokonaispisteet4 Then
    Viesti = MsgBox("Et voi vaihtaa rajaa pienemm‰ksi tai yht‰suureksi, kuin pelaajilla on pisteit‰", 48, "Virhe asetuksissa")
    Exit Sub
End If

Kurkkuikkuna.Enabled = True
Kurkkuikkuna.Pisteraja = Int(txtPisteraja.Text)
Kurkkuikkuna.lblPelaaja2.Caption = txtPelaaja2.Text
Kurkkuikkuna.lblPelaaja3.Caption = txtPelaaja3.Text
Kurkkuikkuna.lblPelaaja4.Caption = txtPelaaja4.Text
Asetuksetikkuna.Visible = False
Unload Asetuksetikkuna

Exit Sub

Virhe:

Viesti = MsgBox("Pisteraja oltava v‰lill‰ 30-99", 48, "Virhe asetuksissa")
Exit Sub

End Sub

Private Sub cmdPeruuta_Click()

Kurkkuikkuna.Enabled = True
Unload Asetuksetikkuna

End Sub


Private Sub Form_Load()

Kurkkuikkuna.Enabled = False
txtPisteraja.Text = Kurkkuikkuna.Pisteraja
txtPelaaja2.Text = Kurkkuikkuna.lblPelaaja2.Caption
txtPelaaja3.Text = Kurkkuikkuna.lblPelaaja3.Caption
txtPelaaja4.Text = Kurkkuikkuna.lblPelaaja4.Caption

End Sub


Private Sub Form_Unload(Cancel As Integer)

Kurkkuikkuna.Enabled = True

End Sub


Private Sub txtPelaaja2_GotFocus()

txtPelaaja2.SelStart = 0
txtPelaaja2.SelLength = 99

End Sub


Private Sub txtPelaaja3_GotFocus()

txtPelaaja3.SelStart = 0
txtPelaaja3.SelLength = 99

End Sub


Private Sub txtPelaaja4_GotFocus()

txtPelaaja4.SelStart = 0
txtPelaaja4.SelLength = 99

End Sub


Private Sub txtPisteraja_GotFocus()

txtPisteraja.SelStart = 0
txtPisteraja.SelLength = 99

End Sub


