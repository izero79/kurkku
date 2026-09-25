VERSION 5.00
Begin VB.Form Avausikkuna 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TESI software Kurkku"
   ClientHeight    =   1515
   ClientLeft      =   3090
   ClientTop       =   4170
   ClientWidth     =   4680
   ClipControls    =   0   'False
   Icon            =   "Avausikkuna.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1515
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLopeta 
      Cancel          =   -1  'True
      Caption         =   "Lopeta"
      Height          =   345
      Left            =   3480
      TabIndex        =   4
      Top             =   570
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3480
      TabIndex        =   3
      Top             =   150
      Width           =   1110
   End
   Begin VB.TextBox txtNimi 
      Height          =   300
      Left            =   1830
      MaxLength       =   14
      TabIndex        =   2
      Top             =   615
      Width           =   1440
   End
   Begin VB.Label lblNimi 
      Caption         =   "Mikä on nimesi?"
      Height          =   255
      Left            =   50
      TabIndex        =   1
      Top             =   645
      Width           =   1335
   End
   Begin VB.Label lblTervetuloa 
      Caption         =   "Tervetuloa pelaamaan TESI software Kurkkua."
      Height          =   255
      Left            =   45
      TabIndex        =   0
      Top             =   285
      Width           =   3405
   End
End
Attribute VB_Name = "Avausikkuna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private alku As Integer
Public Sub Aloita_peli()

Kurkkuikkuna.lblPelaaja1.Caption = Avausikkuna.txtNimi.Text
Kurkkuikkuna.Enabled = True
Kurkkuikkuna.Aani_Sekoitus
Kurkkuikkuna.Soita_Aani

If Kurkkuikkuna.mnuAani.Checked = True Then
    Do Until Kurkkuikkuna.MMControl1.Mode = 525
    Loop
End If

Kurkkuikkuna.Kortit_Alussa
Kurkkuikkuna.Jaa_Kortit
Kurkkuikkuna.Jaa_Kortit2
Kurkkuikkuna.Jaa_Kortit3
Kurkkuikkuna.Jaa_Kortit4
Vuoro = 1
Kurkkuikkuna.sbarTilapalkki.SimpleText = "Valitse pelattava kortti."
Avausikkuna.Visible = False
Unload Avausikkuna

End Sub


Private Sub cmdLopeta_Click()

alku = 1
End

End Sub

Private Sub cmdOK_Click()

alku = 1

If txtNimi.Text <> "" Then
    Pelaajannimi = txtNimi.Text
    Aloita_peli
End If

End Sub

Private Sub Form_Load()

alku = 0
txtNimi.Text = A
txtNimi.SelStart = 0
txtNimi.SelLength = 99
Kurkkuikkuna.Enabled = False
Kurkkuikkuna.Visible = False
Kurkkuikkuna.Visible = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

If alku <> 1 Then
    End
End If

End Sub


Private Sub txtNimi_GotFocus()

txtNimi.SelStart = 0
txtNimi.SelLength = 99

End Sub


