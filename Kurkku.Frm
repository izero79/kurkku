VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.Form Kurkkuikkuna 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TESI software Kurkku"
   ClientHeight    =   6150
   ClientLeft      =   1830
   ClientTop       =   3165
   ClientWidth     =   8010
   Icon            =   "Kurkku.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6150
   ScaleWidth      =   8010
   Begin VB.Timer tmrAika 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5760
      Top             =   5040
   End
   Begin VB.PictureBox pcPelattuP4 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   4080
      ScaleHeight     =   1440
      ScaleWidth      =   1080
      TabIndex        =   7
      Top             =   2160
      Width           =   1080
   End
   Begin VB.PictureBox pcPelattuP3 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   3480
      ScaleHeight     =   1440
      ScaleWidth      =   1080
      TabIndex        =   6
      Top             =   1680
      Width           =   1080
   End
   Begin VB.PictureBox pcPelattuP2 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   2880
      ScaleHeight     =   1440
      ScaleWidth      =   1080
      TabIndex        =   5
      Top             =   2160
      Width           =   1080
   End
   Begin VB.PictureBox pcPelattuP1 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   3480
      ScaleHeight     =   1440
      ScaleWidth      =   1080
      TabIndex        =   4
      Top             =   2640
      Width           =   1080
   End
   Begin ComctlLib.StatusBar sbarTilapalkki 
      Height          =   345
      Left            =   60
      TabIndex        =   0
      Top             =   5760
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   609
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MCI.MMControl MMControl1 
      Height          =   330
      Left            =   2280
      TabIndex        =   8
      Top             =   7000
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   582
      _Version        =   327680
      DeviceType      =   ""
      FileName        =   ""
   End
   Begin PicClip.PictureClip pcKortinkuvat 
      Left            =   7440
      Top             =   5880
      _ExtentX        =   26300
      _ExtentY        =   10160
      _Version        =   327680
      Rows            =   4
      Cols            =   14
      Picture         =   "Kurkku.frx":0E42
   End
   Begin VB.Image pcP2K 
      Height          =   1440
      Index           =   5
      Left            =   240
      Top             =   2760
      Width           =   1080
   End
   Begin VB.Image pcP2K 
      Height          =   1440
      Index           =   4
      Left            =   240
      Top             =   2520
      Width           =   1080
   End
   Begin VB.Image pcP2K 
      Height          =   1440
      Index           =   3
      Left            =   240
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Image pcP2K 
      Height          =   1440
      Index           =   2
      Left            =   240
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Image pcP2K 
      Height          =   1440
      Index           =   1
      Left            =   240
      Top             =   1800
      Width           =   1080
   End
   Begin VB.Image pcP2K 
      Height          =   1440
      Index           =   0
      Left            =   240
      Top             =   1560
      Width           =   1080
   End
   Begin VB.Image pcP1K 
      Height          =   1440
      Index           =   5
      Left            =   4080
      Top             =   4200
      Width           =   1080
   End
   Begin VB.Image pcP1K 
      Height          =   1440
      Index           =   4
      Left            =   3840
      Top             =   4200
      Width           =   1080
   End
   Begin VB.Image pcP1K 
      Height          =   1440
      Index           =   3
      Left            =   3600
      Top             =   4200
      Width           =   1080
   End
   Begin VB.Image pcP1K 
      Height          =   1440
      Index           =   2
      Left            =   3360
      Top             =   4200
      Width           =   1080
   End
   Begin VB.Image pcP1K 
      Height          =   1440
      Index           =   1
      Left            =   3120
      Top             =   4200
      Width           =   1080
   End
   Begin VB.Image pcP1K 
      Height          =   1440
      Index           =   0
      Left            =   2880
      Top             =   4200
      Width           =   1080
   End
   Begin VB.Label lblPelaaja2 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   960
      UseMnemonic     =   0   'False
      Width           =   2340
   End
   Begin VB.Label lblPelaaja4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5520
      TabIndex        =   3
      Top             =   5160
      UseMnemonic     =   0   'False
      Width           =   2355
   End
   Begin VB.Label lblPelaaja3 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5640
      TabIndex        =   2
      Top             =   30
      UseMnemonic     =   0   'False
      Width           =   2295
   End
   Begin VB.Label lblPelaaja1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   5280
      UseMnemonic     =   0   'False
      Width           =   2295
   End
   Begin VB.Shape shpPalkki 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      Height          =   405
      Left            =   -15
      Top             =   5740
      Width           =   8500
   End
   Begin VB.Image pcP3K 
      Height          =   1440
      Index           =   0
      Left            =   2880
      Top             =   60
      Width           =   1080
   End
   Begin VB.Image pcP3K 
      Height          =   1440
      Index           =   1
      Left            =   3120
      Top             =   60
      Width           =   1080
   End
   Begin VB.Image pcP3K 
      Height          =   1440
      Index           =   2
      Left            =   3360
      Top             =   60
      Width           =   1080
   End
   Begin VB.Image pcP3K 
      Height          =   1440
      Index           =   3
      Left            =   3600
      Top             =   60
      Width           =   1080
   End
   Begin VB.Image pcP3K 
      Height          =   1440
      Index           =   4
      Left            =   3840
      Top             =   60
      Width           =   1080
   End
   Begin VB.Image pcP3K 
      Height          =   1440
      Index           =   5
      Left            =   4080
      Top             =   60
      Width           =   1080
   End
   Begin VB.Image pcP4k 
      Height          =   1440
      Index           =   0
      Left            =   6720
      Top             =   1560
      Width           =   1080
   End
   Begin VB.Image pcP4k 
      Height          =   1440
      Index           =   1
      Left            =   6720
      Top             =   1800
      Width           =   1080
   End
   Begin VB.Image pcP4k 
      Height          =   1440
      Index           =   2
      Left            =   6720
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Image pcP4k 
      Height          =   1440
      Index           =   3
      Left            =   6720
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Image pcP4k 
      Height          =   1440
      Index           =   4
      Left            =   6720
      Top             =   2520
      Width           =   1080
   End
   Begin VB.Image pcP4k 
      Height          =   1440
      Index           =   5
      Left            =   6720
      Top             =   2760
      Width           =   1080
   End
   Begin VB.Menu mnuPeli 
      Caption         =   "&Peli"
      Begin VB.Menu mnuUusipeli 
         Caption         =   "&Uusi peli"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnu1Valiviiva 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAsetukset 
         Caption         =   "&Asetukset..."
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuAani 
         Caption         =   "&ƒ‰ni"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuPistetilanne 
         Caption         =   "P&istetilanne..."
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnu2Valiviiva 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLopeta 
         Caption         =   "&Lopeta"
      End
   End
   Begin VB.Menu mnuOhje 
      Caption         =   "&Ohje"
      Begin VB.Menu mnuOhjeenAiheet 
         Caption         =   "&Ohjeen aiheet"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnu3Valiviiva 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInternet 
         Caption         =   "T&ESI software Internetiss‰"
      End
      Begin VB.Menu mnusahkoposti 
         Caption         =   "&L‰het‰ s‰hkˆpostia"
      End
      Begin VB.Menu mnu4ValiViiva 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTietojaKurkusta 
         Caption         =   "&Tietoja Kurkusta"
      End
   End
End
Attribute VB_Name = "Kurkkuikkuna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

#If Win32 Then
Private Declare Function ShellExecute Lib _
    "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

#Else
Private Declare Function ShellExecute Lib _
    "shell.dll" (ByVal _
    hWnd As Integer, ByVal lpOperation _
    As String, ByVal lpFile _
    As String, ByVal lpParameters _
    As String, ByVal lpDirectory _
    As String, ByVal nShowCmd _
    As Integer) As Integer
#End If


Public A, Pelaajannimi As String
Public Kortteja, Pelikaynnissa, Vali, Vali2, Vuoro, Korttipelattu, Pelinalku As Integer
Public P1Kortit As Variant
Public P2Kortit As Variant
Public P3Kortit As Variant
Public P4Kortit As Variant
Public PelattuP1, PelattuP2, PelattuP3, PelattuP4, Johtokortti As Integer
Public Pelattu, Kierros, Haviaja1, Haviaja2, Haviaja3, Haviaja4 As Integer
Public johKortti, tyhKortti, tKortti6, tKortti5, tKortti4, tKortti3, tKortti2, tKortti1, Suurin, Pienin, Seuraava, Pisteet1, Pisteet2, Pisteet3, Pisteet4 As Integer
Public i, Kortti1, Kortti2, Kortti3, Kortti4 As Integer
Public Kokonaispisteet1, Kokonaispisteet2, Kokonaispisteet3, Kokonaispisteet4 As Integer
Public Havio, PVoitto, Pisteraja, Aloittava, H, P, R, Ru, Tarjottava As Integer
Public Uusiteksti1, Uusiteksti2, Uusiteksti3, Uusiteksti4 As String
Public Aloitettu, Aloitusvuoro, Monesko As Integer
Public moodi As Integer
Public pistetext As Integer
Public Aika As Integer
Public Ajanotto As Integer
Public OdotusArvo As Integer
Public Sijoitus As Integer

Public Sub Aani_Havio()

If mnuAani.Checked = True Then
    MMControl1.Command = "Close"
    MMControl1.DeviceType = "WaveAudio"
    MMControl1.filename = App.Path + "\Havio.wav"
    MMControl1.TimeFormat = 0
    MMControl1.Command = "Open"
End If

End Sub

Public Sub Aani_Sekoitus()

If mnuAani.Checked = True Then
    MMControl1.Command = "Close"
    MMControl1.DeviceType = "WaveAudio"
    MMControl1.filename = App.Path + "\Sekoitus.wav"
    MMControl1.TimeFormat = 0
    MMControl1.Command = "Open"
End If

End Sub
Public Sub Aani_Voitto()

If mnuAani.Checked = True Then
    MMControl1.Command = "Close"
    MMControl1.DeviceType = "WaveAudio"
    MMControl1.filename = App.Path + "\Voitto.wav"
    MMControl1.TimeFormat = 0
    MMControl1.Command = "Open"
End If

End Sub
Public Sub Jaa_Kortit()

P1Kortit = Array(0, 1, 2, 3, 4, 5, 6)

arvonta:
For i = 1 To 6
    Randomize
    P1Kortit(i) = Int(Rnd * 51) + 1
Next
    
If P1Kortit(2) = P1Kortit(1) Then GoTo arvonta
If P1Kortit(3) = P1Kortit(1) Or P1Kortit(3) = P1Kortit(2) Then GoTo arvonta
If P1Kortit(4) = P1Kortit(1) Or P1Kortit(4) = P1Kortit(2) Or P1Kortit(4) = P1Kortit(3) Then GoTo arvonta
If P1Kortit(5) = P1Kortit(1) Or P1Kortit(5) = P1Kortit(2) Or P1Kortit(5) = P1Kortit(3) Or P1Kortit(5) = P1Kortit(4) Then GoTo arvonta
If P1Kortit(6) = P1Kortit(1) Or P1Kortit(6) = P1Kortit(2) Or P1Kortit(6) = P1Kortit(3) Or P1Kortit(6) = P1Kortit(4) Or P1Kortit(6) = P1Kortit(5) Then GoTo arvonta

Jarjesta_1

For i = 0 To 5
    pcP1K(i).Picture = pcKortinkuvat.GraphicCell(P1Kortit(i + 1))
Next

End Sub

Public Sub Jaa_Kortit2()

P2Kortit = Array(0, 1, 2, 3, 4, 5, 6)

arvonta:
For i = 1 To 6
arpa:
    Randomize
    P2Kortit(i) = Int(Rnd * 51) + 1
    For j = 1 To 6
        If P2Kortit(i) = P1Kortit(j) Then GoTo arpa
    Next
Next
    
If P2Kortit(2) = P2Kortit(1) Then GoTo arvonta
If P2Kortit(3) = P2Kortit(1) Or P2Kortit(3) = P2Kortit(2) Then GoTo arvonta
If P2Kortit(4) = P2Kortit(1) Or P2Kortit(4) = P2Kortit(2) Or P2Kortit(4) = P2Kortit(3) Then GoTo arvonta
If P2Kortit(5) = P2Kortit(1) Or P2Kortit(5) = P2Kortit(2) Or P2Kortit(5) = P2Kortit(3) Or P2Kortit(5) = P2Kortit(4) Then GoTo arvonta
If P2Kortit(6) = P2Kortit(1) Or P2Kortit(6) = P2Kortit(2) Or P2Kortit(6) = P2Kortit(3) Or P2Kortit(6) = P2Kortit(4) Or P2Kortit(6) = P2Kortit(5) Then GoTo arvonta

For j = 1 To 100
    For i = 1 To 5
        If P2Kortit(i) >= P2Kortit(i + 1) Then
            Vali = P2Kortit(i)
            P2Kortit(i) = P2Kortit(i + 1)
            P2Kortit(i + 1) = Vali
        End If
    Next
Next

End Sub

Public Sub Jaa_Kortit3()

P3Kortit = Array(0, 1, 2, 3, 4, 5, 6)

arvonta:
For i = 1 To 6
arpa:
    Randomize
    P3Kortit(i) = Int(Rnd * 51) + 1
    For j = 1 To 6
        If P3Kortit(i) = P1Kortit(j) Then GoTo arpa
    Next
    For K = 1 To 6
        If P3Kortit(i) = P2Kortit(K) Then GoTo arpa
    Next
Next
    
If P3Kortit(2) = P3Kortit(1) Then GoTo arvonta
If P3Kortit(3) = P3Kortit(1) Or P3Kortit(3) = P3Kortit(2) Then GoTo arvonta
If P3Kortit(4) = P3Kortit(1) Or P3Kortit(4) = P3Kortit(2) Or P3Kortit(4) = P3Kortit(3) Then GoTo arvonta
If P3Kortit(5) = P3Kortit(1) Or P3Kortit(5) = P3Kortit(2) Or P3Kortit(5) = P3Kortit(3) Or P3Kortit(5) = P3Kortit(4) Then GoTo arvonta
If P3Kortit(6) = P3Kortit(1) Or P3Kortit(6) = P3Kortit(2) Or P3Kortit(6) = P3Kortit(3) Or P3Kortit(6) = P3Kortit(4) Or P3Kortit(6) = P3Kortit(5) Then GoTo arvonta

For j = 1 To 100
    For i = 1 To 5
        If P3Kortit(i) >= P3Kortit(i + 1) Then
            Vali = P3Kortit(i)
            P3Kortit(i) = P3Kortit(i + 1)
            P3Kortit(i + 1) = Vali
        End If
    Next
Next

End Sub


Public Sub Jaa_Kortit4()

P4Kortit = Array(0, 1, 2, 3, 4, 5, 6)

arvonta:
For i = 1 To 6
arpa:
    Randomize
    P4Kortit(i) = Int(Rnd * 51) + 1
    For j = 1 To 6
        If P4Kortit(i) = P1Kortit(j) Then GoTo arpa
    Next
    For K = 1 To 6
        If P4Kortit(i) = P2Kortit(K) Then GoTo arpa
    Next
    For L = 1 To 6
        If P4Kortit(i) = P3Kortit(L) Then GoTo arpa
    Next
Next
    
If P4Kortit(2) = P4Kortit(1) Then GoTo arvonta
If P4Kortit(3) = P4Kortit(1) Or P4Kortit(3) = P4Kortit(2) Then GoTo arvonta
If P4Kortit(4) = P4Kortit(1) Or P4Kortit(4) = P4Kortit(2) Or P4Kortit(4) = P4Kortit(3) Then GoTo arvonta
If P4Kortit(5) = P4Kortit(1) Or P4Kortit(5) = P4Kortit(2) Or P4Kortit(5) = P4Kortit(3) Or P4Kortit(5) = P4Kortit(4) Then GoTo arvonta
If P4Kortit(6) = P4Kortit(1) Or P4Kortit(6) = P4Kortit(2) Or P4Kortit(6) = P4Kortit(3) Or P4Kortit(6) = P4Kortit(4) Or P4Kortit(6) = P4Kortit(5) Then GoTo arvonta

For j = 1 To 100
    For i = 1 To 5
        If P4Kortit(i) >= P4Kortit(i + 1) Then
            Vali = P4Kortit(i)
            P4Kortit(i) = P4Kortit(i + 1)
            P4Kortit(i + 1) = Vali
        End If
    Next
Next

End Sub


Public Sub Jarjesta_1()

For j = 1 To 100
    For i = 1 To 5
        If P1Kortit(i) >= P1Kortit(i + 1) Then
            Vali = P1Kortit(i)
            P1Kortit(i) = P1Kortit(i + 1)
            P1Kortit(i + 1) = Vali
        End If
    Next
Next

End Sub

Private Sub jatkaa3()

Aloittava = 0
Pelaaja2
For j = 1 To 30000
Next
Tarkista_KierrosVoitto
Odotus

End Sub

Public Sub Kortit_Alussa()

For i = 0 To 5
    pcP1K(i).Picture = pcKortinkuvat.GraphicCell(0)
    pcP2K(i).Picture = pcKortinkuvat.GraphicCell(0)
    pcP3K(i).Picture = pcKortinkuvat.GraphicCell(0)
    pcP4k(i).Picture = pcKortinkuvat.GraphicCell(0)
Next

End Sub


Private Sub Lataa_avausikkuna()

Load Avausikkuna
Avausikkuna.Top = Kurkkuikkuna.Top + 1300
Avausikkuna.Left = Kurkkuikkuna.Left + 300
Avausikkuna.Visible = True
Avausikkuna.ZOrder 0

End Sub

Public Sub Lisaa_Pisteet()

If Haviaja1 = 1 Then
    Kokonaispisteet1 = Kokonaispisteet1 + Kortti1
End If

If Haviaja2 = 1 Then
    Kokonaispisteet2 = Kokonaispisteet2 + Kortti2
End If

If Haviaja3 = 1 Then
    Kokonaispisteet3 = Kokonaispisteet3 + Kortti3
End If

If Haviaja4 = 1 Then
    Kokonaispisteet4 = Kokonaispisteet4 + Kortti4
End If

End Sub
Public Sub LisaaPist()
tmrAika.Enabled = False
OdotusArvo = 0
pcPelattuP1.Picture = LoadPicture()
pcPelattuP2.Picture = LoadPicture()
pcPelattuP3.Picture = LoadPicture()
pcPelattuP4.Picture = LoadPicture()
PelattuP1 = 0
PelattuP2 = 0
PelattuP3 = 0
PelattuP4 = 0
Johtokortti = 0

Lisaa_Pisteet

Pelikaynnissa = 0
Load Pisteikkuna
Pisteikkuna.Top = Kurkkuikkuna.Top + 2100
Pisteikkuna.Left = Kurkkuikkuna.Left + 1250
If Monesko < 11 Then
    Pisteikkuna.Pisteet1(Monesko).Caption = Str(Kokonaispisteet1)
    Pisteikkuna.Pisteet2(Monesko).Caption = Str(Kokonaispisteet2)
    Pisteikkuna.Pisteet3(Monesko).Caption = Str(Kokonaispisteet3)
    Pisteikkuna.Pisteet4(Monesko).Caption = Str(Kokonaispisteet4)
Else
    Pisteikkuna.Pisteet1(12).Caption = Str(Kokonaispisteet1)
    Pisteikkuna.Pisteet2(12).Caption = Str(Kokonaispisteet2)
    Pisteikkuna.Pisteet3(12).Caption = Str(Kokonaispisteet3)
    Pisteikkuna.Pisteet4(12).Caption = Str(Kokonaispisteet4)

End If

If Kokonaispisteet1 >= Pisteraja Then
    Havio = 1
    Pisteikkuna.Caption = "Pelin loppu"
    sbarTilapalkki.SimpleText = "Peli loppunut, Aloita uusi peli painamalla F2."
    Pisteikkuna.lblPelaaja1.ForeColor = &HFF&
    For i = 0 To 12
    Pisteikkuna.Pisteet1(i).ForeColor = &HFF&
    Next
    Pisteikkuna.imgKuva.Visible = True
    Pelikaynnissa = 1
End If
If Kokonaispisteet2 >= Pisteraja Then
    Pisteikkuna.Caption = "Pelin loppu"
    sbarTilapalkki.SimpleText = "Peli loppunut, Aloita uusi peli painamalla F2."
    Pisteikkuna.lblPelaaja2.ForeColor = &HFF&
    For i = 0 To 12
    Pisteikkuna.Pisteet2(i).ForeColor = &HFF&
    Next
    Pisteikkuna.imgKuva.Visible = True
    Pelikaynnissa = 1
End If
If Kokonaispisteet3 >= Pisteraja Then
    Pisteikkuna.Caption = "Pelin loppu"
    sbarTilapalkki.SimpleText = "Peli loppunut, Aloita uusi peli painamalla F2."
    Pisteikkuna.lblPelaaja3.ForeColor = &HFF&
    For i = 0 To 12
    Pisteikkuna.Pisteet3(i).ForeColor = &HFF&
    Next
    Pisteikkuna.imgKuva.Visible = True
    Pelikaynnissa = 1
End If
If Kokonaispisteet4 >= Pisteraja Then
    Pisteikkuna.Caption = "Pelin loppu"
    sbarTilapalkki.SimpleText = "Peli loppunut, Aloita uusi peli painamalla F2."
    Pisteikkuna.lblPelaaja4.ForeColor = &HFF&
    For i = 0 To 12
    Pisteikkuna.Pisteet4(i).ForeColor = &HFF&
    Next
    Pisteikkuna.imgKuva.Visible = True
    Pelikaynnissa = 1
End If

If Kurkkuikkuna.Kokonaispisteet2 >= Pisteraja Or Kurkkuikkuna.Kokonaispisteet3 >= Pisteraja Or Kurkkuikkuna.Kokonaispisteet4 >= Pisteraja Then
PVoitto = 1
End If

Pisteikkuna.Visible = True
Pisteikkuna.ZOrder 0

Pisteikkuna.Caption = "Pistetilanne"
TarkistaPisteet
If Sijoitus = 4 And Kurkkuikkuna.Havio = 1 Then
Pisteikkuna.Caption = "Pistetilanne - h‰visit"
Kurkkuikkuna.Aani_Havio
Kurkkuikkuna.Soita_Aani
End If

If Sijoitus = 1 And Kurkkuikkuna.PVoitto = 1 Then
Pisteikkuna.Caption = "Pistetilanne - voitit"
Kurkkuikkuna.Aani_Voitto
Kurkkuikkuna.Soita_Aani
End If

Kurkkuikkuna.Enabled = False

End Sub

Public Sub Muuta_pisteet()

Select Case Pisteet1

Case 1, 2, 3, 4
    Kortti1 = 2
Case 5, 6, 7, 8
    Kortti1 = 3
Case 9, 10, 11, 12
    Kortti1 = 4
Case 13, 14, 15, 16
    Kortti1 = 5
Case 17, 18, 19, 20
    Kortti1 = 6
Case 21, 22, 23, 24
    Kortti1 = 7
Case 25, 26, 27, 28
    Kortti1 = 8
Case 29, 30, 31, 32
    Kortti1 = 9
Case 33, 34, 35, 36
    Kortti1 = 10
Case 37, 38, 39, 40
    Kortti1 = 11
Case 41, 42, 43, 44
    Kortti1 = 12
Case 45, 46, 47, 48
    Kortti1 = 13
Case 49, 50, 51, 52
    Kortti1 = 14
    
End Select

Select Case Pisteet2

Case 1, 2, 3, 4
    Kortti2 = 2
Case 5, 6, 7, 8
    Kortti2 = 3
Case 9, 10, 11, 12
    Kortti2 = 4
Case 13, 14, 15, 16
    Kortti2 = 5
Case 17, 18, 19, 20
    Kortti2 = 6
Case 21, 22, 23, 24
    Kortti2 = 7
Case 25, 26, 27, 28
    Kortti2 = 8
Case 29, 30, 31, 32
    Kortti2 = 9
Case 33, 34, 35, 36
    Kortti2 = 10
Case 37, 38, 39, 40
    Kortti2 = 11
Case 41, 42, 43, 44
    Kortti2 = 12
Case 45, 46, 47, 48
    Kortti2 = 13
Case 49, 50, 51, 52
    Kortti2 = 14
    
End Select

Select Case Pisteet3

Case 1, 2, 3, 4
    Kortti3 = 2
Case 5, 6, 7, 8
    Kortti3 = 3
Case 9, 10, 11, 12
    Kortti3 = 4
Case 13, 14, 15, 16
    Kortti3 = 5
Case 17, 18, 19, 20
    Kortti3 = 6
Case 21, 22, 23, 24
    Kortti3 = 7
Case 25, 26, 27, 28
    Kortti3 = 8
Case 29, 30, 31, 32
    Kortti3 = 9
Case 33, 34, 35, 36
    Kortti3 = 10
Case 37, 38, 39, 40
    Kortti3 = 11
Case 41, 42, 43, 44
    Kortti3 = 12
Case 45, 46, 47, 48
    Kortti3 = 13
Case 49, 50, 51, 52
    Kortti3 = 14
    
End Select

Select Case Pisteet4

Case 1, 2, 3, 4
    Kortti4 = 2
Case 5, 6, 7, 8
    Kortti4 = 3
Case 9, 10, 11, 12
    Kortti4 = 4
Case 13, 14, 15, 16
    Kortti4 = 5
Case 17, 18, 19, 20
    Kortti4 = 6
Case 21, 22, 23, 24
    Kortti4 = 7
Case 25, 26, 27, 28
    Kortti4 = 8
Case 29, 30, 31, 32
    Kortti4 = 9
Case 33, 34, 35, 36
    Kortti4 = 10
Case 37, 38, 39, 40
    Kortti4 = 11
Case 41, 42, 43, 44
    Kortti4 = 12
Case 45, 46, 47, 48
    Kortti4 = 13
Case 49, 50, 51, 52
    Kortti4 = 14
    
End Select

End Sub

Public Sub Nollaus()

Aloittava = 0
Vuoro = 1
Vali = 0
Vali2 = 0
Vuoro = 0
Korttipelattu = 1
Pelinalku = 0

For i = 0 To 6
    P1Kortit(i) = 0
    P2Kortit(i) = 0
    P3Kortit(i) = 0
    P4Kortit(i) = 0
Next

PelattuP1 = 0
PelattuP2 = 0
PelattuP3 = 0
PelattuP4 = 0
Johtokortti = 0
Pelattu = 0
Kierros = 1
Haviaja1 = 0
Haviaja2 = 0
Haviaja3 = 0
Haviaja4 = 0
Suurin = 0
Pienin = 0
Seuraava = 0
Pisteet1 = 0
Pisteet2 = 0
Pisteet3 = 0
Pisteet4 = 0
i = 0
Kortti1 = 0
Kortti2 = 0
Kortti3 = 0
Kortti4 = 0
Aloittava = 0


End Sub

Public Sub Nollaus_alku()

pcPelattuP1.Picture = LoadPicture()
pcPelattuP2.Picture = LoadPicture()
pcPelattuP3.Picture = LoadPicture()
pcPelattuP4.Picture = LoadPicture()
Pisteikkuna.Caption = "Pistetilanne"
Pisteikkuna.lblPelaaja1.ForeColor = &H80000012
Pisteikkuna.lblPelaaja2.ForeColor = &H80000012
Pisteikkuna.lblPelaaja3.ForeColor = &H80000012
Pisteikkuna.lblPelaaja4.ForeColor = &H80000012
Pisteikkuna.Pisteet1(0) = "0"
Pisteikkuna.Pisteet2(0) = "0"
Pisteikkuna.Pisteet3(0) = "0"
Pisteikkuna.Pisteet4(0) = "0"
Pisteikkuna.Pisteet1(0).ForeColor = &H80000012
Pisteikkuna.Pisteet2(0).ForeColor = &H80000012
Pisteikkuna.Pisteet3(0).ForeColor = &H80000012
Pisteikkuna.Pisteet4(0).ForeColor = &H80000012

For i = 1 To 12
    Pisteikkuna.Pisteet1(i) = ""
    Pisteikkuna.Pisteet2(i) = ""
    Pisteikkuna.Pisteet3(i) = ""
    Pisteikkuna.Pisteet4(i) = ""
    Pisteikkuna.Pisteet1(i).ForeColor = &H80000012
    Pisteikkuna.Pisteet2(i).ForeColor = &H80000012
    Pisteikkuna.Pisteet3(i).ForeColor = &H80000012
    Pisteikkuna.Pisteet4(i).ForeColor = &H80000012
Next

Pisteikkuna.Visible = False
Havio = 0
PVoitto = 0
Monesko = 0
Kokonaispisteet1 = 0
Kokonaispisteet2 = 0
Kokonaispisteet3 = 0
Kokonaispisteet4 = 0
Kurkkuikkuna.Enabled = True

End Sub

Public Sub Odotus()

tmrAika.Enabled = True
Ajanotto = 0
Aika = 2

End Sub

Private Sub P2()

Aloittava = 1
Pelaaja2
For j = 1 To 30000
Next
Aloittava = 0
Pelaaja3
For j = 1 To 30000
Next
Aloittava = 0
Pelaaja4
For j = 1 To 30000
Next
Seuraava = 0

End Sub

Private Sub P3()

Aloittava = 1
Pelaaja3
For j = 1 To 30000
Next
Aloittava = 0
Pelaaja4
For j = 1 To 30000
Next
Seuraava = 2
End Sub



Private Sub P4()
    Aloittava = 1
    Pelaaja4
    For j = 1 To 30000
    Next
    Seuraava = 2
End Sub

Private Sub jatkaa4()

Aloittava = 0
Pelaaja2
For j = 1 To 30000
Next
Aloittava = 0
Pelaaja3
Tarkista_KierrosVoitto
Odotus

End Sub


Public Sub Pelaaja2()

Tutki_Koneen2kortit
sbarTilapalkki.SimpleText = "Valitse pelattava kortti."
If PelattuP2 = 0 And Kierros <> 6 And Aloittava = 0 Then
For i = 1 To 6
    Select Case Johtokortti
    Case 1, 2, 3, 4
        If P2Kortit(i) >= 1 And P2Kortit(i) <= 52 Then
            PelattuP2 = P2Kortit(i)
            pcPelattuP2.Picture = pcKortinkuvat.GraphicCell(P2Kortit(i))
            Johtokortti = P2Kortit(i)
            P2Kortit(i) = 0
            pcP2K(i - 1).Visible = False
            Exit Sub
        End If
    Case 5, 6, 7, 8
        If P2Kortit(i) >= 5 And P2Kortit(i) <= 52 Then
            PelattuP2 = P2Kortit(i)
            pcPelattuP2.Picture = pcKortinkuvat.GraphicCell(P2Kortit(i))
            Johtokortti = P2Kortit(i)
            P2Kortit(i) = 0
            pcP2K(i - 1).Visible = False
            Exit Sub
        End If
    Case 9, 10, 11, 12
        If P2Kortit(i) >= 9 And P2Kortit(i) <= 52 Then
            PelattuP2 = P2Kortit(i)
            pcPelattuP2.Picture = pcKortinkuvat.GraphicCell(P2Kortit(i))
            Johtokortti = P2Kortit(i)
            P2Kortit(i) = 0
            pcP2K(i - 1).Visible = False
            Exit Sub
        End If
    Case 13, 14, 15, 16
        If P2Kortit(i) >= 13 And P2Kortit(i) <= 52 Then
            PelattuP2 = P2Kortit(i)
            pcPelattuP2.Picture = pcKortinkuvat.GraphicCell(P2Kortit(i))
            Johtokortti = P2Kortit(i)
            P2Kortit(i) = 0
            pcP2K(i - 1).Visible = False
            Exit Sub
        End If
    Case 17, 18, 19, 20
        If P2Kortit(i) >= 17 And P2Kortit(i) <= 52 Then
            PelattuP2 = P2Kortit(i)
            pcPelattuP2.Picture = pcKortinkuvat.GraphicCell(P2Kortit(i))
            Johtokortti = P2Kortit(i)
            P2Kortit(i) = 0
            pcP2K(i - 1).Visible = False
            Exit Sub
        End If
    Case 21, 22, 23, 24
        If P2Kortit(i) >= 21 And P2Kortit(i) <= 52 Then
            PelattuP2 = P2Kortit(i)
            pcPelattuP2.Picture = pcKortinkuvat.GraphicCell(P2Kortit(i))
            Johtokortti = P2Kortit(i)
            P2Kortit(i) = 0
            pcP2K(i - 1).Visible = False
            Exit Sub
        End If
    Case 25, 26, 27, 28
        If P2Kortit(i) >= 25 And P2Kortit(i) <= 52 Then
            PelattuP2 = P2Kortit(i)
            pcPelattuP2.Picture = pcKortinkuvat.GraphicCell(P2Kortit(i))
            Johtokortti = P2Kortit(i)
            P2Kortit(i) = 0
            pcP2K(i - 1).Visible = False
            Exit Sub
        End If
    Case 29, 30, 31, 32
        If P2Kortit(i) >= 29 And P2Kortit(i) <= 52 Then
            PelattuP2 = P2Kortit(i)
            pcPelattuP2.Picture = pcKortinkuvat.GraphicCell(P2Kortit(i))
            Johtokortti = P2Kortit(i)
            P2Kortit(i) = 0
            pcP2K(i - 1).Visible = False
            Exit Sub
        End If
    Case 33, 34, 35, 36
        If P2Kortit(i) >= 33 And P2Kortit(i) <= 52 Then
            PelattuP2 = P2Kortit(i)
            pcPelattuP2.Picture = pcKortinkuvat.GraphicCell(P2Kortit(i))
            Johtokortti = P2Kortit(i)
            P2Kortit(i) = 0
            pcP2K(i - 1).Visible = False
            Exit Sub
        End If
    Case 37, 38, 39, 40
        If P2Kortit(i) >= 37 And P2Kortit(i) <= 52 Then
            PelattuP2 = P2Kortit(i)
            pcPelattuP2.Picture = pcKortinkuvat.GraphicCell(P2Kortit(i))
            Johtokortti = P2Kortit(i)
            P2Kortit(i) = 0
            pcP2K(i - 1).Visible = False
            Exit Sub
        End If
    Case 41, 42, 43, 44
        If P2Kortit(i) >= 41 And P2Kortit(i) <= 52 Then
            PelattuP2 = P2Kortit(i)
            pcPelattuP2.Picture = pcKortinkuvat.GraphicCell(P2Kortit(i))
            Johtokortti = P2Kortit(i)
            P2Kortit(i) = 0
            pcP2K(i - 1).Visible = False
            Exit Sub
        End If
    Case 45, 46, 47, 48
        If P2Kortit(i) >= 45 And P2Kortit(i) <= 52 Then
            PelattuP2 = P2Kortit(i)
            pcPelattuP2.Picture = pcKortinkuvat.GraphicCell(P2Kortit(i))
            Johtokortti = P2Kortit(i)
            P2Kortit(i) = 0
            pcP2K(i - 1).Visible = False
            Exit Sub
        End If
    Case 49, 50, 51, 52
        If P2Kortit(i) >= 49 And P2Kortit(i) <= 52 Then
            PelattuP2 = P2Kortit(i)
            pcPelattuP2.Picture = pcKortinkuvat.GraphicCell(P2Kortit(i))
            Johtokortti = P2Kortit(i)
            P2Kortit(i) = 0
            pcP2K(i - 1).Visible = False
            Exit Sub
        End If
    
    End Select

Next

Else

    PelattuP2 = Pelattavakortti
    pcPelattuP2.Picture = pcKortinkuvat.GraphicCell(johKortti)
    Johtokortti = johKortti
    P2Kortit(tyhKortti) = 0
    pcP2K(tyhKortti - 1).Visible = False
    Aloittava = 0
    Exit Sub

End If

For i = 1 To 6
    If P2Kortit(i) <> 0 Then
        PelattuP2 = P2Kortit(i)
        pcPelattuP2.Picture = pcKortinkuvat.GraphicCell(P2Kortit(i))
        P2Kortit(i) = 0
        pcP2K(i - 1).Visible = False
        Exit Sub
    End If
Next


End Sub

Public Sub Soita_Aani()

If mnuAani.Checked = True Then
    MMControl1.Command = "Prev"
    MMControl1.Command = "Play"
End If

End Sub

Public Sub TarkistaPisteet()

Sijoitus = 1
If Kurkkuikkuna.Kokonaispisteet1 <= Kurkkuikkuna.Kokonaispisteet2 And Kurkkuikkuna.Kokonaispisteet1 <= Kurkkuikkuna.Kokonaispisteet3 And Kurkkuikkuna.Kokonaispisteet1 <= Kurkkuikkuna.Kokonaispisteet4 And Kurkkuikkuna.PVoitto = 1 Then
    Sijoitus = 1
End If

If Kokonaispisteet1 > Kokonaispisteet2 Then
    Sijoitus = Sijoitus + 1
End If

If Kokonaispisteet1 > Kokonaispisteet3 Then
    Sijoitus = Sijoitus + 1
End If

If Kokonaispisteet1 > Kokonaispisteet4 Then
    Sijoitus = Sijoitus + 1
End If

Select Case Sijoitus
Case 1
Pisteikkuna.Caption = "Pistetilanne - olet ensimm‰inen"
Case 2
Pisteikkuna.Caption = "Pistetilanne - olet toinen"
Case 3
Pisteikkuna.Caption = "Pistetilanne - olet kolmas"
Case 4
Pisteikkuna.Caption = "Pistetilanne - olet nelj‰s"
End Select
End Sub

Public Sub Tutki_Koneen2kortit()

If pcP2K(0) <> 0 Then
tKortti1 = P2Kortit(1)
End If
If pcP2K(1) <> 0 Then
tKortti2 = P2Kortit(2)
End If
If pcP2K(2) <> 0 Then
tKortti3 = P2Kortit(3)
End If
If pcP2K(3) <> 0 Then
tKortti4 = P2Kortit(4)
End If
If pcP2K(4) <> 0 Then
tKortti5 = P2Kortit(5)
End If
If pcP2K(5) <> 0 Then
tKortti6 = P2Kortit(6)
End If

If tKortti4 <> 0 Then
    Pelattavakortti = tKortti4
    johKortti = P2Kortit(4)
    tyhKortti = 4
Else
    If tKortti5 <> 0 Then
        Pelattavakortti = tKortti5
        johKortti = P2Kortit(5)
        tyhKortti = 5
    Else
        If tKortti6 <> 0 Then
            Pelattavakortti = tKortti6
            johKortti = P2Kortit(6)
            tyhKortti = 6
        Else
            If tKortti3 <> 0 Then
                Pelattavakortti = tKortti3
                johKortti = P2Kortit(3)
                tyhKortti = 3
            Else
                If tKortti2 <> 0 Then
                    Pelattavakortti = tKortti2
                    johKortti = P2Kortit(2)
                    tyhKortti = 2
                Else
                    If tKortti1 <> 0 Then
                        Pelattavakortti = tKortti1
                        johKortti = P2Kortit(1)
                        tyhKortti = 1
                    End If
End If
End If
End If
End If
End If

End Sub

Public Sub Tutki_Koneen3kortit()

If pcP3K(0) <> 0 Then
tKortti1 = P3Kortit(1)
End If
If pcP3K(1) <> 0 Then
tKortti2 = P3Kortit(2)
End If
If pcP3K(2) <> 0 Then
tKortti3 = P3Kortit(3)
End If
If pcP3K(3) <> 0 Then
tKortti4 = P3Kortit(4)
End If
If pcP3K(4) <> 0 Then
tKortti5 = P3Kortit(5)
End If
If pcP3K(5) <> 0 Then
tKortti6 = P3Kortit(6)
End If

If tKortti4 <> 0 Then
    Pelattavakortti = tKortti4
    johKortti = P3Kortit(4)
    tyhKortti = 4
Else
    If tKortti5 <> 0 Then
        Pelattavakortti = tKortti5
        johKortti = P3Kortit(5)
        tyhKortti = 5
    Else
        If tKortti6 <> 0 Then
            Pelattavakortti = tKortti6
            johKortti = P3Kortit(6)
            tyhKortti = 6
        Else
            If tKortti3 <> 0 Then
                Pelattavakortti = tKortti3
                johKortti = P3Kortit(3)
                tyhKortti = 3
            Else
                If tKortti2 <> 0 Then
                    Pelattavakortti = tKortti2
                    johKortti = P3Kortit(2)
                    tyhKortti = 2
                Else
                    If tKortti1 <> 0 Then
                        Pelattavakortti = tKortti1
                        johKortti = P3Kortit(1)
                        tyhKortti = 1
                    End If
End If
End If
End If
End If
End If

End Sub

Public Sub Tutki_Koneen4kortit()

If pcP4k(0) <> 0 Then
tKortti1 = P4Kortit(1)
End If
If pcP4k(1) <> 0 Then
tKortti2 = P4Kortit(2)
End If
If pcP4k(2) <> 0 Then
tKortti3 = P4Kortit(3)
End If
If pcP4k(3) <> 0 Then
tKortti4 = P4Kortit(4)
End If
If pcP4k(4) <> 0 Then
tKortti5 = P4Kortit(5)
End If
If pcP4k(5) <> 0 Then
tKortti6 = P4Kortit(6)
End If

If tKortti4 <> 0 Then
    Pelattavakortti = tKortti4
    johKortti = P4Kortit(4)
    tyhKortti = 4
Else
    If tKortti5 <> 0 Then
        Pelattavakortti = tKortti5
        johKortti = P4Kortit(5)
        tyhKortti = 5
    Else
        If tKortti6 <> 0 Then
            Pelattavakortti = tKortti6
            johKortti = P4Kortit(6)
            tyhKortti = 6
        Else
            If tKortti3 <> 0 Then
                Pelattavakortti = tKortti3
                johKortti = P4Kortit(3)
                tyhKortti = 3
            Else
                If tKortti2 <> 0 Then
                    Pelattavakortti = tKortti2
                    johKortti = P4Kortit(2)
                    tyhKortti = 2
                Else
                    If tKortti1 <> 0 Then
                        Pelattavakortti = tKortti1
                        johKortti = P4Kortit(1)
                        tyhKortti = 1
                    End If
End If
End If
End If
End If
End If

End Sub


Public Sub Pelaaja4()

If Kurkkuikkuna.mnuAani.Checked = True Then
End If
Tutki_Koneen4kortit
sbarTilapalkki.SimpleText = "Valitse pelattava kortti."
If PelattuP4 = 0 And Kierros <> 6 And Aloittava = 0 Then
For i = 1 To 6
    Select Case Johtokortti
    Case 1, 2, 3, 4
        If P4Kortit(i) >= 1 And P4Kortit(i) <= 52 Then
            PelattuP4 = P4Kortit(i)
            pcPelattuP4.Picture = pcKortinkuvat.GraphicCell(P4Kortit(i))
            Johtokortti = P4Kortit(i)
            P4Kortit(i) = 0
            pcP4k(i - 1).Visible = False
            Exit Sub
        End If
    Case 5, 6, 7, 8
        If P4Kortit(i) >= 5 And P4Kortit(i) <= 52 Then
            PelattuP4 = P4Kortit(i)
            pcPelattuP4.Picture = pcKortinkuvat.GraphicCell(P4Kortit(i))
            Johtokortti = P4Kortit(i)
            P4Kortit(i) = 0
            pcP4k(i - 1).Visible = False
            Exit Sub
        End If
    Case 9, 10, 11, 12
        If P4Kortit(i) >= 9 And P4Kortit(i) <= 52 Then
            PelattuP4 = P4Kortit(i)
            pcPelattuP4.Picture = pcKortinkuvat.GraphicCell(P4Kortit(i))
            Johtokortti = P4Kortit(i)
            P4Kortit(i) = 0
            pcP4k(i - 1).Visible = False
            Exit Sub
        End If
    Case 13, 14, 15, 16
        If P4Kortit(i) >= 13 And P4Kortit(i) <= 52 Then
            PelattuP4 = P4Kortit(i)
            pcPelattuP4.Picture = pcKortinkuvat.GraphicCell(P4Kortit(i))
            Johtokortti = P4Kortit(i)
            P4Kortit(i) = 0
            pcP4k(i - 1).Visible = False
            Exit Sub
        End If
    Case 17, 18, 19, 20
        If P4Kortit(i) >= 17 And P4Kortit(i) <= 52 Then
            PelattuP4 = P4Kortit(i)
            pcPelattuP4.Picture = pcKortinkuvat.GraphicCell(P4Kortit(i))
            Johtokortti = P4Kortit(i)
            P4Kortit(i) = 0
            pcP4k(i - 1).Visible = False
            Exit Sub
        End If
    Case 21, 22, 23, 24
        If P4Kortit(i) >= 21 And P4Kortit(i) <= 52 Then
            PelattuP4 = P4Kortit(i)
            pcPelattuP4.Picture = pcKortinkuvat.GraphicCell(P4Kortit(i))
            Johtokortti = P4Kortit(i)
            P4Kortit(i) = 0
            pcP4k(i - 1).Visible = False
            Exit Sub
        End If
    Case 25, 26, 27, 28
        If P4Kortit(i) >= 25 And P4Kortit(i) <= 52 Then
            PelattuP4 = P4Kortit(i)
            pcPelattuP4.Picture = pcKortinkuvat.GraphicCell(P4Kortit(i))
            Johtokortti = P4Kortit(i)
            P4Kortit(i) = 0
            pcP4k(i - 1).Visible = False
            Exit Sub
        End If
    Case 29, 30, 31, 32
        If P4Kortit(i) >= 29 And P4Kortit(i) <= 52 Then
            PelattuP4 = P4Kortit(i)
            pcPelattuP4.Picture = pcKortinkuvat.GraphicCell(P4Kortit(i))
            Johtokortti = P4Kortit(i)
            P4Kortit(i) = 0
            pcP4k(i - 1).Visible = False
            Exit Sub
        End If
    Case 33, 34, 35, 36
        If P4Kortit(i) >= 33 And P4Kortit(i) <= 52 Then
            PelattuP4 = P4Kortit(i)
            pcPelattuP4.Picture = pcKortinkuvat.GraphicCell(P4Kortit(i))
            Johtokortti = P4Kortit(i)
            P4Kortit(i) = 0
            pcP4k(i - 1).Visible = False
            Exit Sub
        End If
    Case 37, 38, 39, 40
        If P4Kortit(i) >= 37 And P4Kortit(i) <= 52 Then
            PelattuP4 = P4Kortit(i)
            pcPelattuP4.Picture = pcKortinkuvat.GraphicCell(P4Kortit(i))
            Johtokortti = P4Kortit(i)
            P4Kortit(i) = 0
            pcP4k(i - 1).Visible = False
            Exit Sub
        End If
    Case 41, 42, 43, 44
        If P4Kortit(i) >= 41 And P4Kortit(i) <= 52 Then
            PelattuP4 = P4Kortit(i)
            pcPelattuP4.Picture = pcKortinkuvat.GraphicCell(P4Kortit(i))
            Johtokortti = P4Kortit(i)
            P4Kortit(i) = 0
            pcP4k(i - 1).Visible = False
            Exit Sub
        End If
    Case 45, 46, 47, 48
        If P4Kortit(i) >= 45 And P4Kortit(i) <= 52 Then
            PelattuP4 = P4Kortit(i)
            pcPelattuP4.Picture = pcKortinkuvat.GraphicCell(P4Kortit(i))
            Johtokortti = P4Kortit(i)
            P4Kortit(i) = 0
            pcP4k(i - 1).Visible = False
            Exit Sub
        End If
    Case 49, 50, 51, 52
        If P4Kortit(i) >= 49 And P4Kortit(i) <= 52 Then
            PelattuP4 = P4Kortit(i)
            pcPelattuP4.Picture = pcKortinkuvat.GraphicCell(P4Kortit(i))
            Johtokortti = P4Kortit(i)
            P4Kortit(i) = 0
            pcP4k(i - 1).Visible = False
            Exit Sub
        End If
    End Select
    
Next

Else

    PelattuP4 = Pelattavakortti
    pcPelattuP4.Picture = pcKortinkuvat.GraphicCell(johKortti)
    Johtokortti = johKortti
    P4Kortit(tyhKortti) = 0
    pcP4k(tyhKortti - 1).Visible = False
    Aloittava = 0
    Exit Sub

End If

For i = 1 To 6
    If P4Kortit(i) <> 0 Then
        PelattuP4 = P4Kortit(i)
        pcPelattuP4.Picture = pcKortinkuvat.GraphicCell(P4Kortit(i))
        P4Kortit(i) = 0
        pcP4k(i - 1).Visible = False
        Exit Sub
    End If
Next

End Sub
Public Sub Pelaaja3()

If Kurkkuikkuna.mnuAani.Checked = True Then
End If
Tutki_Koneen3kortit
sbarTilapalkki.SimpleText = "Valitse pelattava kortti."
If PelattuP3 = 0 And Kierros <> 6 And Aloittava = 0 Then
For i = 1 To 6
    Select Case Johtokortti
    Case 1, 2, 3, 4
        If P3Kortit(i) >= 1 And P3Kortit(i) <= 52 Then
            PelattuP3 = P3Kortit(i)
            pcPelattuP3.Picture = pcKortinkuvat.GraphicCell(P3Kortit(i))
            Johtokortti = P3Kortit(i)
            P3Kortit(i) = 0
            pcP3K(i - 1).Visible = False
            Exit Sub
        End If
    Case 5, 6, 7, 8
        If P3Kortit(i) >= 5 And P3Kortit(i) <= 52 Then
            PelattuP3 = P3Kortit(i)
            pcPelattuP3.Picture = pcKortinkuvat.GraphicCell(P3Kortit(i))
            Johtokortti = P3Kortit(i)
            P3Kortit(i) = 0
            pcP3K(i - 1).Visible = False
            Exit Sub
        End If
    Case 9, 10, 11, 12
        If P3Kortit(i) >= 9 And P3Kortit(i) <= 52 Then
            PelattuP3 = P3Kortit(i)
            pcPelattuP3.Picture = pcKortinkuvat.GraphicCell(P3Kortit(i))
            Johtokortti = P3Kortit(i)
            P3Kortit(i) = 0
            pcP3K(i - 1).Visible = False
            Exit Sub
        End If
    Case 13, 14, 15, 16
        If P3Kortit(i) >= 13 And P3Kortit(i) <= 52 Then
            PelattuP3 = P3Kortit(i)
            pcPelattuP3.Picture = pcKortinkuvat.GraphicCell(P3Kortit(i))
            Johtokortti = P3Kortit(i)
            P3Kortit(i) = 0
            pcP3K(i - 1).Visible = False
            Exit Sub
        End If
    Case 17, 18, 19, 20
        If P3Kortit(i) >= 17 And P3Kortit(i) <= 52 Then
            PelattuP3 = P3Kortit(i)
            pcPelattuP3.Picture = pcKortinkuvat.GraphicCell(P3Kortit(i))
            Johtokortti = P3Kortit(i)
            P3Kortit(i) = 0
            pcP3K(i - 1).Visible = False
            Exit Sub
        End If
    Case 21, 22, 23, 24
        If P3Kortit(i) >= 21 And P3Kortit(i) <= 52 Then
            PelattuP3 = P3Kortit(i)
            pcPelattuP3.Picture = pcKortinkuvat.GraphicCell(P3Kortit(i))
            Johtokortti = P3Kortit(i)
            P3Kortit(i) = 0
            pcP3K(i - 1).Visible = False
            Exit Sub
        End If
    Case 25, 26, 27, 28
        If P3Kortit(i) >= 25 And P3Kortit(i) <= 52 Then
            PelattuP3 = P3Kortit(i)
            pcPelattuP3.Picture = pcKortinkuvat.GraphicCell(P3Kortit(i))
            Johtokortti = P3Kortit(i)
            P3Kortit(i) = 0
            pcP3K(i - 1).Visible = False
            Exit Sub
        End If
    Case 29, 30, 31, 32
        If P3Kortit(i) >= 29 And P3Kortit(i) <= 52 Then
            PelattuP3 = P3Kortit(i)
            pcPelattuP3.Picture = pcKortinkuvat.GraphicCell(P3Kortit(i))
            Johtokortti = P3Kortit(i)
            P3Kortit(i) = 0
            pcP3K(i - 1).Visible = False
            Exit Sub
        End If
    Case 33, 34, 35, 36
        If P3Kortit(i) >= 33 And P3Kortit(i) <= 52 Then
            PelattuP3 = P3Kortit(i)
            pcPelattuP3.Picture = pcKortinkuvat.GraphicCell(P3Kortit(i))
            Johtokortti = P3Kortit(i)
            P3Kortit(i) = 0
            pcP3K(i - 1).Visible = False
            Exit Sub
        End If
    Case 37, 38, 39, 40
        If P3Kortit(i) >= 37 And P3Kortit(i) <= 52 Then
            PelattuP3 = P3Kortit(i)
            pcPelattuP3.Picture = pcKortinkuvat.GraphicCell(P3Kortit(i))
            Johtokortti = P3Kortit(i)
            P3Kortit(i) = 0
            pcP3K(i - 1).Visible = False
            Exit Sub
        End If
    Case 41, 42, 43, 44
        If P3Kortit(i) >= 41 And P3Kortit(i) <= 52 Then
            PelattuP3 = P3Kortit(i)
            pcPelattuP3.Picture = pcKortinkuvat.GraphicCell(P3Kortit(i))
            Johtokortti = P3Kortit(i)
            P3Kortit(i) = 0
            pcP3K(i - 1).Visible = False
            Exit Sub
        End If
    Case 45, 46, 47, 48
        If P3Kortit(i) >= 45 And P3Kortit(i) <= 52 Then
            PelattuP3 = P3Kortit(i)
            pcPelattuP3.Picture = pcKortinkuvat.GraphicCell(P3Kortit(i))
            Johtokortti = P3Kortit(i)
            P3Kortit(i) = 0
            pcP3K(i - 1).Visible = False
            Exit Sub
        End If
    Case 49, 50, 51, 52
        If P3Kortit(i) >= 49 And P3Kortit(i) <= 52 Then
            PelattuP3 = P3Kortit(i)
            pcPelattuP3.Picture = pcKortinkuvat.GraphicCell(P3Kortit(i))
            Johtokortti = P3Kortit(i)
            P3Kortit(i) = 0
            pcP3K(i - 1).Visible = False
            Exit Sub
        End If
    End Select
Next

Else

    PelattuP3 = Pelattavakortti
    pcPelattuP3.Picture = pcKortinkuvat.GraphicCell(johKortti)
    Johtokortti = johKortti
    P3Kortit(tyhKortti) = 0
    pcP3K(tyhKortti - 1).Visible = False
    Aloittava = 0
    Exit Sub

End If

For i = 1 To 6
    If P3Kortit(i) <> 0 Then
        PelattuP3 = P3Kortit(i)
        pcPelattuP3.Picture = pcKortinkuvat.GraphicCell(P3Kortit(i))
        P3Kortit(i) = 0
        pcP3K(i - 1).Visible = False
        Exit Sub
    End If
Next

End Sub

Public Sub Kirjoita_Asetukset()

Dim File As String
File = App.Path + "\Settings.dat"
Open File For Random As #1

A = Kurkkuikkuna.lblPelaaja1.Caption
B = Kurkkuikkuna.lblPelaaja2.Caption
C = Kurkkuikkuna.lblPelaaja3.Caption
D = Kurkkuikkuna.lblPelaaja4.Caption
e = Kurkkuikkuna.Pisteraja
f = Kurkkuikkuna.mnuAani.Checked
H = Kurkkuikkuna.Top
j = Kurkkuikkuna.Left

Put #1, , A
Put #1, , B
Put #1, , C
Put #1, , D
Put #1, , e
Put #1, , f
Put #1, , H
Put #1, , j

Close #1

End Sub


Public Sub Tarkista_OmaKortti()
Select Case Johtokortti

Case 1, 2, 3, 4
    Vali = 2
Case 5, 6, 7, 8
    Vali = 3
Case 9, 10, 11, 12
    Vali = 4
Case 13, 14, 15, 16
    Vali = 5
Case 17, 18, 19, 20
    Vali = 6
Case 21, 22, 23, 24
    Vali = 7
Case 25, 26, 27, 28
    Vali = 8
Case 29, 30, 31, 32
    Vali = 9
Case 33, 34, 35, 36
    Vali = 10
Case 37, 38, 39, 40
    Vali = 11
Case 41, 42, 43, 44
    Vali = 12
Case 45, 46, 47, 48
    Vali = 13
Case 49, 50, 51, 52
    Vali = 14

End Select
Select Case PelattuP1

Case 1, 2, 3, 4
    Vali2 = 2
Case 5, 6, 7, 8
    Vali2 = 3
Case 9, 10, 11, 12
    Vali2 = 4
Case 13, 14, 15, 16
    Vali2 = 5
Case 17, 18, 19, 20
    Vali2 = 6
Case 21, 22, 23, 24
    Vali2 = 7
Case 25, 26, 27, 28
    Vali2 = 8
Case 29, 30, 31, 32
    Vali2 = 9
Case 33, 34, 35, 36
    Vali2 = 10
Case 37, 38, 39, 40
    Vali2 = 11
Case 41, 42, 43, 44
    Vali2 = 12
Case 45, 46, 47, 48
    Vali2 = 13
Case 49, 50, 51, 52
    Vali2 = 14

End Select

End Sub

Public Sub Tarkista_OmatKortit()

    
    Pienin = 100
    Suurin = -100

For i = 1 To 6
    If P1Kortit(i) <> 0 Then
        If P1Kortit(i) >= Suurin Then
            Suurin = P1Kortit(i)
        End If
    End If
Next

For i = 6 To 1 Step -1
    If P1Kortit(i) <> 0 Then
        If P1Kortit(i) <= Pienin Then
            Pienin = P1Kortit(i)
        End If
    End If
Next

Select Case Suurin

Case 1, 2, 3, 4
    Suurin = 2
Case 5, 6, 7, 8
    Suurin = 3
Case 9, 10, 11, 12
    Suurin = 4
Case 13, 14, 15, 16
    Suurin = 5
Case 17, 18, 19, 20
    Suurin = 6
Case 21, 22, 23, 24
    Suurin = 7
Case 25, 26, 27, 28
    Suurin = 8
Case 29, 30, 31, 32
    Suurin = 9
Case 33, 34, 35, 36
    Suurin = 10
Case 37, 38, 39, 40
    Suurin = 11
Case 41, 42, 43, 44
    Suurin = 12
Case 45, 46, 47, 48
    Suurin = 13
Case 49, 50, 51, 52
    Suurin = 14

End Select

Select Case Pienin

Case 1, 2, 3, 4
    Pienin = 2
Case 5, 6, 7, 8
    Pienin = 3
Case 9, 10, 11, 12
    Pienin = 4
Case 13, 14, 15, 16
    Pienin = 5
Case 17, 18, 19, 20
    Pienin = 6
Case 21, 22, 23, 24
    Pienin = 7
Case 25, 26, 27, 28
    Pienin = 8
Case 29, 30, 31, 32
    Pienin = 9
Case 33, 34, 35, 36
    Pienin = 10
Case 37, 38, 39, 40
    Pienin = 11
Case 41, 42, 43, 44
    Pienin = 12
Case 45, 46, 47, 48
    Pienin = 13
Case 49, 50, 51, 52
    Pienin = 14

End Select

End Sub


Private Sub Tyhjays()
tmrAika.Enabled = False
pcPelattuP1.Picture = LoadPicture()
pcPelattuP2.Picture = LoadPicture()
pcPelattuP3.Picture = LoadPicture()
pcPelattuP4.Picture = LoadPicture()
PelattuP1 = 0
PelattuP2 = 0
PelattuP3 = 0
PelattuP4 = 0
Johtokortti = 0


If Kierros <> 6 Then

Select Case Vuoro

Case 1
    Aloittava = 0
    Exit Sub
Case 2
    P2
    Exit Sub
Case 3
    P3
    Exit Sub
Case 4
    P4
    Exit Sub
End Select

End If

Tarkista_KierrosVoitto

End Sub

Public Sub UusiJako()

Aani_Sekoitus
Soita_Aani
If Kurkkuikkuna.mnuAani.Checked = True Then
    Do Until Kurkkuikkuna.MMControl1.Mode = 525
    Loop
End If
For i = 0 To 5
pcP1K(i).Visible = True
pcP2K(i).Visible = True
pcP3K(i).Visible = True
pcP4k(i).Visible = True
pcP1K(i).Enabled = True
pcP2K(i).Enabled = True
pcP3K(i).Enabled = True
pcP4k(i).Enabled = True
Next
Kurkkuikkuna.Kortit_Alussa
Kurkkuikkuna.Jaa_Kortit
Kurkkuikkuna.Jaa_Kortit2
Kurkkuikkuna.Jaa_Kortit3
Kurkkuikkuna.Jaa_Kortit4
If Aloitusvuoro = 4 Then
    Aloitusvuoro = 1
    Vuoro = 1
Else
    Aloitusvuoro = Aloitusvuoro + 1
    Vuoro = Vuoro + 1
End If

Kurkkuikkuna.sbarTilapalkki.SimpleText = "Valitse pelattava kortti."

End Sub

Public Sub Voitto()


For i = 0 To 5
    pcP2K(i).Picture = pcKortinkuvat.GraphicCell(P2Kortit(i + 1))
    pcP3K(i).Picture = pcKortinkuvat.GraphicCell(P3Kortit(i + 1))
    pcP4k(i).Picture = pcKortinkuvat.GraphicCell(P4Kortit(i + 1))
    pcP1K(i).Enabled = False
Next

For i = 0 To 6
    Pisteet1 = Pisteet1 + P1Kortit(i)
    Pisteet2 = Pisteet2 + P2Kortit(i)
    Pisteet3 = Pisteet3 + P3Kortit(i)
    Pisteet4 = Pisteet4 + P4Kortit(i)
Next

Muuta_pisteet

If Kortti1 >= Kortti2 And Kortti1 >= Kortti3 And Kortti1 >= Kortti4 Then
Haviaja1 = 1
End If

If Kortti2 >= Kortti1 And Kortti2 >= Kortti3 And Kortti2 >= Kortti4 Then
Haviaja2 = 1
End If

If Kortti3 >= Kortti2 And Kortti3 >= Kortti1 And Kortti3 >= Kortti4 Then
Haviaja3 = 1
End If

If Kortti4 >= Kortti2 And Kortti4 >= Kortti3 And Kortti4 >= Kortti1 Then
Haviaja4 = 1
End If

OdotusArvo = 1
Odotus

End Sub

Private Sub Form_Load()

moodi = MMControl1.Mode
App.HelpFile = App.Path + "\Kurkku.hlp"
Lue_Asetukset
Aloitusvuoro = 1
Vuoro = 1
Kierros = 1
Lataa_avausikkuna
OdotusArvo = 0
End Sub

Public Sub Lue_Asetukset()

Dim File As String

File = App.Path + "\Settings.dat"
Open File For Random As #1

Get #1, , A
Get #1, , B
Get #1, , C
Get #1, , D
Get #1, , e
Get #1, , f
Get #1, , H
Get #1, , j

Avausikkuna.txtNimi.Text = A
Kurkkuikkuna.lblPelaaja2.Caption = B
Kurkkuikkuna.lblPelaaja3.Caption = C
Kurkkuikkuna.lblPelaaja4.Caption = D
Kurkkuikkuna.Pisteraja = e
Kurkkuikkuna.mnuAani.Checked = f
Kurkkuikkuna.Top = H
Kurkkuikkuna.Left = j

Close #1

End Sub

Private Sub Form_Unload(Cancel As Integer)

Kurkkuikkuna.Kirjoita_Asetukset
End

End Sub



Private Sub mnuAani_Click()

If mnuAani.Checked <> True Then
    mnuAani.Checked = True
Else
    mnuAani.Checked = False
End If

End Sub


Private Sub mnuAsetukset_Click()

Load Asetuksetikkuna
Asetuksetikkuna.Top = Kurkkuikkuna.Top + 1300
Asetuksetikkuna.Left = Kurkkuikkuna.Left + 500
Asetuksetikkuna.Visible = True
Asetuksetikkuna.ZOrder 0

End Sub

Private Sub mnuInternet_Click()

Pathi = App.Path
Dim i As Long
i = ShellExecute(Me.hWnd, vbNullString, "http://members.xoom.com/tesisoftware", vbNullString, Pathi, SW_SHOWNORMAL)

End Sub

Private Sub mnuLopeta_Click()

Unload Kurkkuikkuna

End Sub


Private Sub mnuMotto_Click()

Load Mottoikkuna
Mottoikkuna.Top = Kurkkuikkuna.Top + 1800
Mottoikkuna.Left = Kurkkuikkuna.Left + 600
Mottoikkuna.Visible = True
Mottoikkuna.ZOrder 0

End Sub


Private Sub mnuOhjeenAiheet_Click()

Dim helppi As String
helppi = App.Path + "\Kurkku.hlp"
helppi = "winhlp32 " + helppi
Dim X
X = Shell(helppi, 1)

End Sub

Private Sub mnuPistetilanne_Click()

Pelikaynnissa = 1
Load Pisteikkuna
Pisteikkuna.Top = Kurkkuikkuna.Top + 2100
Pisteikkuna.Left = Kurkkuikkuna.Left + 1250
If Kokonaispisteet1 >= Pisteraja Then
    Pisteikkuna.Caption = "Pelin loppu"
    sbarTilapalkki.SimpleText = "Peli loppu, Aloita uusi peli painamalla F2."
    Pisteikkuna.lblPelaaja1.ForeColor = &HFF&
    For i = 0 To 12
    Pisteikkuna.Pisteet1(i).ForeColor = &HFF&
    Next
    Pisteikkuna.imgKuva.Visible = True
    Pelikaynnissa = 1
End If
If Kokonaispisteet2 >= Pisteraja Then
    Pisteikkuna.Caption = "Pelin loppu"
    sbarTilapalkki.SimpleText = "Peli loppu, Aloita uusi peli painamalla F2."
    Pisteikkuna.lblPelaaja2.ForeColor = &HFF&
    For i = 0 To 12
    Pisteikkuna.Pisteet2(i).ForeColor = &HFF&
    Next
    Pisteikkuna.imgKuva.Visible = True
    Pelikaynnissa = 1
End If
If Kokonaispisteet3 >= Pisteraja Then
    Pisteikkuna.Caption = "Pelin loppu"
    sbarTilapalkki.SimpleText = "Peli loppu, Aloita uusi peli painamalla F2."
    Pisteikkuna.lblPelaaja3.ForeColor = &HFF&
    For i = 0 To 12
    Pisteikkuna.Pisteet3(i).ForeColor = &HFF&
    Next
    Pisteikkuna.imgKuva.Visible = True
    Pelikaynnissa = 1
End If
If Kokonaispisteet4 >= Pisteraja Then
    Pisteikkuna.Caption = "Pelin loppu"
    sbarTilapalkki.SimpleText = "Peli loppu, Aloita uusi peli painamalla F2."
    Pisteikkuna.lblPelaaja4.ForeColor = &HFF&
    For i = 0 To 12
    Pisteikkuna.Pisteet4(i).ForeColor = &HFF&
    Next
    Pisteikkuna.imgKuva.Visible = True
    Pelikaynnissa = 1
End If
Pisteikkuna.Visible = True
Pisteikkuna.ZOrder 0

End Sub

Private Sub mnusahkoposti_Click()

Pathi = App.Path
Dim i As Long
i = ShellExecute(Me.hWnd, vbNullString, "mailto:tesisoftware@usa.net", vbNullString, Pathi, SW_SHOWNORMAL)

End Sub

Private Sub mnuTietojaKurkusta_Click()

Load Tietojaikkuna
Tietojaikkuna.Top = Kurkkuikkuna.Top + 1100
Tietojaikkuna.Left = Kurkkuikkuna.Left + 500
Tietojaikkuna.Visible = True
Tietojaikkuna.ZOrder 0

End Sub


Private Sub mnuUusipeli_Click()

Aloitusvuoro = 0
Nollaus_alku
Nollaus
UusiJako

End Sub

Public Sub pcP1k_Click(index As Integer)

If Aloitettu = 0 Then
Select Case Aloitusvuoro
Case 1
    Aloitettu = 1
    Vuoro = 1
Case 2
    Aloitettu = 1
    Vuoro = 2
    P2
    Exit Sub
Case 3
    Aloitettu = 1
    Vuoro = 3
    P3
    Exit Sub
Case 4
    Aloitettu = 1
    Vuoro = 4
    P4
    Exit Sub
End Select
End If

Tarkista_OmatKortit
PelattuP1 = P1Kortit(index + 1)
If Vuoro <> 1 Then
Tarkista_OmaKortti
    
    If Vali2 >= Vali Then
        pcPelattuP1.Picture = pcP1K(index).Picture
        Johtokortti = P1Kortit(index + 1)
        pcP1K(index).Picture = LoadPicture()
        pcP1K(index).Visible = False

        P1Kortit(index + 1) = 0
        If Vuoro <> 2 Then
            GoTo jatkuu
        Else
            If Kurkkuikkuna.mnuAani.Checked = True Then
                Do Until Kurkkuikkuna.MMControl1.Mode = 525
                Loop
            End If
            GoTo tarkistus
        End If
    Else
        If Vali2 < Vali And Suurin >= Vali Then
            sbarTilapalkki.SimpleText = "Sinun on laitettava v‰hint‰‰n yht‰ suuri kortti."
            Exit Sub
        End If
        
        If Suurin < Vali And Vali2 > Pienin Then
            sbarTilapalkki.SimpleText = "Sinun on laitettava pienin korttisi."
            Exit Sub
        End If
        If Suurin < Vali And Vali2 <= Pienin Then
            pcPelattuP1.Picture = pcP1K(index).Picture
            pcP1K(index).Picture = LoadPicture()
            pcP1K(index).Visible = False
            P1Kortit(index + 1) = 0
            If Vuoro <> 2 Then
                GoTo jatkuu
            Else
                If Kurkkuikkuna.mnuAani.Checked = True Then
                    Do Until Kurkkuikkuna.MMControl1.Mode = 525
                    Loop
                End If
                GoTo tarkistus
            End If
        End If
        
    End If

Else
    pcPelattuP1.Picture = pcP1K(index).Picture
    Johtokortti = P1Kortit(index + 1)
    pcP1K(index).Picture = LoadPicture()
    pcP1K(index).Visible = False
    P1Kortit(index + 1) = 0
End If

jatkuu:

If Vuoro = 2 Then Tarkista_KierrosVoitto
If Vuoro = 3 Then
    jatkaa3
    Exit Sub
End If
If Vuoro = 4 Then
    jatkaa4
    Exit Sub
End If

If Kierros <> 6 And Aloittava = 0 Then
Pelaaja2
End If

For j = 1 To 30000
Next

If Kierros <> 6 And Aloittava = 0 Then
Pelaaja3
End If

For j = 1 To 30000
Next

If Kierros <> 6 And Aloittava = 0 Then
Pelaaja4
End If

tarkistus:
Tarkista_KierrosVoitto
If OdotusArvo = 0 Then
    Odotus
End If

End Sub


Public Sub Tarkista_KierrosVoitto()

Pelattu = 0
Kierros = Kierros + 1

Select Case Johtokortti

Case PelattuP1
    Vuoro = 1
Case PelattuP2
    Vuoro = 2
Case PelattuP3
    Vuoro = 3
Case PelattuP4
    Vuoro = 4

End Select

If Kierros = 6 Then
    Voitto
End If

End Sub


Private Sub tmrAika_Timer()

If Aika = Ajanotto Then
    If OdotusArvo = 0 Then
        Tyhjays
    End If
    If OdotusArvo = 1 Then
        LisaaPist
    End If
    
End If
Ajanotto = Ajanotto + 1

End Sub


