VERSION 4.00
Begin VB.Form Mottoikkuna 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "TESI softare Kurkun motto"
   ClientHeight    =   1785
   ClientLeft      =   5985
   ClientTop       =   3585
   ClientWidth     =   3720
   Height          =   2190
   Icon            =   "Mottoikkuna.frx":0000
   Left            =   5925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   Top             =   3240
   Width           =   3840
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   340
      Left            =   1720
      TabIndex        =   1
      Top             =   1280
      Width           =   1120
   End
   Begin VB.Image imgKuva 
      Height          =   480
      Left            =   240
      Picture         =   "Mottoikkuna.frx":000C
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblMotto 
      Caption         =   $"Mottoikkuna.frx":028E
      Height          =   855
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Mottoikkuna"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Private Sub cmdOK_Click()

Kurkkuikkuna.Enabled = True
Unload Mottoikkuna

End Sub

Private Sub Form_Load()

Kurkkuikkuna.Enabled = False

End Sub


