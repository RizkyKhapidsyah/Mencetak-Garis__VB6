VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Mencetak Garis"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   2040
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub PrintLine(Width As Single)
   Printer.Line (0, Printer.CurrentY)- _
   (Printer.ScaleWidth, Printer.CurrentY + Width), , BF
   'Gunakan perintah EndDoc jika garis ini adalah garis
   'terakhir yang akan Anda cetak ke dalam kertas
    Printer.EndDoc
End Sub

Private Sub Command1_Click()
    'Ganti 40 di bawah dengan lebar garis yang Anda
    'inginkan.
    PrintLine (40)
End Sub


