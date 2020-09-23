VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   ScaleHeight     =   6840
   ScaleWidth      =   9945
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Adjustments:"
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2415
      Begin VB.CommandButton cmdApply 
         Caption         =   "&Apply"
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtFontSize 
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Text            =   "12"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Font Size:"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   255
         Width           =   855
      End
   End
   Begin Project1.ucCommentry ucCommentry 
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5040
      Width           =   9255
      _extentx        =   16325
      _extenty        =   2990
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlGap       As Long
Private mlMaxHeight As Long

Private Sub cmdApply_Click()
    pSetFontSize
End Sub

Private Sub Form_Load()
    mlGap = 4 * Screen.TwipsPerPixelX
    With ucCommentry
        mlMaxHeight = .Height
        .Caption = "Memory Allocation"
        .Description = "Have you ever wondered about the hidden features in the Windows 95, 98 and NT registry that can improve performance, add cool features and increase security? The RegEdit.com : Windows Registry Guide has all the best tips, tricks and tweaks from the RegEdit.com web site (www.regedit.com) in a convenient Windows help file."
    End With
    pSetFontSize

End Sub

Private Sub Form_Resize()

    Dim lHeight As Long
    Dim lTop    As Long

    On Error Resume Next
    With ucCommentry
        lHeight = .Height + mlGap
        If Me.ScaleHeight > lHeight Then
            lTop = Me.ScaleHeight - lHeight
            lHeight = mlMaxHeight
        Else
            lTop = mlGap
            lHeight = Me.ScaleHeight - mlGap * 2
        End If
        .Move .Left, lTop, Me.ScaleWidth - .Left * 2, lHeight
    End With

End Sub

Private Sub pSetFontSize()
    ucCommentry.FontHeight = CLng(txtFontSize.Text)
End Sub
