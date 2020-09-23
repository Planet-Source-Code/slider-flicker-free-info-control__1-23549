VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Test 'Commentry' User Control"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   ScaleHeight     =   3540
   ScaleWidth      =   9510
   StartUpPosition =   1  'CenterOwner
   Begin TestCommentry.ucCommentry ucCommentry 
      Height          =   3255
      Left            =   2880
      TabIndex        =   11
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   5741
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackgroundPicture=   "Form1.frx":0000
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   405
      Left            =   1680
      TabIndex        =   10
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "C&ommentry:"
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   2655
      Begin VB.TextBox txtFont 
         Height          =   285
         Index           =   2
         Left            =   840
         TabIndex        =   7
         Text            =   "Comic Sans MS"
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtFont 
         Height          =   285
         Index           =   3
         Left            =   840
         TabIndex        =   9
         Text            =   "12"
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblFont 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblFont 
         Alignment       =   1  'Right Justify
         Caption         =   "Size:"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   735
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Caption:"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.TextBox txtFont 
         Height          =   285
         Index           =   0
         Left            =   840
         TabIndex        =   2
         Text            =   "Arial"
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtFont 
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   4
         Text            =   "24"
         Top             =   705
         Width           =   375
      End
      Begin VB.Label lblFont 
         Alignment       =   1  'Right Justify
         Caption         =   "Name:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblFont 
         Alignment       =   1  'Right Justify
         Caption         =   "Size:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   495
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlGap       As Long

Private Sub cmdApply_Click()
    pSetFontSize
End Sub

Private Sub Form_Load()
    mlGap = 4 * Screen.TwipsPerPixelX
    With ucCommentry
        .Caption = "Hidden great features"
        .Description = "Have you ever wondered about the hidden features in the Windows 95, 98 and NT registry that can improve performance, add cool features and increase security? The RegEdit.com : Windows Registry Guide has all the best tips, tricks and tweaks from the RegEdit.com web site (www.regedit.com) in a convenient Windows help file."
    End With
    pSetFontSize
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With ucCommentry
        .Move .Left, _
              mlGap, _
              Me.ScaleWidth - .Left - mlGap, _
              Me.ScaleHeight - mlGap * 2
    End With
End Sub

Private Sub pSetFontSize()
    With ucCommentry
        With .CaptionFont
            .Name = txtFont(0).Text         '!! Acidic - Nice...
            .Size = CLng(txtFont(1).Text)
            .Bold = True
        End With
        With .DescFont
            .Name = txtFont(2).Text
            .Size = CLng(txtFont(3).Text)
            .Bold = False
        End With
        .Refresh
    End With
End Sub
