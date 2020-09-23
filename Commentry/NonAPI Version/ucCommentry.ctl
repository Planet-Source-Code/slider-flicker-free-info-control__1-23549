VERSION 5.00
Begin VB.UserControl ucCommentry 
   BackColor       =   &H8000000C&
   ClientHeight    =   1665
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5385
   ScaleHeight     =   1665
   ScaleWidth      =   5385
   Begin VB.Label lblDesc 
      BackColor       =   &H80000002&
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      ForeColor       =   &H80000014&
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   3975
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5175
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   150
      TabIndex        =   2
      Top             =   30
      Width           =   5175
   End
End
Attribute VB_Name = "ucCommentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const clSTEP As Long = 20

Private mlGap As Long

Public Property Let Caption(Text As String)
    lblCaption(0).Caption = Text
    lblCaption(1).Caption = Text
End Property

Public Property Let Description(Text As String)
    lblDesc = Text
End Property

Public Property Let FontHeight(Size As Long)
    With lblCaption(0)
        .FontSize = Size + 2
        FontSize = .FontSize
        .Height = TextHeight(.Caption)
        lblCaption(1).FontSize = .FontSize
        lblCaption(1).Height = .Height
    End With
    With lblDesc
        .FontSize = Size
    End With
    pResize
End Property

Private Sub UserControl_Initialize()
    mlGap = 2 * Screen.TwipsPerPixelX
End Sub

Private Sub UserControl_Paint()
    ThinBorder UserControl.hwnd, True
End Sub

Private Sub UserControl_Resize()
    pResize
End Sub

Private Sub pResize()

    Dim lTop As Long

    On Error Resume Next
    With lblCaption(0)
        .Move mlGap, mlGap, UserControl.ScaleWidth - 2 * mlGap, .Height
        lblCaption(1).Move .Left + clSTEP, .Top + clSTEP, .Width - clSTEP, .Height - clSTEP
        lTop = .Height
    End With
    With lblDesc
        .Move mlGap * 4, lTop, lblCaption(0).Width - mlGap * 3, UserControl.ScaleHeight - lTop - mlGap * 2
    End With

End Sub
