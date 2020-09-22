VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   6930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8700
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   462
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHTxt 
      Caption         =   "Hook text box"
      Height          =   495
      Left            =   120
      TabIndex        =   14
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdUTxt 
      Caption         =   "Unhook text box"
      Height          =   495
      Left            =   3120
      TabIndex        =   13
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdUChk 
      Caption         =   "Unhook check box"
      Height          =   495
      Left            =   3120
      TabIndex        =   11
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton cmdHChk 
      Caption         =   "Hook check box"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Form1 subclass"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.CommandButton Command3 
         Caption         =   "Remove subclass"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Get user data"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Install window subclass"
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtrefd 
         Height          =   285
         Left            =   3480
         TabIndex        =   1
         Text            =   "0"
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Try to resize the form and move it to screen edge when subclass installed."
         Height          =   435
         Left            =   5160
         TabIndex        =   15
         Top             =   1680
         Width           =   2745
      End
      Begin VB.Label lblinst 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   4200
         TabIndex        =   8
         Top             =   480
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "with user data"
         Height          =   195
         Left            =   2280
         TabIndex        =   7
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label lblclr 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   2280
         TabIndex        =   6
         Top             =   1680
         Width           =   45
      End
      Begin VB.Label lblrefd 
         AutoSize        =   -1  'True
         Caption         =   "Reference data: "
         Height          =   195
         Left            =   2280
         TabIndex        =   4
         Top             =   1080
         Width           =   1230
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Winand's SubCL Project"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   6960
      TabIndex        =   20
      Top             =   3000
      Width           =   1365
   End
   Begin VB.Label Label4 
      Caption         =   "These hooks are not removed on form_unload so the'll be removed automatically on WM_NCDESTROY message to each of them."
      Height          =   1515
      Left            =   4920
      TabIndex        =   19
      Top             =   2880
      Width           =   1785
   End
   Begin VB.Label Label3 
      Caption         =   $"Form1.frx":0000
      Height          =   1545
      Left            =   120
      TabIndex        =   18
      Top             =   4560
      Width           =   8160
   End
   Begin VB.Label lblTxtMsg 
      AutoSize        =   -1  'True
      Caption         =   "Msg:"
      Height          =   195
      Left            =   1800
      TabIndex        =   17
      Top             =   4125
      Width           =   345
   End
   Begin VB.Label lblChkMsg 
      AutoSize        =   -1  'True
      Caption         =   "Msg:"
      Height          =   195
      Left            =   1800
      TabIndex        =   16
      Top             =   3120
      Width           =   345
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Safe subclass test project 0.0.1
'Winand, 2010-11-09 (http://sourceforge.net/projects/audica/)

Option Explicit
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Const WM_GETMINMAXINFO As Long = &H24
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type MINMAXINFO
   ptReserved As POINTAPI
   ptMaxSize As POINTAPI
   ptMaxPosition As POINTAPI
   ptMinTrackSize As POINTAPI
   ptMaxTrackSize As POINTAPI
End Type
Private Const WM_WINDOWPOSCHANGING As Long = &H46
Private Type WINDOWPOS
   hWnd As Long
   hWndInsertAfter As Long
   x As Long
   y As Long
   cx As Long
   cy As Long
   Flags As Long
End Type
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uiAction As Long, ByVal uiParam As Long, pvParam As Any, ByVal fWinIni As Long) As Long
Private Const SPI_GETWORKAREA As Long = 48
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cBytes&)

Private subc As New csubcl

Private Sub cmdHChk_Click()
    Call subc.HookSet(Me, Check1.hWnd)
End Sub

Private Sub cmdHTxt_Click()
    Call subc.HookSet(Me, Text1.hWnd)
End Sub

Private Sub cmdUChk_Click()
    Call subc.HookClear(Me, Check1.hWnd)
End Sub

Private Sub cmdUTxt_Click()
    Call subc.HookClear(Me, Text1.hWnd)
End Sub

Private Sub Command1_Click()
    Dim userdata As Long, installed As Long
    installed = subc.HookGetData(Me, , userdata)
    If installed Then _
        lblrefd.Caption = "Reference data: " & userdata _
    Else lblrefd.Caption = "Hook not installed"
End Sub

Private Sub Command2_Click()
    If subc.HookSet(Me, hWnd, Val(txtrefd.Text)) Then _
        lblinst.Caption = "Hook installed (or user data updated) successfully"
End Sub

Private Sub Command3_Click()
    If subc.HookClear(Me, hWnd) Then _
        lblclr.Caption = "Hook cleared succefully" _
    Else lblclr.Caption = "Cannot clear hook"
End Sub

Private Sub Form_Load()
    Caption = "I'm Form1. My object pointer is " & ObjPtr(Me) & ", my window handle is " & hWnd
    Debug.Print "---app starts---"
    If subc.HookSet(Me, hWnd, Val(txtrefd.Text)) Then _
        lblinst.Caption = "Hook installed successfully"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If you don't clear hook manually
    'it'll be cleared automatically on WM_NCDESTROY message
    Call subc.HookClear(Me, hWnd)
End Sub

Public Function WindowProc(ByRef bHandled As Boolean, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal dwRefData As Long) As Long
    If hWnd = Me.hWnd Then
        'From Karl E. Peterson's SnapDialog http://vb.mvps.org/samples/SnapDialog/
        If uMsg = WM_GETMINMAXINFO Then
            Dim mmi As MINMAXINFO
            ' Snatch copy of current minmax.
            Call CopyMemory(mmi, ByVal lParam, Len(mmi))
            mmi.ptMinTrackSize.x = 570
            mmi.ptMinTrackSize.y = 372
            ' Make sure maximum really is maximum, to avoid flash.
            mmi.ptMaxSize = mmi.ptMaxTrackSize
            ' Send altered values back to Windows.
            Call CopyMemory(ByVal lParam, mmi, Len(mmi))
            bHandled = True 'do not call DefSubclassProc
        ElseIf uMsg = WM_WINDOWPOSCHANGING Then
            Dim pos As WINDOWPOS, mon As RECT
            ' Snag copy of position structure and process it.
            Call CopyMemory(pos, ByVal lParam, Len(pos))
            ' Get coordinates for main work area.
            Call SystemParametersInfo(SPI_GETWORKAREA, 0&, mon, 0&)
            If Abs(pos.x - mon.Left) <= 10 Then ' Snap X axis
               pos.x = mon.Left
            ElseIf Abs(pos.x + pos.cx - mon.Right) <= 10 Then
               pos.x = mon.Right - pos.cx
            End If
            If Abs(pos.y - mon.Top) <= 10 Then ' Snap Y axis
               pos.y = mon.Top
            ElseIf Abs(pos.y + pos.cy - mon.Bottom) <= 10 Then
               pos.y = mon.Bottom - pos.cy
            End If
            ' Send altered values back to Windows.
            Call CopyMemory(ByVal lParam, pos, Len(pos))
            bHandled = True 'do not call DefSubclassProc
        End If
    ElseIf hWnd = Check1.hWnd Then
        lblChkMsg.Caption = "Msg: " & Hex(uMsg)
    ElseIf hWnd = Text1.hWnd Then
        lblTxtMsg.Caption = "Msg: " & Hex(uMsg)
    End If
    'Be careful in this proc, don't freeze app with infinite messages
End Function

