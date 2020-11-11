VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Form1"
   ClientHeight    =   5616
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7452
   LinkTopic       =   "Form1"
   ScaleHeight     =   5616
   ScaleWidth      =   7452
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtExpr 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   852
      Left            =   2268
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   3612
      Width           =   2952
   End
   Begin Project1.TixyLand TixyLand1 
      Height          =   2952
      Left            =   2268
      Top             =   504
      Width           =   2952
      _extentx        =   5207
      _extenty        =   5207
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' VB6 TixyLand Control (c) 2020 by wqweto@gmail.com
'
' Based on the original idea of https://tixy.land
'
'=========================================================================
Option Explicit
DefObj A-Z
Private Const STR_MODULE_NAME As String = "Form1"

Private Const STR_EXPR As String = "sin(y/8+t)|~|random() < 0.1|~|random()|~|sin(t)|~|i / 256|~|x / 16|~|y / 16|~|y - 7.5|~|y - t|~|y - t*4|~|[1, 0, -1][i%3]|~|sin(t-sqrt(pow(x-7.5,2)+pow(y-6,2)))|~|sin(y/8 + t)|~|" & _
                                   "y - x|~|(y > x) && (14-x < y)|~|i%4 - y%4|~|x%4 && y%4|~|x>3 & y>3 & x<12 & y<12|~|-(x>t & y>t & x<15-t & y<15-t)|~|(y-6) * (x-6)|~|(y-4*t|0) * (x-2-t|0)|~|" & _
                                   "4 * t & i & x & y|~|(t*10) & (1<<x) && y==8|~|random() * 2 - 1|~|sin(pow(i, 2))|~|cos(t + i + x * y)|~|sin(x/2) - sin(x-t) - y+6|~|sin(t-sqrt(x*x+y*y))|~|(x-8)*(y-8) - sin(t)*64|~|" & _
                                   "d=y*y%5.9+1,!((x+t*50/d)&15)/d|~|y == x || -(15-x == y)|~|x==0 | x==15 | y==0 | y==15|~|8*t%13 - hypot(x-7.5, y-7.5)|~|sin(PI*2*atan((y-8)/(x-8))+5*t)|~|" & _
                                   "(x-y) - sin(t) * 16|~|(x-y)/24 - sin(t)|~|sin(t*5) * tan(t*7)|~|pow(x-5,2) + pow(y-5,2) - 99*sin(t)|~|YOUR EXPRESSION HERE!"
Private m_lCurrent      As Long
Private m_bInSet        As Boolean

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunction As String)
    Debug.Print "Critical error: " & Err.Description & " [" & STR_MODULE_NAME & "." & sFunction & "]", Timer
End Sub

Private Sub Form_Load()
    Const FUNC_NAME     As String = "Form_Load"
    
    On Error GoTo EH
    m_lCurrent = -1
    TixyLand1_Click
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub Form_Resize()
    Const FUNC_NAME     As String = "Form_Resize"
    
    On Error GoTo EH
    If WindowState <> vbMinimized Then
        TixyLand1.Left = (ScaleWidth - TixyLand1.Width) / 2
        txtExpr.Left = (ScaleWidth - txtExpr.Width) / 2
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub TixyLand1_Click()
    Const FUNC_NAME     As String = "TixyLand1_Click"
    Dim vSplit          As Variant
    
    On Error GoTo EH
    vSplit = Split(STR_EXPR, "|~|")
    m_lCurrent = (m_lCurrent + 1) Mod (UBound(vSplit) + 1)
    Caption = "TixyLand"
    TixyLand1.Expression = vSplit(m_lCurrent)
    m_bInSet = True
    txtExpr.Text = vSplit(m_lCurrent)
    txtExpr.SelLength = &H7FFF
    m_bInSet = False
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub TixyLand1_ScriptError()
    Const FUNC_NAME     As String = "TixyLand1_ScriptError"
    
    On Error GoTo EH
    Caption = TixyLand1.LastError
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub

Private Sub txtExpr_Change()
    On Error GoTo QH
    If Not m_bInSet Then
        Caption = "TixyLand"
        TixyLand1.Expression = txtExpr.Text
    End If
QH:
End Sub
