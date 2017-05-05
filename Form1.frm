VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   2760
   LinkTopic       =   "Form1"
   ScaleHeight     =   2430
   ScaleWidth      =   2760
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton calcBtn 
      Caption         =   "计算"
      Height          =   375
      Left            =   500
      TabIndex        =   1
      Top             =   1700
      Width           =   1600
   End
   Begin VB.TextBox inputTxt 
      Height          =   375
      Left            =   500
      TabIndex        =   0
      Top             =   1000
      Width           =   1600
   End
   Begin VB.Label Label1 
      Caption         =   "请输入一个奇数"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   500
      TabIndex        =   2
      Top             =   300
      Width           =   2000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub calcBtn_Click()
    n = inputTxt.Text
    n2 = n * n
    showMat n, n2
End Sub

Private Sub inputTxt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub



Private Sub showMat(ByVal n As Integer, ByVal n2 As Integer)
    Dim x As Integer
    Dim y As Integer
    x = 0
    y = n / 2
    Dim mat(100, 100) As Integer
    mat(x, y) = 1
    For i = 2 To n2
        If (i - 1) Mod n = 0 Then
            nx = x + 1
            ny = y
        Else
            nx = x - 1
            ny = y + 1
        End If
        If nx < 0 Then
            nx = n - 1
        End If
        If ny >= n Then
            ny = 0
        End If
        x = nx
        y = ny
        mat(x, y) = i
    Next i
    
    Dim str As String
    
    str = ""
    
    For x = 0 To n - 1
        For y = 0 To n - 1
            str = str & mat(x, y) & Chr(9)
        Next y
        str = str & vbCrLf
    Next x
    MsgBox str
End Sub

