VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form splash 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Nfl Qb Efficiency Rating Calculator."
   ClientHeight    =   4170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5490
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   5490
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   0
      Top             =   2280
   End
   Begin MSComctlLib.ProgressBar ProgLoad 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Freeware"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   $"Form.frx":0000
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   3255
   End
   Begin VB.Image Image2 
      Height          =   1200
      Left            =   3480
      Picture         =   "Form.frx":00A6
      ToolTipText     =   "www.Nfl.com"
      Top             =   840
      Width           =   1905
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "Form.frx":0893
      Top             =   120
      Width           =   720
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Loading...please stand by"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "De_killer_bee@hotmail.com"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   2760
      Width           =   3975
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'All form ontop stuff :)
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private Declare Sub SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Sub Form_Activate()
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub

Private Sub Form_Load()
    
    'Centers the form.
    Left = (Screen.Width - Width) \ 2
    Top = (Screen.Height - Height) \ 2

End Sub

Private Sub Timer1_Timer()
    
    ProgLoad.Value = ProgLoad.Value + 5
    'If the Progress Bar (ProgLoad) is 100% then your function happens.
    If ProgLoad.Value = 100 Then
        
        'Your function, can be anything. Open another form, frmMain.show... Ect.
             Form1.Show
             
        'Unloads this form
     
       Unload Me
  End If

End Sub

