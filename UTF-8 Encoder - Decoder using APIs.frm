VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form Form1 
   Caption         =   "UTF-8 Encoder/Decoder"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   7950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   855
      Left            =   6480
      TabIndex        =   16
      Top             =   4440
      Width           =   1335
      Begin VB.CommandButton Command6 
         Caption         =   "Exit"
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   4440
      Width           =   1695
      Begin VB.CommandButton Command5 
         Caption         =   "Clear All Boxes"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Standard VB Controls"
      Height          =   2175
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   7695
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2535
         TabIndex        =   0
         Top             =   360
         Width           =   4920
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   1080
         Width           =   5640
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   1560
         Width           =   5640
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Show Decoded"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Show Encoded"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   1800
         TabIndex        =   18
         Top             =   840
         Width           =   5535
      End
      Begin VB.Label Label1 
         Caption         =   "Paste In Characters To Encode"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Enter your Unicode characters here, they will be encoded using the specified codepage"
         Top             =   405
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Unicode Controls (Forms 2.0)"
      Height          =   2175
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   7695
      Begin VB.CommandButton Command2 
         Caption         =   "Show Decoded"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Show Encoded"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Paste In Characters To Encode"
         Height          =   255
         Index           =   1
         Left            =   135
         TabIndex        =   11
         ToolTipText     =   "Enter your Unicode characters here, they will be encoded using the specified codepage"
         Top             =   435
         Width           =   2295
      End
      Begin MSForms.TextBox TextBox1 
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   360
         Width           =   4935
         VariousPropertyBits=   746604571
         Size            =   "8705;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox TextBox2 
         Height          =   375
         Left            =   1920
         TabIndex        =   7
         Top             =   1080
         Width           =   5535
         VariousPropertyBits=   746604571
         Size            =   "9763;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox TextBox3 
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         Top             =   1560
         Width           =   5535
         VariousPropertyBits=   746604571
         Size            =   "9763;661"
         FontName        =   "Tahoma"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' You can use this source code as you please, all I ask is that you respect the author's
' *very few* copyright restrictions of the basInCodePage Module as that is not my work.

Option Explicit
Public sCodePage As Long
Public cnvUni2 As String
Public cnvUni As String
Private Sub Form_Load()
' Simple UTF-8 Encoder/Decoder using API

' Use character sets such as Japanese on the Forms 2.0 controls as only these are unicode enabled

' The actual functions used from the API's are MultiByteToWideChar and WideCharToMultiByte

' For a complete list of Code Pages you can use look in the Module basInCodePage !
' Example... If you want to use UTF-7 Encoding/Decoding you would set sCodePage = CP_UTF7

sCodePage = CP_UTF8
End Sub
Function EncodeUTF8(ByVal cnvUni As String)
    If cnvUni = vbNullString Then Exit Function
    EncodeUTF8 = StrConv(WToA(cnvUni, sCodePage, 0), vbUnicode)
End Function
Function DecodeUTF8(ByVal cnvUni As String)
    If cnvUni = vbNullString Then Exit Function
    cnvUni2 = WToA(cnvUni, CP_ACP)
    DecodeUTF8 = AToW(cnvUni2, sCodePage)
End Function
Private Sub Command1_Click()
    TextBox2.Text = EncodeUTF8(TextBox1.Text)
End Sub
Private Sub Command2_Click()
    TextBox3.Text = DecodeUTF8(TextBox2.Text)
End Sub
Private Sub Command3_Click()
' Hint... If you paste in �Ќ then click Show Encoded it will encode to å©ÐŒ
    Text2.Text = EncodeUTF8(Text1.Text)
End Sub
Private Sub Command4_Click()
' Hint... If you paste in å©ÐŒ into the box next to 'Show Encoded' and click
' Show Decoded it will decode the characters and display �Ќ
    Text3.Text = DecodeUTF8(Text2.Text)
End Sub
Private Sub Command5_Click()
    Text1.Text = vbNullString
    Text2.Text = vbNullString
    Text3.Text = vbNullString
    TextBox1.Text = vbNullString
    TextBox2.Text = vbNullString
    TextBox3.Text = vbNullString
End Sub
Private Sub Command6_Click()
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Command6_Click
End Sub
