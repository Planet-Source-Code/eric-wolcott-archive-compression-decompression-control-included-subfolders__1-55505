VERSION 5.00
Begin VB.UserControl CheckBox 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1320
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1305
   ScaleHeight     =   1320
   ScaleWidth      =   1305
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   0
      Width           =   315
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CheckBox"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   330
      TabIndex        =   1
      Top             =   30
      Width           =   4455
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   930
      Picture         =   "UserControl3.ctx":0000
      Top             =   1035
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   600
      Picture         =   "UserControl3.ctx":0342
      Top             =   1035
      Width           =   240
   End
End
Attribute VB_Name = "CheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Event ButtonClicked(Checked As Boolean)
Private Control_Enabled As Boolean
Private Control_Checked As Boolean

Private Sub Picture1_Click()
    If Control_Enabled = True Then
    If Control_Checked = True Then
        Control_Checked = False
        Picture1.Picture = Image2.Picture
    Else
        Control_Checked = True
        Picture1.Picture = Image1.Picture
    End If
    RaiseEvent ButtonClicked(Control_Checked)
    End If
End Sub

Private Sub UserControl_Initialize()
    Picture1.Picture = Image2.Picture
    Label1.Caption = UserControl.Name
End Sub

Public Property Get Checked() As Boolean
    Checked = Control_Checked
End Property

Public Property Let Checked(Checked As Boolean)
    If Checked = True Then
        Picture1.Picture = Image1.Picture
    Else
        Picture1.Picture = Image2.Picture
    End If
    Control_Checked = Checked
End Property

Public Property Get Enabled() As Boolean
    Enabled = Control_Enabled
End Property

Public Property Let Enabled(Enabled As Boolean)
    Control_Enabled = Enabled
End Property

Public Property Get Caption() As String
    Caption = Label1.Caption
End Property

Public Property Let Caption(Caption As String)
    Label1.Caption = Caption
End Property

Public Property Get Font() As String
    Font = Label1.Font
End Property

Public Property Let Font(Font As String)
    Label1.Font = Font
End Property

Public Property Get Bold() As Boolean
    Bold = Label1.FontBold
End Property

Public Property Let Bold(Bold As Boolean)
    Label1.FontBold = Bold
End Property

Public Property Get ForeColor() As Long
    ForeColor = Label1.ForeColor
End Property

Public Property Let ForeColor(ForeColor As Long)
    Label1.ForeColor = ForeColor
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Label1.Caption = PropBag.ReadProperty("Caption", UserControl.Name)
    Let Checked = PropBag.ReadProperty("Checked", False)
    Let Enabled = PropBag.ReadProperty("Enabled", True)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", Caption, Nothing)
    Call PropBag.WriteProperty("Checked", Control_Checked, False)
    Call PropBag.WriteProperty("Enabled", Control_Enabled, True)
End Sub
