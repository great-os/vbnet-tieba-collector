VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMain 
   Caption         =   "Baidu Tieba Collector"
   ClientHeight    =   5445
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   13725
   LinkTopic       =   "Form1"
   ScaleHeight     =   363
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   915
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      Caption         =   "搜索结果"
      Height          =   5250
      Left            =   4560
      TabIndex        =   7
      Top             =   120
      Width           =   9135
      Begin MSComctlLib.ListView lstPosts 
         Height          =   4880
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   8599
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "总控台"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton btnStop 
         Caption         =   "停止"
         Height          =   495
         Left            =   1800
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton btnBegin 
         Caption         =   "开始采集"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox txtPages 
         Height          =   270
         Left            =   960
         TabIndex        =   4
         Top             =   570
         Width           =   1215
      End
      Begin VB.TextBox txtBarName 
         Height          =   270
         Left            =   600
         TabIndex        =   2
         Top             =   240
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "搜索页数"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "贴吧"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Width           =   375
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'http://tieba.baidu.com/f?kw=vb&ie=utf-8&pn=50
Private Sub btnBegin_Click()
    Dim barName As String
    Dim pageIndex As Integer
    
    barName = txtBarName.Text
    pageIndex = CInt(txtPages.Text)
    
    Dim Tb As New TiebaCollector
    Tb.s
End Sub
