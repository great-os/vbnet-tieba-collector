VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TiebaCollector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mPageIndex As Integer
Private mBarName As String
Private mWeb As New WebCode

Public Type TieBaData
    Dim Title As String
    Dim url As String
    Dim Poster As String
End Type


Private Sub Class_Initialize()

End Sub

Public Sub SetCollector(ByVal BarName As String)
    mBarName = BarName
End Sub

Public Sub MoveToPage(ByVal PageIndex As Integer)
    mPageIndex = PageIndex
End Sub

Public Function GetPageData() As TieBaData()
    Dim htmlCode As String
    Dim url As String
    
    url = mBuildUrl(mBarName, mPageIndex)
    htmlCode = mWeb.GetHttpCode(url, "smart")
    Debug.Print htmlCode
End Function

Private Function mBuildUrl(ByVal BarName As String, ByVal PageIndex As Integer) As String
    Dim url As String
    PageIndex = IIf(PageIndex < 1, 1, PageIndex)
    url = "http://tieba.baidu.com/f?kw=" & BarName & "&ie=utf-8&pn=" & (PageIndex - 1) * 50
    mBuildUrl = url
End Function
