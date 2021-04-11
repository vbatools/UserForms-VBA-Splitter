VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FSplitter 
   Caption         =   "Splitter Control"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7890
   OleObjectBlob   =   "FSplitter.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : FSplitter
'* Created    : 04-04-2021 10:34
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

Private WithEvents m_clsSplitterH As CAJPiSplitter
Attribute m_clsSplitterH.VB_VarHelpID = -1
Private WithEvents m_clsSplitterV1 As CAJPiSplitter
Attribute m_clsSplitterV1.VB_VarHelpID = -1
Private WithEvents m_clsSplitterV2 As CAJPiSplitter
Attribute m_clsSplitterV2.VB_VarHelpID = -1

Private Sub CommandButton1_Click()
    Set m_clsSplitterV1 = Nothing
    Set m_clsSplitterV2 = Nothing
    Set m_clsSplitterH = Nothing
    Unload Me
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Select Case ListBox1.ListIndex
    Case 0
        m_clsSplitterV1.GrabHandle = ajpiGrabHandleDown
        m_clsSplitterH.GrabHandle = ajpiGrabHandleDown
    Case 1
        m_clsSplitterV1.GrabHandle = ajpiGrabHandleRight
        m_clsSplitterH.GrabHandle = ajpiGrabHandleRight
    Case 2
        m_clsSplitterV1.GrabHandle = ajpiGrabHandleUp
        m_clsSplitterH.GrabHandle = ajpiGrabHandleUp
    Case 3
        m_clsSplitterV1.GrabHandle = ajpiGrabHandleLeft
        m_clsSplitterH.GrabHandle = ajpiGrabHandleLeft
    Case 4
        m_clsSplitterV1.GrabHandle = ajpiGrabHandleDotSmall
        m_clsSplitterH.GrabHandle = ajpiGrabHandleDotSmall
    Case 5
        m_clsSplitterV1.GrabHandle = ajpiGrabHandleSquare
        m_clsSplitterH.GrabHandle = ajpiGrabHandleSquare
    Case 6
        m_clsSplitterV1.GrabHandle = ajpiGrabHandleDotMedium
        m_clsSplitterH.GrabHandle = ajpiGrabHandleDotMedium
    Case 7
        m_clsSplitterV1.GrabHandle = ajpiGrabHandleDotLarge
        m_clsSplitterH.GrabHandle = ajpiGrabHandleDotLarge
    Case 8
        m_clsSplitterV1.GrabHandle = ajpiGrabHandleDash
        m_clsSplitterH.GrabHandle = ajpiGrabHandleDash
    Case 9
        m_clsSplitterV1.GrabHandle = ajpiGrabHandleCross
        m_clsSplitterH.GrabHandle = ajpiGrabHandleCross
    Case 10
        m_clsSplitterV1.GrabHandle = ajpiGrabHandleUpDown
        m_clsSplitterH.GrabHandle = ajpiGrabHandleUpDown
    Case 11
        m_clsSplitterV1.GrabHandle = ajpiGrabHandleDownUnderscore
        m_clsSplitterH.GrabHandle = ajpiGrabHandleDownUnderscore
    End Select
End Sub


Private Sub m_clsSplitterH_ResizeCompleted()

    ' handle resizing of splitter between listbox2 and listbox3
    With m_clsSplitterV1
        .Width = .Width + (.Left - (m_clsSplitterH.Left + m_clsSplitterH.Width))
        .Left = m_clsSplitterH.Left + m_clsSplitterH.Width
    End With
    
    ' handle resizing of splitter between listbox1+listbox4 and listbox2+listbox3
    With m_clsSplitterV2
        .Width = m_clsSplitterH.Left - .Left
    End With
    
End Sub

Private Sub UserForm_Initialize()

    With ListBox1
        .Clear
        .AddItem "ajpiGrabHandleDown = 54"
        .AddItem "ajpiGrabHandleRight = 52"
        .AddItem "ajpiGrabHandleUp = 53"
        .AddItem "ajpiGrabHandleLeft = 51"
        .AddItem "ajpiGrabHandleDotSmall = 105"
        .AddItem "ajpiGrabHandleSquare = 103"
        .AddItem "ajpiGrabHandleDotMedium = 104"
        .AddItem "ajpiGrabHandleDotLarge = 110"
        .AddItem "ajpiGrabHandleDash = 113"
        .AddItem "ajpiGrabHandleCross = 114"
        .AddItem "ajpiGrabHandleUpDown = 118"
        .AddItem "ajpiGrabHandleDownUnderscore = 55"
    End With
    ListBox2.List = Array("One", "Two", "Three")
    ListBox3.List = ListBox2.List
    ListBox4.List = ListBox2.List
    
    ListBox3.Top = ListBox2.Top + ListBox2.Height + 7
    ListBox4.Top = ListBox1.Top + ListBox1.Height + 7
    
    Set m_clsSplitterH = New CAJPiSplitter
    m_clsSplitterH.Initialize Me
    m_clsSplitterH.AddControlsLeft ListBox1, ListBox4
    m_clsSplitterH.AddControlsRight ListBox2, ListBox3
    m_clsSplitterH.GrabHandleBackcolor = RGB(0, 255, 0)
    
    Set m_clsSplitterV1 = New CAJPiSplitter
    m_clsSplitterV1.Initialize Me
    m_clsSplitterV1.AddControlsAbove ListBox2
    m_clsSplitterV1.AddControlsBelow ListBox3
    
    With m_clsSplitterV1
        .GrabHandleForecolor = RGB(0, 0, 255)
        .Tooltip = "Изменение вертикального разделения между списками"
    End With
    
    Set m_clsSplitterV2 = New CAJPiSplitter
    m_clsSplitterV2.Initialize Me
    m_clsSplitterV2.AddControlsAbove ListBox1
    m_clsSplitterV2.AddControlsBelow ListBox4
    
    Label1.Caption = "Используйте мышь, чтобы перетащить область между списками."
End Sub
