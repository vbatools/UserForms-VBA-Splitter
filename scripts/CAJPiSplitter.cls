VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAJPiSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : CAJPiSplitter
'* Created    : 04-04-2021 10:34
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit

Public Enum ajpiGrabHandle
    ajpiGrabHandleDown = 54
    ajpiGrabHandleRight = 52
    ajpiGrabHandleUp = 53
    ajpiGrabHandleLeft = 51
    ajpiGrabHandleDotSmall = 104
    ajpiGrabHandleSquare = 103
    ajpiGrabHandleDotMedium = 105
    ajpiGrabHandleDotLarge = 110
    ajpiGrabHandleDash = 113
    ajpiGrabHandleCross = 114
    ajpiGrabHandleUpDown = 118
    ajpiGrabHandleDownUnderscore = 55
End Enum

Private Const mSPLITTERNAME = "Splitter"
Private Const mSPLIT_LEFTRIGHT = 1
Private Const mSPLIT_TOPBOTTOM = 2

Private m_blnSplitting As Boolean
Private m_frmParent As MSForms.UserForm
Private m_objHandle As MSForms.Label
Attribute m_objHandle.VB_VarHelpID = -1
Private m_lngLHSNItems As Long
Private m_lngRHSNItems As Long
Private m_objLHS() As Object
Private m_objRHS() As Object
Private m_lngOrientation As Long
Private m_lngUsingOrientation As fmOrientation

Private WithEvents m_objSplitter As MSForms.Label
Attribute m_objSplitter.VB_VarHelpID = -1

Public Event ResizeCompleted()
Public Sub Initialize(Parent As MSForms.UserForm)
    
    Set m_frmParent = Parent
    
End Sub
Public Sub AddControlsLeft(ParamArray Items() As Variant)

    Dim intIndex As Integer
    
    m_lngLHSNItems = UBound(Items) - LBound(Items)
    ReDim m_objLHS(m_lngLHSNItems) As Object
    For intIndex = 0 To m_lngLHSNItems
        Set m_objLHS(intIndex) = Items(intIndex)
    Next
    
    m_lngOrientation = mSPLIT_LEFTRIGHT
    
    If m_lngRHSNItems >= 0 Then
        mCreateSplitter
    End If
    
End Sub
Public Sub AddControlsRight(ParamArray Items() As Variant)

    Dim intIndex As Integer
    
    m_lngRHSNItems = UBound(Items) - LBound(Items)
    ReDim m_objRHS(m_lngRHSNItems) As Object
    For intIndex = 0 To m_lngRHSNItems
        Set m_objRHS(intIndex) = Items(intIndex)
    Next
    
    m_lngOrientation = mSPLIT_LEFTRIGHT
    
    If m_lngLHSNItems >= 0 Then
        mCreateSplitter
    End If
    
End Sub
Public Sub AddControlsAbove(ParamArray Items() As Variant)

    Dim intIndex As Integer
    
    m_lngLHSNItems = UBound(Items) - LBound(Items)
    ReDim m_objLHS(m_lngLHSNItems) As Object
    For intIndex = 0 To m_lngLHSNItems
        Set m_objLHS(intIndex) = Items(intIndex)
    Next
    
    m_lngOrientation = mSPLIT_TOPBOTTOM
    
    If m_lngRHSNItems >= 0 Then
        mCreateSplitter
    End If
    
End Sub
Public Sub AddControlsBelow(ParamArray Items() As Variant)

    Dim intIndex As Integer
    
    m_lngRHSNItems = UBound(Items) - LBound(Items)
    ReDim m_objRHS(m_lngRHSNItems) As Object
    For intIndex = 0 To m_lngRHSNItems
        Set m_objRHS(intIndex) = Items(intIndex)
    Next
    
    m_lngOrientation = mSPLIT_TOPBOTTOM
    
    If m_lngLHSNItems >= 0 Then
        mCreateSplitter
    End If
    
End Sub

Public Property Let Enabled(RHS As Boolean)
    m_objSplitter.Enabled = RHS
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_objSplitter.Enabled
End Property

Public Property Let CustomGrabHandle(RHS As String)
    With m_objHandle
        .Caption = String(5, ChrW(RHS))
        If m_lngUsingOrientation = mSPLIT_TOPBOTTOM Then
            .Width = m_objSplitter.Width
            .AutoSize = True
            .AutoSize = False
            .Height = m_objSplitter.Height
            .Left = m_objSplitter.Left + ((m_objSplitter.Width - .Width) / 2)
            .Top = m_objSplitter.Top
            .ZOrder 1
        Else
            .Width = m_objSplitter.Width
            .Height = m_objSplitter.Height
            .AutoSize = True
            .AutoSize = False
            .Top = m_objSplitter.Top + ((m_objSplitter.Height - .Height) / 2)
            .Left = m_objSplitter.Left
            .ZOrder 1
        End If
    End With
End Property
Public Property Let GrabHandle(RHS As ajpiGrabHandle)
    
    With m_objHandle
        .Caption = String(5, ChrW(RHS))
        If m_lngUsingOrientation = mSPLIT_TOPBOTTOM Then
            .Width = m_objSplitter.Width
            .AutoSize = True
            .AutoSize = False
            .Height = m_objSplitter.Height
            .Left = m_objSplitter.Left + ((m_objSplitter.Width - .Width) / 2)
            .Top = m_objSplitter.Top
            .ZOrder 1
        Else
            .Width = m_objSplitter.Width
            .Height = m_objSplitter.Height
            .AutoSize = True
            .AutoSize = False
            .Top = m_objSplitter.Top + ((m_objSplitter.Height - .Height) / 2)
            .Left = m_objSplitter.Left
            .ZOrder 1
        End If
    End With

End Property


Public Property Let GrabHandleBackcolor(RHS As Long)
    m_objHandle.BackColor = RHS
End Property
Public Property Let GrabHandleForecolor(RHS As Long)
    m_objHandle.ForeColor = RHS
End Property

Public Property Get GrabHandleBackcolor() As Long
    GrabHandleBackcolor = m_objHandle.BackColor
End Property
Public Property Get GrabHandleForecolor() As Long
    GrabHandleForecolor = m_objHandle.ForeColor
End Property

Public Property Get Height() As Single
    Height = m_objSplitter.Height
End Property

Public Property Let Left(RHS As Single)
    m_objSplitter.Left = RHS
    m_PositionGrabHandle
End Property
Public Property Let Height(RHS As Single)
    m_objSplitter.Height = RHS
    m_PositionGrabHandle
End Property

Public Property Get Left() As Single
    Left = m_objSplitter.Left
End Property
Private Sub m_GetVerticalSizes(Left As Single, Top As Single, Width As Single, Height As Single)
'
'  +-----+
'  |     |
'  | LHS |
'  |     |
'  +-----+  ^
'  .......  |
'  +-----+  v
'  |     |
'  | RHS |
'  |     |
'  +-----+
'
    Dim intIndex As Integer
    Dim sngTop As Single
    Dim sngRight As Single
    Dim sngLeft As Single
    
    sngLeft = m_objLHS(0).Left
    sngTop = m_objLHS(0).Top + m_objLHS(0).Height
    sngRight = m_objLHS(0).Left + m_objLHS(0).Width
    For intIndex = LBound(m_objLHS) To UBound(m_objLHS)
        If m_objLHS(intIndex).Left < sngLeft Then
            sngLeft = m_objLHS(intIndex).Left
        End If
        If (m_objLHS(intIndex).Top + m_objLHS(intIndex).Height) > sngTop Then
            sngTop = (m_objLHS(intIndex).Top + m_objLHS(intIndex).Height)
        End If
        If (m_objLHS(intIndex).Left + m_objLHS(intIndex).Width) > sngRight Then
            sngRight = m_objLHS(intIndex).Left + m_objLHS(intIndex).Width
        End If
    Next
    Left = sngLeft
    Top = sngTop
    Width = sngRight - Left
    If Width < 0 Then Width = 0
    
    sngTop = m_objRHS(0).Top
    For intIndex = LBound(m_objRHS) To UBound(m_objRHS)
        If m_objRHS(intIndex).Top < sngTop Then
            sngTop = m_objRHS(intIndex).Top
        End If
    Next
    Height = sngTop - Top
    If Height < 0 Then Height = 0

End Sub
Private Sub m_GetHorizontalSizes(Left As Single, Top As Single, Width As Single, Height As Single)
'         <->
'  +-----+ : +-----+
'  |     | : |     |
'  | LHS | : | RHS |
'  |     | : |     |
'  +-----+ : +-----+
'         <->
'
'
    Dim intIndex As Integer
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim sngBottom As Single
    
    sngLeft = m_objLHS(0).Left + m_objLHS(0).Width
    sngTop = m_objLHS(0).Top
    sngBottom = m_objLHS(0).Top + m_objLHS(0).Height
    For intIndex = LBound(m_objLHS) To UBound(m_objLHS)
        If (m_objLHS(intIndex).Left + m_objLHS(intIndex).Width) > sngLeft Then
            sngLeft = m_objLHS(intIndex).Left + m_objLHS(intIndex).Width
        End If
        If m_objLHS(intIndex).Top < sngTop Then
            sngTop = m_objLHS(0).Top
        End If
        If (m_objLHS(intIndex).Top + m_objLHS(intIndex).Height) > sngBottom Then
            sngBottom = m_objLHS(intIndex).Top + m_objLHS(intIndex).Height
        End If
    Next
    Left = sngLeft
    Top = sngTop
    Height = sngBottom - sngTop
    If Height < 0 Then Height = 0
    
    sngLeft = m_objRHS(0).Left
    For intIndex = LBound(m_objRHS) To UBound(m_objRHS)
        If m_objRHS(intIndex).Left < sngLeft Then
            sngLeft = m_objRHS(intIndex).Left
        End If
    Next
    Width = sngLeft - Left
    If Width < 0 Then Width = 0
    
End Sub
Private Sub m_PositionGrabHandle()
        
    If m_objHandle Is Nothing Then Exit Sub
    
    With m_objHandle
        If m_lngUsingOrientation = mSPLIT_TOPBOTTOM Then
            .Width = m_objSplitter.Width
            .AutoSize = True
            .AutoSize = False
            .Height = m_objSplitter.Height
            .Left = m_objSplitter.Left + ((m_objSplitter.Width - .Width) / 2)
            .Top = m_objSplitter.Top
            .ZOrder 1
        Else
            .Width = m_objSplitter.Width
            .Height = m_objSplitter.Height
            .AutoSize = True
            .AutoSize = False
            .Top = m_objSplitter.Top + ((m_objSplitter.Height - .Height) / 2)
            .Left = m_objSplitter.Left
            .ZOrder 1
        End If
    End With
End Sub
Public Property Let Orientation(RHS As fmOrientation)
    m_lngOrientation = RHS
End Property

Public Property Get Orientation() As fmOrientation
    Orientation = m_lngOrientation
End Property
Public Sub Remove()
    m_frmParent.Controls.Remove mSPLITTERNAME
End Sub

Private Sub mCreateSplitter()
'
' Create splitter
'
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim sngWidth As Single
    Dim sngHeight As Single
    Dim strCaption As String
    Dim lngGrabHandle As ajpiGrabHandle
    
    If m_objSplitter Is Nothing Then
        Set m_objSplitter = m_frmParent.Controls.Add("Forms.label.1", mSPLITTERNAME, True)
        Set m_objHandle = m_frmParent.Controls.Add("Forms.label.1", mSPLITTERNAME, True)
    End If
    
    If m_lngOrientation = mSPLIT_LEFTRIGHT Then
        m_GetHorizontalSizes sngLeft, sngTop, sngWidth, sngHeight
        m_objSplitter.MousePointer = fmMousePointerSizeWE
        m_lngUsingOrientation = mSPLIT_LEFTRIGHT
        lngGrabHandle = ajpiGrabHandleLeft
    Else
        m_GetVerticalSizes sngLeft, sngTop, sngWidth, sngHeight
        m_objSplitter.MousePointer = fmMousePointerSizeNS
        m_lngUsingOrientation = mSPLIT_TOPBOTTOM
        lngGrabHandle = ajpiGrabHandleDown
    End If
    
    With m_objSplitter
        If sngLeft > 0 Then .Left = sngLeft
        If sngWidth > 0 Then .Width = sngWidth
        If sngTop > 0 Then .Top = sngTop
        If sngHeight > 0 Then .Height = sngHeight
        .BackColor = m_frmParent.BackColor
        .BackStyle = fmBackStyleTransparent
    End With
    With m_objHandle
        .Font.Name = "Marlett"
        .Font.Size = 6
        .Font.Name = "Marlett"
        .Font.Charset = 1
        Me.GrabHandle = lngGrabHandle
        If m_lngUsingOrientation = mSPLIT_TOPBOTTOM Then
            .Width = m_objSplitter.Width
            .AutoSize = True
            .AutoSize = False
            .Height = m_objSplitter.Height
            .Left = m_objSplitter.Left + ((m_objSplitter.Width - .Width) / 2)
            .Top = m_objSplitter.Top
            .ZOrder 1
        Else
            .Width = m_objSplitter.Width
            .Height = m_objSplitter.Height
            .AutoSize = True
            .AutoSize = False
            .Top = m_objSplitter.Top + ((m_objSplitter.Height - .Height) / 2)
            .Left = m_objSplitter.Left
            .ZOrder 1
        End If
    End With
    
End Sub
Public Property Let Tooltip(RHS As String)
    m_objSplitter.ControlTipText = RHS
End Property


Public Property Let Top(RHS As Single)
    m_objSplitter.Top = RHS
    m_PositionGrabHandle
End Property

Public Property Get Top() As Single
    Top = m_objSplitter.Top
End Property

Public Property Let Width(RHS As Single)
    m_objSplitter.Width = RHS
    m_PositionGrabHandle
End Property
Public Property Get Width() As Single
    Width = m_objSplitter.Width
End Property
Private Sub Class_Initialize()

    m_lngLHSNItems = -1
    m_lngRHSNItems = -1
    
End Sub

Private Sub Class_Terminate()

    m_frmParent.Controls.Remove mSPLITTERNAME
    Set m_objSplitter = Nothing
    Set m_objHandle = Nothing
    Set m_frmParent = Nothing

End Sub


Private Sub m_objSplitter_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    m_blnSplitting = (Button = 1)
    
End Sub


Private Sub m_objSplitter_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    Dim intRHSIndex As Integer
    Dim intLHSIndex As Integer
    
    If m_blnSplitting Then
        If m_lngUsingOrientation = mSPLIT_LEFTRIGHT Then
            If m_objSplitter.Left + X >= 0 Then
                If m_objLHS(0).Width + X >= 0 Then
                    If m_objRHS(0).Width - X >= 0 Then
                        With m_objSplitter
                            .Move .Left + X, .Top, .Width, .Height
                            .Visible = True
                        End With
                        With m_objHandle
                            .Top = m_objSplitter.Top + ((m_objSplitter.Height - .Height) / 2)
                            .Left = m_objSplitter.Left
                            .Visible = True
                        End With
                        For intLHSIndex = 0 To m_lngLHSNItems
                            With m_objLHS(intLHSIndex)
                                .Width = .Width + X
                            End With
                        Next
                        For intRHSIndex = 0 To m_lngRHSNItems
                            With m_objRHS(intRHSIndex)
                                .Move .Left + X, .Top, .Width, .Height
                                .Width = .Width - X
                            End With
                        Next
                    End If
'                    For intRHSIndex = 1 To UBound(m_objRHS)
'                        m_objRHS(intRHSIndex).Left = m_objRHS(0).Left
'                        m_objRHS(intRHSIndex).Width = m_objRHS(intRHSIndex).Width - X
'                    Next
                End If
            End If
        Else
            If m_objSplitter.Top + Y >= 0 Then
                If m_objLHS(0).Height + Y >= 0 Then
                    If m_objRHS(0).Height - Y >= 0 Then
                        With m_objSplitter
                            .Move .Left, .Top + Y, .Width, .Height
                            .Visible = True
                        End With
                        With m_objHandle
                            .Top = m_objSplitter.Top + ((m_objSplitter.Height - .Height) / 2)
                            .Visible = True
                        End With
                        For intLHSIndex = 0 To m_lngLHSNItems
                            With m_objLHS(intLHSIndex)
                                .Height = .Height + Y
                            End With
                        Next
                        For intRHSIndex = 0 To m_lngRHSNItems
                            With m_objRHS(intRHSIndex)
                                .Move .Left, .Top + Y, .Width, .Height
                                .Height = .Height - Y
                            End With
                        Next
                    End If
'                    For intRHSIndex = 1 To UBound(m_objRHS)
'                        m_objRHS(intRHSIndex).Top = m_objRHS(0).Top
'                        m_objRHS(intRHSIndex).Height = m_objRHS(intRHSIndex).Height - Y
'                    Next
                End If
            End If
        End If
    End If

End Sub

Private Sub m_objSplitter_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    RaiseEvent ResizeCompleted
    m_blnSplitting = False
    
End Sub
