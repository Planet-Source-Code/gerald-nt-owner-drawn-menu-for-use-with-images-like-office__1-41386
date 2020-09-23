VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'selected menu-icons (3D)
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const BF_BOTTOM = &H8
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

'owner-drawn menus have to process these messages
Private Const WM_DRAWITEM = &H2B
Private Const WM_MEASUREITEM = &H2C

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'draw menu-item
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, _
        ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As _
        String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const SM_CXFRAME = 32
Private Const SRCAND = &H8800C6
Private Const SRCCOPY = &HCC0020
Private Const COLOR_MENU = 4
Private Const COLOR_HIGHLIGHT = 13 'Selected item background
Private Const COLOR_HIGHLIGHTTEXT = 14 'Selected item-text
Private Const ODS_SELECTED = &H1
Private Const ODS_CHECKED = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0

Private Const TRANSPARENT = 1
Private Type DRAWITEMSTRUCT
        CtlType As Long
        CtlID As Long
        itemID As Long
        itemAction As Long
        itemState As Long
        hwndItem As Long
        hdc As Long
        rcItem As RECT
        itemData As Long
End Type

'measure item
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Const DT_CALCRECT = &H400
Private Type MEASUREITEMSTRUCT
        CtlType As Long
        CtlID As Long
        itemID As Long
        itemWidth As Long
        itemHeight As Long
        itemData As Long
End Type

'create owner-drawn
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Long, lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Private Const MIIM_DATA = &H20
Private Const MF_OWNERDRAW = &H100&
Private Const MIIM_TYPE = &H10
Private Const MFT_STRING = 0
Private Const MFT_SEPARATOR = &H800
Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

Private Enum MenuType
    Main
    SubMenu
    Separator
End Enum

'structure to hold information for each menu-item
Private Type MenuItemDescription
    Caption As String
    Type As MenuType
    picBackground As PictureBox
    picChecked As PictureBox
    picUnchecked As PictureBox
    picSelect As PictureBox
    id As Long
    Items As Long 'all items of same submenu
    ItemNumber As Long 'number in submenu
    Y As Long
    TextHeight As Long 'center between top and bottom
End Type
Private m_mnuItem() As MenuItemDescription

'the form of our menu
Private m_hWnd As Long
Private m_mnuItemCounter As Integer

'leave space between picChecked and caption
'only for submenu
Private m_MnuSpaceLeft As Long

Private m_MaxPicWidth As Long
Private m_MaxPicHeight As Long
Private m_MnuSeparatorHeight As Long

'if no image is set for an item, these will be used
Private m_PicStandardMainBackground As PictureBox
Private m_PicStandardBackground As PictureBox
Private m_PicStandardSelect As PictureBox
Private m_PicStandardChecked As PictureBox
Private m_PicStandardUnChecked As PictureBox
Private m_PicStandardSeparator As PictureBox

'used for painting the area between last menu-item and
'right border of window
Private oldStruct As DRAWITEMSTRUCT
Private oldIndex As Long


'fill the array with menu-entries
Private Function addItem(ByVal sCaption As String, tMenuType As MenuType, ByVal lItems As Long, ByVal lItemNumber As Long) As Boolean
    
    If m_mnuItemCounter = 0 Then
        ReDim m_mnuItem(10) As MenuItemDescription
    ElseIf m_mnuItemCounter > UBound(m_mnuItem) Then
        ReDim Preserve m_mnuItem(m_mnuItemCounter + 10) As MenuItemDescription
    End If
    
    With m_mnuItem(m_mnuItemCounter)
        .Caption = sCaption
        .Type = tMenuType
        .id = m_mnuItemCounter
        .Items = lItems
        .ItemNumber = lItemNumber
        If .picBackground Is Nothing Then
            If .Type = Main Then
                Set .picBackground = m_PicStandardMainBackground
            ElseIf .Type = SubMenu Then
                Set .picBackground = m_PicStandardBackground
            ElseIf .Type = Separator Then
                Set .picBackground = m_PicStandardSeparator
                addItem = True
                Exit Function
            Else
                MsgBox "Unknown Type: " & .Type
                Exit Function
            End If
        End If
        
        If .picChecked Is Nothing Then
            Set .picChecked = m_PicStandardChecked
        End If
        If Not .picChecked Is Nothing Then
            If .picChecked.scaleWidth > m_MaxPicWidth Then m_MaxPicWidth = .picChecked.scaleWidth
            If .picChecked.ScaleHeight > m_MaxPicHeight Then m_MaxPicHeight = .picChecked.ScaleHeight
        End If
        
        If .picUnchecked Is Nothing Then
            Set .picUnchecked = m_PicStandardUnChecked
        End If
        If Not .picUnchecked Is Nothing Then
            If .picUnchecked.scaleWidth > m_MaxPicWidth Then m_MaxPicWidth = .picUnchecked.scaleWidth
            If .picUnchecked.ScaleHeight > m_MaxPicHeight Then m_MaxPicHeight = .picUnchecked.ScaleHeight
        End If
        
        If .picSelect Is Nothing Then
            Set .picSelect = m_PicStandardSelect
        End If
    End With

    addItem = True
End Function

'overwrite standard properties for a single item
'if you set the background picture with this function
'the whole image will be sized to dimensions of one menu-item
Public Function setItemProperties(id As Long, Optional picBackground As Variant, _
    Optional picSelect As Variant, Optional picChecked As Variant, _
    Optional picUnchecked As Variant) As Boolean
    
    If m_mnuItemCounter <= id Then
        'Init has to be called first or ID is too big
        Exit Function
    End If
    
    With m_mnuItem(id)
        
        If Not IsMissing(picBackground) Then
            .Items = 1
            .ItemNumber = 0
            setPictureProperty picBackground
            Set .picBackground = picBackground
        End If
        
        If Not IsMissing(picSelect) Then
            setPictureProperty picSelect
            Set .picSelect = picSelect
        End If
        
        If Not IsMissing(picChecked) Then
            setPictureProperty picChecked
            Set .picChecked = picChecked
            'maximum of picChecked will determine width and height
            If Not .picChecked Is Nothing Then
                If picChecked.scaleWidth > m_MaxPicWidth Then m_MaxPicWidth = picChecked.scaleWidth
                If picChecked.ScaleHeight > m_MaxPicHeight Then m_MaxPicHeight = picChecked.ScaleHeight
            End If
        End If
        
        If Not IsMissing(picUnchecked) Then
            setPictureProperty picUnchecked
            Set .picUnchecked = picUnchecked
            'maximum of picUnchecked will also determine width and height
            If Not .picUnchecked Is Nothing Then
                If picUnchecked.scaleWidth > m_MaxPicWidth Then m_MaxPicWidth = picUnchecked.scaleWidth
                If picUnchecked.ScaleHeight > m_MaxPicHeight Then m_MaxPicHeight = picUnchecked.ScaleHeight
            End If
        End If
        
    End With
        
    setItemProperties = True
    
End Function


Public Sub processMessage(ByVal hWnd As Long, ByVal lMsg As Long, _
        ByVal wParam As Long, ByVal lParam As Long)

    Dim mis As MEASUREITEMSTRUCT
    Dim dis As DRAWITEMSTRUCT

    Select Case lMsg
        Case WM_DRAWITEM
            'copy data pointed to by lParam into structure
            CopyMemory dis, ByVal lParam, Len(dis)
            OnDrawMenuItem hWnd, dis
        Case WM_MEASUREITEM
            CopyMemory mis, ByVal (lParam), LenB(mis)
            mis = OnMeasureItem(hWnd, mis)
            'copy new structure to lParam
            CopyMemory ByVal lParam, ByVal VarPtr(mis), Len(mis)
    End Select

End Sub

Public Function init(ByVal hWnd As Long, Optional picStandardMainBackground As PictureBox, Optional picStandardBackground As PictureBox, _
    Optional picStandardSelect As PictureBox, Optional picStandardChecked As PictureBox, Optional picStandardUnChecked As PictureBox, Optional picStandardSeparator As PictureBox) As Boolean
    
    Dim hMenu&
    
    If Not picStandardMainBackground Is Nothing Then
        setPictureProperty picStandardMainBackground
        Set m_PicStandardMainBackground = picStandardMainBackground
    End If
    
    If Not picStandardBackground Is Nothing Then
        setPictureProperty picStandardBackground
        Set m_PicStandardBackground = picStandardBackground
    End If
    
    If Not picStandardSelect Is Nothing Then
        setPictureProperty picStandardSelect
        Set m_PicStandardSelect = picStandardSelect
    End If
    
    If Not picStandardChecked Is Nothing Then
        setPictureProperty picStandardChecked
        Set m_PicStandardChecked = picStandardChecked
    End If
    
    If Not picStandardUnChecked Is Nothing Then
        setPictureProperty picStandardUnChecked
        Set m_PicStandardUnChecked = picStandardUnChecked
    End If
    
    If Not picStandardSeparator Is Nothing Then
        setPictureProperty picStandardSeparator
        Set m_PicStandardSeparator = picStandardSeparator
    End If
        
    m_mnuItemCounter = 0
    
    m_hWnd = hWnd
    
    hMenu = GetMenu(hWnd)
    
    createOwnerDrawnMnuItem hMenu, Main, 0, 0
    
    init = True
    
End Function

Private Function createOwnerDrawnMnuItem(ByVal hMenu As Long, ByVal tMenuType As MenuType, ByVal ItemNumber As Long, ItemCount As Long) As Boolean

    Dim mnuItemCount&
    Dim ret&
    Dim mii As MENUITEMINFO
    Dim i As Long
        
    mnuItemCount = GetMenuItemCount(hMenu)
    
    For i = 0 To mnuItemCount - 1
    
        'set these flags before calling GetMenuItemInfo
        mii.cbSize = Len(mii)
        mii.fType = MFT_STRING
        mii.fMask = MIIM_TYPE
        mii.dwTypeData = Space$(256)
        mii.cch = Len(mii.dwTypeData)
           
        ret = GetMenuItemInfo(hMenu, i, True, mii)

        mii.fType = MF_OWNERDRAW
        mii.fMask = MIIM_TYPE Or MIIM_DATA
        mii.dwItemData = m_mnuItemCounter
   
        ret = SetMenuItemInfo(hMenu, i, True, mii)

        'save into array
        If mii.dwTypeData = "" Then
            addItem "", Separator, 1, 0
        Else
            addItem getString(mii.dwTypeData), tMenuType, mnuItemCount, i
        End If
                
        m_mnuItemCounter = m_mnuItemCounter + 1
                
        createOwnerDrawnMnuItem GetSubMenu(hMenu, i), SubMenu, i, mnuItemCount
    Next i
        
End Function
 
Private Function OnMeasureItem(hWnd As Long, lpmis As MEASUREITEMSTRUCT) As MEASUREITEMSTRUCT
    Dim xM As MEASUREITEMSTRUCT
    Dim hdc As Long
    Dim ret
    Dim r As RECT
    Dim sText As String
    
    hdc = GetDC(hWnd)

    With m_mnuItem(lpmis.itemData)
        
        'calculate rect needed for writing the text
        sText = .Caption
        ret = DrawText(hdc, sText, Len(sText), r, DT_CALCRECT)
        
        .TextHeight = r.Bottom - r.Top
    
        'set width and height of the item
        xM.itemWidth = r.Right - r.Left
        
        'space for checkmarks
        If .Type = SubMenu Then
            xM.itemWidth = xM.itemWidth + m_MnuSpaceLeft + m_MaxPicWidth
        End If
        
        xM.itemHeight = (r.Bottom - r.Top)
        
        'height of menu-bar won't be changed
        If .Type = SubMenu Then
            If m_MaxPicHeight > xM.itemHeight Then xM.itemHeight = m_MaxPicHeight
        ElseIf .Type = Separator Then
            xM.itemHeight = m_MnuSeparatorHeight
        End If
                    
        OnMeasureItem = xM
        
        'clean up
        ReleaseDC hWnd, hdc

    End With

End Function

Private Sub OnDrawMenuItem(hWnd As Long, lpdis As DRAWITEMSTRUCT)

    Dim srcX&, srcY&, destX&, destY&
    Dim srcWidth&, srcHeight&, destWidth&, destHeight&
        
    destX = lpdis.rcItem.Left
    destY = lpdis.rcItem.Top

    destWidth = lpdis.rcItem.Right - destX
    destHeight = lpdis.rcItem.Bottom - destY
    
    With m_mnuItem(lpdis.itemData)
        Dim lEmptySpace As Long

        If .Type = Main Then
            'this item has changed y-coord. now we have to repaint
            'the previous item so it will fill up empty space
            If (.Y <> destY) And Not (m_mnuItem(oldIndex).picBackground Is Nothing) Then
                .Y = destY
                OnDrawMenuItem m_hWnd, oldStruct
            End If

            oldStruct = lpdis

            lEmptySpace = getEmptySpace(lpdis.itemData, lpdis.rcItem.Right)
        End If

        'BACKGROUND
        If Not .picBackground Is Nothing Then
            If (.Type = Main) Then
                srcWidth = Fix(.picBackground.scaleWidth / .Items)
                srcHeight = .picBackground.ScaleHeight
                srcX = .ItemNumber * srcWidth
                srcY = 0
            ElseIf (.Type = Separator) Then
                srcWidth = .picBackground.scaleWidth
                srcHeight = .picBackground.ScaleHeight
                srcX = 0
                srcY = 0
            Else
                srcWidth = .picBackground.scaleWidth
                srcHeight = Fix(.picBackground.ScaleHeight / .Items)
                srcX = 0
                srcY = srcHeight * .ItemNumber
            End If
            Dim ret As Long
            ret = StretchBlt(lpdis.hdc, destX, destY, destWidth + lEmptySpace, destHeight, .picBackground.hdc, srcX, srcY, srcWidth, srcHeight, SRCCOPY)
        Else
            'if no background-picture, use menu-color to clear the area
            Dim br As Long
            br = CreateSolidBrush(GetSysColor(COLOR_MENU))
            FillRect lpdis.hdc, lpdis.rcItem, br
            DeleteObject br
        End If

        'no need for further operations
        If .Type = Separator Then Exit Sub
    
        'CHECKED
        If .Type <> Main Then
            If (lpdis.itemState And ODS_CHECKED) Then
                If Not .picChecked Is Nothing Then
                    BitBlt lpdis.hdc, destX, destY, .picChecked.scaleWidth, .picChecked.ScaleHeight, .picChecked.hdc, 0, 0, SRCAND
                    destX = destX + m_MaxPicWidth
                    destWidth = destWidth - m_MaxPicWidth
                End If
            Else
                If Not .picUnchecked Is Nothing Then
                    BitBlt lpdis.hdc, destX, destY, .picUnchecked.scaleWidth, .picUnchecked.ScaleHeight, .picUnchecked.hdc, 0, 0, SRCAND
                    destX = destX + m_MaxPicWidth
                    destWidth = destWidth - m_MaxPicWidth
                End If
            End If
        End If
            
        'SELECTED
        Dim temp As RECT
        If (lpdis.itemState And ODS_SELECTED) Then
            SetTextColor lpdis.hdc, GetSysColor(COLOR_HIGHLIGHTTEXT)
            If Not .picSelect Is Nothing Then
                StretchBlt lpdis.hdc, destX, destY, destWidth, destHeight, .picSelect.hdc, 0, 0, .picSelect.scaleWidth, .picSelect.ScaleHeight, SRCAND
            Else
                'if no selection picture given, use standard color for highlighting
                Dim br2 As Long
                br2 = CreateSolidBrush(GetSysColor(COLOR_HIGHLIGHT))
                temp = lpdis.rcItem
                temp.Left = destX
                temp.Right = temp.Left + destWidth
                FillRect lpdis.hdc, temp, br2
                DeleteObject br2
            End If
            
            'there is an un/checked image
            If destX > 0 And .Type = SubMenu Then
                temp = lpdis.rcItem
                temp.Bottom = temp.Top + destHeight
                temp.Right = temp.Left + m_MaxPicWidth
                DrawEdge lpdis.hdc, temp, EDGE_RAISED, BF_RECT
            End If
        End If
            
        'CAPTION
        If Not m_MaxPicWidth Then m_MaxPicHeight = destHeight
        
        SetBkMode lpdis.hdc, TRANSPARENT
        Dim drawMethod As Long
        temp = lpdis.rcItem

        If .Type = SubMenu Then
            temp.Left = lpdis.rcItem.Left + m_MnuSpaceLeft + m_MaxPicWidth
            temp.Top = lpdis.rcItem.Top + Fix((m_MaxPicHeight - .TextHeight) / 2)
            drawMethod = DT_LEFT
        ElseIf .Type = Main Then
            temp.Top = temp.Top + Fix((temp.Bottom - temp.Top - .TextHeight) / 2)
            drawMethod = DT_CENTER
        End If

        DrawText lpdis.hdc, .Caption, Len(.Caption), temp, drawMethod
                
    End With

End Sub

'space between last menu entry of a row and the window border
Private Function getEmptySpace(lMnuItem As Long, lItemRight As Long) As Long
    
    Dim nextItem As Long

    With m_mnuItem(lMnuItem)
        If .Type <> Main Then Exit Function

        nextItem = getNextMainMenu(lMnuItem)
        
        'menu bar has more than one line -> next line has different y-coordinate or
        'last item -> nextitem = 0
        If (nextItem <> 0) And (.Y = m_mnuItem(nextItem).Y) Then Exit Function
    End With
        
    'now get window-width without right border
    Dim border As Long
    Dim r As RECT
    Dim scaleWidth As Long
    
    border = GetSystemMetrics(SM_CXFRAME)
    GetWindowRect m_hWnd, r

    scaleWidth = r.Right - r.Left - border
    getEmptySpace = scaleWidth - lItemRight
   
End Function

'seach array for the next item of menu-bar
Private Function getNextMainMenu(lIndex As Long) As Long
    
    Dim i As Long
    
    For i = lIndex + 1 To m_mnuItemCounter - 1
        If m_mnuItem(i).Type = Main Then
            getNextMainMenu = i
            Exit For
        End If
    Next i
    
End Function

Private Function getString(ByVal s As String) As String
    Dim pos&
    pos = InStr(1, s, vbNullChar) - 1
    If pos > 0 Then getString = Left$(s, pos)
End Function

Private Sub Class_Initialize()
    m_MnuSpaceLeft = 10
    m_MnuSeparatorHeight = 1
End Sub

Public Property Let SeparatorHeight(lNewVal As Long)
    m_MnuSeparatorHeight = lNewVal
End Property

Public Property Get SeparatorHeight() As Long
    SeparatorHeight = m_MnuSeparatorHeight
End Property

Public Property Let SpaceLeft(lNewVal As Long)
    m_MnuSpaceLeft = lNewVal
End Property

Public Property Get SpaceLeft() As Long
    SpaceLeft = m_MnuSpaceLeft
End Property

'otherwise blitting operation won't work correctly
Private Function setPictureProperty(ByVal pic As PictureBox) As Boolean
    
    If pic Is Nothing Then Exit Function
    
    pic.AutoRedraw = True
    pic.AutoSize = True
    pic.ScaleMode = vbPixels

    setPictureProperty = True
    
End Function
