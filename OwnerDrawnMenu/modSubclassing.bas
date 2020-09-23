Attribute VB_Name = "modSubclassing"
'This code is from Matthew Curland (slightly modified)

'Read his articles from February 1997 and August 2001 he pointed me at.
'You'll find here: http://www.Fawcette.com/Archives/Magazines/VSM


Option Explicit

Public Const GWL_WNDPROC As Long = -4
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal wndrpcPrev As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Type ThunkBytes
    Thunk(5) As Long
End Type

Private Type PushParamThunk
    pfn As Long
    Code As ThunkBytes
End Type

Private m_Thunk As PushParamThunk
Private m_lOldProcAddr As Long
Private m_hWnd As Long


'function generated dynamically
Public Sub InitPushParamThunk(ByVal pObject As Long, ByVal lFunctionAddress As Long)
'push [esp]
'mov eax, 16h // Dummy value for parameter value
'mov [esp + 4], eax
'nop // Adjustment so the next long is nicely aligned
'nop
'nop
'mov eax, 1234h // Dummy value for function
'jmp eax
'nop
'nop
    
    With m_Thunk.Code
        .Thunk(0) = &HB82434FF
        .Thunk(1) = pObject
        .Thunk(2) = &H4244489
        .Thunk(3) = &HB8909090
        .Thunk(4) = lFunctionAddress
        .Thunk(5) = &H9090E0FF
    End With
    m_Thunk.pfn = VarPtr(m_Thunk.Code)
End Sub

Public Sub SubClass(ByVal hWnd As Long, ByVal pObject As Long, ByVal lFunctionAddress As Long)
    
    m_hWnd = hWnd
    
    InitPushParamThunk pObject, lFunctionAddress
        
    m_lOldProcAddr = SetWindowLong(hWnd, GWL_WNDPROC, m_Thunk.pfn)
    
End Sub

'due to improved subclassing-technique we already have our form-object
'no need for searching in an array or casting
Public Function RedirectWndProc(ByVal This As Form1, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    RedirectWndProc = This.WndProc(hWnd, uMsg, wParam, lParam)
    RedirectWndProc = CallWindowProc(m_lOldProcAddr, hWnd, uMsg, wParam, lParam)
End Function

Public Sub UnSubClass()
    SetWindowLong m_hWnd, GWL_WNDPROC, m_lOldProcAddr
End Sub
