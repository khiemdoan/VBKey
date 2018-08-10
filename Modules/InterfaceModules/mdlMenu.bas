Attribute VB_Name = "mdlMenu"
Option Explicit
' ---------- KHAI BÁO HÀM WIN API ---------------

Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoW" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Long, lpMenuItemInfo As MENUITEMINFOW) As Long
Private Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoW" (ByVal hMenu As Long, ByVal un As Long, ByVal BOOL As Boolean, lpcMenuItemInfo As MENUITEMINFOW) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuW" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long


' --------------- KHAI BÁO KIÊ?U ------------------

Private Type MENUITEMINFOW
   cbSize           As Long
   fMask            As Long
   fType            As Long
   fState           As Long
   wID              As Long
   hSubMenu         As Long
   hbmpChecked      As Long
   hbmpUnchecked    As Long
   dwItemData       As Long
   dwTypeData       As Long
   cch              As Long
   hbmpItem         As Long
End Type

' --------------KHAI BÁO CÁC HA*`NG ------------------

Private Const MIIM_TYPE = &H10
Private Const MIIM_DATA = &H20
Private Const MF_UNCHECKED = &H0&
Private Const MF_CHECKED = &H8&
Private Const MF_BYPOSITION = &H400&
Private Const MF_STRING = &H0&



' ============== THÂN CHU*O*NG TRÌNH =================

Public Sub InitMenuTV(hwnd As Long)
    Dim hMenu&
    hMenu = GetMenu(hwnd)   'lay handle cua thanh menu trong cua so ung dung
    VietnameseMenu hMenu    'lay tung menu con trong thanh menu cua ung dung
End Sub

Private Sub VietnameseMenu(ByVal hMenu As Long)
    Dim hSubMenu&, I%, nCnt%, sTmp$, sStr$
    Dim MII As MENUITEMINFOW
    
    sStr = String(&HFF, 0)
    nCnt = GetMenuItemCount(hMenu)  'dem so menu con trong thanh menu cua cua so ung dung
    If nCnt Then
        For I = 0 To nCnt - 1
            MII.cbSize = LenB(MII)
            MII.fMask = MIIM_TYPE Or MIIM_DATA
            MII.dwTypeData = StrPtr(sStr)  ' String(&HFF, 0)
            MII.cch = Len(sStr)  'MII.dwTypeData)
            MII.hbmpChecked = MF_CHECKED Or MF_UNCHECKED
            
'lay caption cua Menu
            GetMenuItemInfo hMenu, I, True, MII
            sTmp = Left$(sStr, MII.cch)  ' MII.dwTypeData, MII.cch)
            
            If sTmp <> "" Then
                sTmp = TV((sTmp))
                SetUniMenu sTmp, hMenu, I   'thuc hien gan caption cho menu vua tim duoc
            End If
            
'lay Menu con cua mot MenuItem
            hSubMenu = GetSubMenu(hMenu, I)     'lay handle menu con cua menu hien dang xu ly
            If hSubMenu Then    'neu tim thay handle thi goi de quy de xu ly caption cac menu item con trong menu cha dang xu ly
                VietnameseMenu hSubMenu
            End If
        Next
    End If
End Sub

Private Sub SetUniMenu(sCaption As String, MnuHwnd As Long, ByVal mnuItem As Long, Optional ByVal mnuParentItem As Long = -1, Optional isDefault As Boolean = False)
    Dim hMenu As Long
    Dim mInfo As MENUITEMINFOW
    
    If isDefault Then SetMenuDefaultItem hMenu, mnuItem, 1
    With mInfo
        .cbSize = Len(mInfo)
        .fType = &H200
        .fMask = &H10
        .dwTypeData = StrPtr(sCaption)
    End With
    SetMenuItemInfo MnuHwnd, mnuItem, 1, mInfo
End Sub

Public Sub SetIconForMenu(hwnd As Long, MenuIndex As Integer, SubMenuIndex1 As Integer, Optional SubMenuIndex2 As Integer, Optional SubMenuIndex3 As Integer, Optional Icon As Picture, Optional isDefault As Boolean)
On Error GoTo Err
    Dim hMainMenu As Long, hSubMenu1 As Long, hSubMenu2 As Long, hSubMenu3 As Long
    
    MenuIndex = MenuIndex - 1
    SubMenuIndex1 = SubMenuIndex1 - 1
    SubMenuIndex2 = SubMenuIndex2 - 1
    SubMenuIndex3 = SubMenuIndex3 - 1
    
    hMainMenu = GetMenu(hwnd)       'lay handle menu cua form
    
    
    If SubMenuIndex1 >= 0 Then hSubMenu1 = GetSubMenu(hMainMenu, MenuIndex)           'lay menu con thu 1
    If SubMenuIndex2 >= 0 Then hSubMenu2 = GetSubMenu(hSubMenu1, SubMenuIndex1) 'lay menu con thu 2
    If SubMenuIndex3 >= 0 Then hSubMenu3 = GetSubMenu(hSubMenu2, SubMenuIndex2) 'lay menu con thu 3
    
'neu chon Icon cho mot Menu khong ton tai trong Menu hien tai thi thoat khoi thuc tuc
    If (hSubMenu3 = 0 And SubMenuIndex3 >= 0) Or (hSubMenu2 = 0 And SubMenuIndex2 >= 0) Or (hSubMenu1 = 0 And SubMenuIndex1 >= 0) Then Exit Sub
    
'neu chon dat Icon cho menu con cap 3
    If hSubMenu3 <> 0 Then
        If isDefault Then SetMenuDefaultItem hSubMenu3, SubMenuIndex3, 1
        SetMenuItemBitmaps hSubMenu3, SubMenuIndex3, MF_BYPOSITION, Icon, Icon
        Exit Sub
    End If

'neu chon dat Icon cho menu con cap 2
    If hSubMenu2 <> 0 Then
        If isDefault Then SetMenuDefaultItem hSubMenu2, SubMenuIndex2, 1
        SetMenuItemBitmaps hSubMenu2, SubMenuIndex2, MF_BYPOSITION, Icon, Icon
        Exit Sub
    End If
    
'neu chon dat Icon cho menu con cap 1
    If hSubMenu1 <> 0 Then
        If isDefault Then SetMenuDefaultItem hSubMenu1, SubMenuIndex1, 1
        SetMenuItemBitmaps hSubMenu1, SubMenuIndex1, MF_BYPOSITION, Icon, Icon
        Exit Sub
    End If

Err:
'loi xay ra khi chon menu can dat Icon ma khong dat icon
End Sub



Public Sub UniSystemMenu(ByVal hwnd As Long, Optional mString1 As String = "Kho6i phu5c", Optional mString2 As String = "Di chuye63n", Optional mString3 As String = "D9i5nh co74", Optional mString4 As String = "Thu cu75c tie63u", Optional mString5 As String = "Pho1ng cu75c d9a5i", Optional mString6 As String = "D9o1ng          Alt + F4")
    If hwnd <= 0 Then Exit Sub
    Dim hMenu As Long, MenuItem As Long
    hMenu = GetSystemMenu(hwnd, 0)

    MenuItem = GetMenuItemID(hMenu, 0)
    If mString1 <> "" Then ModifyMenu hMenu, MenuItem, MF_STRING, MenuItem, StrPtr(TV((mString1)))

    MenuItem = GetMenuItemID(hMenu, 1)
    If mString2 <> "" Then ModifyMenu hMenu, MenuItem, MF_STRING, MenuItem, StrPtr(TV((mString2)))

    MenuItem = GetMenuItemID(hMenu, 2)
    If mString3 <> "" Then ModifyMenu hMenu, MenuItem, MF_STRING, MenuItem, StrPtr(TV((mString3)))

    MenuItem = GetMenuItemID(hMenu, 3)
    If mString4 <> "" Then ModifyMenu hMenu, MenuItem, MF_STRING, MenuItem, StrPtr(TV((mString4)))

    MenuItem = GetMenuItemID(hMenu, 4)
    If mString5 <> "" Then ModifyMenu hMenu, MenuItem, MF_STRING, MenuItem, StrPtr(TV((mString5)))

    MenuItem = GetMenuItemID(hMenu, 6)
    If mString6 <> "" Then ModifyMenu hMenu, MenuItem, MF_STRING, MenuItem, StrPtr(TV((mString6)))
End Sub


Public Function TV(S As String) As String
    Dim ArrVni() As String, UniArr() As Variant, I As Long, J As Long, sResult As String
    sResult = S
    ArrVni = Split("a2|a1|a3|a4|a5|a6|a62|a61|a63|a64|a65|a8|a82|a81|a83|a84|a85|d9|e2|e1|e3|e4|e5|e6|e62|e61|e63|e64|e65|i2|i1|i3|i4|i5|o2|o1|o3|o4|o5|o6|o62|o61|o63|o64|o65|o7|o72|o71|o73|o74|o75|u2|u1|u3|u4|u5|u7|u72|u71|u73|u74|u75|y2|y1|y3|y4|y5", "|")
    UniArr = Array(ChrW$(224), ChrW$(225), ChrW$(7843), ChrW$(227), ChrW$(7841), ChrW$(226), ChrW$(7847), ChrW$(7845), ChrW$(7849), ChrW$(7851), ChrW$(7853), ChrW$(259), ChrW$(7857), ChrW$(7855), ChrW$(7859), ChrW$(7861), ChrW$(7863), ChrW$(273), ChrW$(232), ChrW$(233), ChrW$(7867), _
                    ChrW$(7869), ChrW$(7865), ChrW$(234), ChrW$(7873), ChrW$(7871), ChrW$(7875), ChrW$(7877), ChrW$(7879), ChrW$(236), ChrW$(237), ChrW$(7881), ChrW$(297), ChrW$(7883), ChrW$(242), ChrW$(243), ChrW$(7887), ChrW$(245), ChrW$(7885), ChrW$(244), ChrW$(7891), ChrW$(7889), _
                    ChrW$(7893), ChrW$(7895), ChrW$(7897), ChrW$(417), ChrW$(7901), ChrW$(7899), ChrW$(7903), ChrW$(7905), ChrW$(7907), ChrW$(249), ChrW$(250), ChrW$(7911), ChrW$(361), ChrW$(7909), ChrW$(432), ChrW$(7915), ChrW$(7913), ChrW$(7917), ChrW$(7919), ChrW$(7921), ChrW$(7923), ChrW$(253), ChrW$(7929), ChrW$(7925))
    
    For I = UBound(UniArr) To LBound(UniArr) Step -1
        For J = 1 To Len(S)
            If LCase$(ArrVni(I)) = LCase$(Mid$(S, J, 3)) Then sResult = Replace$(sResult, Mid$(S, J, 3), IIf(Mid$(S, J, 1) = UCase$(Mid$(S, J, 1)), UCase$(UniArr(I)), UniArr(I)))
        Next J
        
        For J = 1 To Len(S)
            If LCase$(ArrVni(I)) = LCase$(Mid$(S, J, 2)) Then sResult = Replace$(sResult, Mid$(S, J, 2), IIf(Mid$(S, J, 1) = UCase$(Mid$(S, J, 1)), UCase$(UniArr(I)), UniArr(I)))
        Next J
    Next I
    
    TV = sResult
    
End Function


