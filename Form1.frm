VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuMenuTest 
         Caption         =   "This is a test"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'******************************************************************************
'Required API call declarations
'******************************************************************************

'CreateBrushIndirect is used to create the background brush for the menus
Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
'GetMenu is used to get the handle to the menus
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
'GetMenuInfo is used to get the current info for the menu (so we don't change anything we shouldn't by mistake)
Private Declare Function GetMenuInfo Lib "user32" (ByVal hMenu As Long, lpcmi As tagMENUINFO) As Long
'SetMenuInfo is used to set the background brush back to the menu and all sub-menus
Private Declare Function SetMenuInfo Lib "user32" (ByVal hMenu As Long, lpcmi As tagMENUINFO) As Long

'******************************************************************************
'Required API Type Definitions
'******************************************************************************

'Used in the Calls to CreateBrushIndirect
Private Type LOGBRUSH
    lbStyle As Long     'Style type (we only need to create a solid background for this example)
    lbColor As Long     'Set the color of the brush
    lbHatch As Long     'Hatch style (not used in this example because it's ignored for Solid style)
End Type

'Used in GetMenuInfo and SetMenuInfo calls
Private Type tagMENUINFO
    cbSize As Long              'The size of the type structure (use len to calculate)
    fMask As Long               'Mask of information/Actions to process
    dwStyle As Long             'Menu Style (not used in this example)
    cyMax As Long               'Maximum height of menu in pixels (not used in this example)
    hbrBack As Long             'Handle to background brush
    dwContextHelpID As Long     'Help Context ID (not used in this example)
    dwMenuData As Long          'Menu Data (again not used in this example)
End Type

'******************************************************************************
'API Constant declarations
'******************************************************************************

Private Const BS_SOLID = 0      'Solid style for brush
Private Const MIM_APPLYTOSUBMENUS = &H80000000  'Apply to Sub-Menus Mask
Private Const MIM_BACKGROUND = &H2              'Background Mask

Private Sub Form_Load()
    Dim ret As Long                 'Variable to hold return values from GetMenuInfo and SetMenuInfo
    Dim hMenu As Long               'Variable to hold the handle to the menu
    Dim hBrush As Long              'Variable to hold the handle to the background brush we are going to create
    Dim lbBrushInfo As LOGBRUSH     'Variable to hold the information to pass to the CreateBrushIndirect API
    Dim miMenuInfo As tagMENUINFO   'Variable to hold the menu info

    lbBrushInfo.lbStyle = BS_SOLID  'Set our brush type to solid
    lbBrushInfo.lbColor = vbRed     'Here we set our brush color
    lbBrushInfo.lbHatch = 0         'This value is ignored I set it to 0 to make sure nothing weird will happen
    hBrush = CreateBrushIndirect(lbBrushInfo)   'We create our brush
    hMenu = GetMenu(Me.hwnd)                    'Get the handle to the menu that we are modifying (note we pass the form's hWnd because it is the owner of the menu)
    miMenuInfo.cbSize = Len(miMenuInfo)         'Set the MenuInfo structure size so that we don't get errors
    ret = GetMenuInfo(hMenu, miMenuInfo)        'Go and get the actual menu info should return non-zero if successful
    miMenuInfo.fMask = MIM_APPLYTOSUBMENUS Or MIM_BACKGROUND    'Set the mask for the changes (changing the background for menu and all sub-menus)
    miMenuInfo.hbrBack = hBrush                 'Assign our brush to the menu info
    ret = SetMenuInfo(hMenu, miMenuInfo)        'Write our info back to the menu and we're done. (should return non-zero if successful)
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then  'Check to see if it is the right mouse button
        Me.PopupMenu mnuMenu        'Bring up pop-up menu to test
    End If
End Sub
