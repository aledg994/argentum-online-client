Attribute VB_Name = "modCampfirePanel"
Option Explicit

Public Const CAMPFIRE_PANEL_MAX_DISTANCE As Integer = 2
Public Const COAL_OBJ_INDEX              As Integer = 3391

Public Declare Function CampfireGetTickCount Lib "kernel32" Alias "GetTickCount" () As Long

Public CampfirePanelX           As Byte
Public CampfirePanelY           As Byte
Public CampfireSecondsLeft      As Long
Public CookRawObj               As Integer
Public CookCookedObj            As Integer
Public CookTotal                As Integer
Public CookRetrieved            As Integer
Public CookReady                As Integer
Public CookSecondsToNext        As Long
Public CookSecondsPerUnit       As Long
Public CampfireStateReceivedAt  As Long

Public Const CAMPFIRE_SLOT_SIZE        As Integer = 35
Public Const CAMPFIRE_USER_INV_X       As Integer = 252
Public Const CAMPFIRE_USER_INV_WIDTH   As Integer = 210
Public Const CAMPFIRE_COAL_X           As Integer = 88
Public Const CAMPFIRE_COAL_Y           As Integer = 40
Public Const CAMPFIRE_INGREDIENT_X     As Integer = 28
Public Const CAMPFIRE_INGREDIENT_Y     As Integer = 125
Public Const CAMPFIRE_OUTPUT_X         As Integer = 147
Public Const CAMPFIRE_OUTPUT_Y         As Integer = 125

Public Const CAMPFIRE_TEXT_MIN_X    As Integer = 10
Public Const CAMPFIRE_TEXT_MAX_X    As Integer = 235
Public Const CAMPFIRE_TEXT_COOK_Y   As Integer = 180
Public Const CAMPFIRE_TEXT_NEXT_Y   As Integer = 160

Public Const CAMPFIRE_SLOT_COUNT   As Integer = 1
Public Const CAMPFIRE_SLOT_SPACING As Integer = 3

Public Const CAMPFIRE_SLOT_LABEL_Y  As Integer = 111
Public Const CAMPFIRE_FUEL_LABEL_Y  As Integer = 25

Public Function GetCampfireCenteredTextX(ByVal panelText As String, ByVal CenterX As Integer) As Integer
    On Error GoTo GetCampfireCenteredTextX_Err
    Dim textwidth As Integer
    Dim textX     As Integer
    textwidth = Engine_Text_Width(panelText)
    textX = CenterX - textwidth \ 2
    If textX + textwidth > CAMPFIRE_TEXT_MAX_X Then textX = CAMPFIRE_TEXT_MAX_X - textwidth
    If textX < CAMPFIRE_TEXT_MIN_X Then textX = CAMPFIRE_TEXT_MIN_X
    GetCampfireCenteredTextX = textX
    Exit Function
GetCampfireCenteredTextX_Err:
    Call RegistrarError(Err.Number, Err.Description, "modCampfirePanel.GetCampfireCenteredTextX", Erl)
End Function
