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
