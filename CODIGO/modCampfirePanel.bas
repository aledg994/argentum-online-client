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
