VERSION 5.00
Begin VB.Form frmCampfire 
   BorderStyle     =   0  'None
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   481
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   544
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCampfire 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   0
   End
   Begin VB.TextBox cantidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3720
      TabIndex        =   1
      Text            =   "1"
      Top             =   6555
      Width           =   690
   End
   Begin VB.PictureBox interface 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   3660
      Left            =   630
      MousePointer    =   99  'Custom
      ScaleHeight     =   244
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   459
      TabIndex        =   0
      Top             =   1605
      Width           =   6885
   End
   Begin VB.Image cmdApagar 
      Height          =   420
      Left            =   5505
      Tag             =   "0"
      Top             =   6465
      Width           =   1605
   End
   Begin VB.Image salir 
      Height          =   375
      Left            =   7680
      Top             =   0
      Width           =   495
   End
   Begin VB.Image cmdCocinar 
      Height          =   420
      Left            =   825
      Tag             =   "0"
      Top             =   6465
      Width           =   1605
   End
   Begin VB.Image cmdMas 
      Height          =   315
      Left            =   4650
      Tag             =   "0"
      Top             =   6510
      Width           =   315
   End
   Begin VB.Image cmdMenos 
      Height          =   315
      Left            =   3195
      Tag             =   "0"
      Top             =   6510
      Width           =   315
   End
   Begin VB.Label lblFogataTiempo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   2850
      TabIndex        =   2
      Top             =   5745
      Width           =   1.03995e5
   End
End
Attribute VB_Name = "frmCampfire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    Argentum 20 - Game Client Program
'    Copyright (C) 2022 - Noland Studios
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU Affero General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU Affero General Public License for more details.
'    You should have received a copy of the GNU Affero General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'
Option Explicit

Private Const WM_SYSCOMMAND  As Long = &H112&
Private Const MOUSE_MOVE     As Long = &HF012&
Private Const MIN_QUANTITY   As Integer = 1
Private Const MAX_QUANTITY   As Integer = 10000
Private Const MS_PER_SECOND  As Long = 1000
Private Const SECONDS_PER_MINUTE As Long = 60
Private Const SECONDS_PER_HOUR   As Long = 3600

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lparam As Long) As Long

Public WithEvents InvCampfireUser       As clsGrapchicalInventory
Attribute InvCampfireUser.VB_VarHelpID = -1
Public WithEvents InvCampfireCoal       As clsGrapchicalInventory
Attribute InvCampfireCoal.VB_VarHelpID = -1
Public WithEvents InvCampfireIngredient As clsGrapchicalInventory
Attribute InvCampfireIngredient.VB_VarHelpID = -1
Public WithEvents InvCampfireOutput     As clsGrapchicalInventory
Attribute InvCampfireOutput.VB_VarHelpID = -1

Private m_IngredientSlot   As Byte
Private m_IngredientAmount As Integer
Private m_IngredientObj    As Integer
Private m_DragSource       As Integer
Private m_DragSlot         As Byte

Private Const DRAG_SOURCE_NONE       As Integer = 0
Private Const DRAG_SOURCE_USER       As Integer = 1
Private Const DRAG_SOURCE_INGREDIENT As Integer = 2
Private Const DRAG_SOURCE_OUTPUT     As Integer = 3

Private m_OpenMap          As Integer

Private cBotonMas              As clsGraphicalButton
Private cBotonMenos            As clsGraphicalButton
Private cBotonCerrar           As clsGraphicalButton
Private cBotonCocinar          As clsGraphicalButton
Private cBotonApagar           As clsGraphicalButton

Private Const BTN_COOK_DEFAULT As String = "boton-cocinar-default.bmp"
Private Const BTN_COOK_OVER    As String = "boton-cocinar-over.bmp"
Private Const BTN_COOK_OFF     As String = "boton-cocinar-off.bmp"
Private Const BTN_OFF_DEFAULT  As String = "boton-apagar-default.bmp"
Private Const BTN_OFF_OVER     As String = "boton-apagar-over.bmp"
Private Const BTN_OFF_OFF      As String = "boton-apagar-off.bmp"

Public campfireCookText    As String
Public campfireNextText    As String
Public campfireNeedsRedraw As Boolean

Public Sub ShowCampfireState()
    On Error GoTo ShowCampfireState_Err
    If Not Me.visible Then
        m_IngredientSlot = 0
        m_IngredientAmount = 0
        m_OpenMap = UserMap
        Me.Show , GetGameplayForm()
    End If
    tmrCampfire.enabled = True
    Call RefreshPanel
    Exit Sub
ShowCampfireState_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.ShowCampfireState", Erl)
End Sub

Public Sub ClosePanel()
    On Error GoTo ClosePanel_Err
    Unload Me
    Exit Sub
ClosePanel_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.ClosePanel", Erl)
End Sub

Public Sub RefreshPanel()
    On Error GoTo RefreshPanel_Err
    Dim ElapsedSeconds   As Long
    Dim campfireSeconds  As Long
    Dim readyNow         As Integer
    Dim secondsToNext    As Long
    ElapsedSeconds = (CampfireGetTickCount() - CampfireStateReceivedAt) \ MS_PER_SECOND
    If ElapsedSeconds < 0 Then ElapsedSeconds = 0
    campfireSeconds = CampfireSecondsLeft - ElapsedSeconds
    If campfireSeconds < 0 Then campfireSeconds = 0
    If campfireSeconds >= SECONDS_PER_HOUR Then
        lblFogataTiempo.Caption = JsonLanguage.Item("FOGATA_DURA") & " " & (campfireSeconds \ SECONDS_PER_HOUR) & " h " & ((campfireSeconds Mod SECONDS_PER_HOUR) \ SECONDS_PER_MINUTE) & " " & _
                JsonLanguage.Item("FOGATA_MINUTOS")
    Else
        lblFogataTiempo.Caption = JsonLanguage.Item("FOGATA_DURA") & " " & (campfireSeconds \ SECONDS_PER_MINUTE) & " " & JsonLanguage.Item("FOGATA_MINUTOS")
    End If
    If CookTotal > 0 Then
        readyNow = GetLocalReadyCount(ElapsedSeconds)
        If Not ValidateDeposit() Then m_IngredientAmount = 0
        If CookTotal - CookRetrieved - readyNow + m_IngredientAmount > 0 Then
            Call InvCampfireIngredient.SetItem(1, CookRawObj, CookTotal - CookRetrieved - readyNow + m_IngredientAmount, False, ObjData(CookRawObj).grhIndex, 0, 0, 0, 0, 0, _
                    ObjData(CookRawObj).name, 0, 1)
        Else
            Call InvCampfireIngredient.ClearSlot(1)
        End If
        If readyNow > 0 Then
            Call InvCampfireOutput.SetItem(1, CookCookedObj, readyNow, False, ObjData(CookCookedObj).grhIndex, 0, 0, 0, 0, 0, ObjData(CookCookedObj).name, 0, 1)
        Else
            Call InvCampfireOutput.ClearSlot(1)
        End If
        campfireCookText = ObjData(CookRawObj).name & " -> " & ObjData(CookCookedObj).name
        If CookRetrieved + readyNow >= CookTotal Then
            campfireNextText = JsonLanguage.Item("FOGATA_TODO_LISTO")
        Else
            secondsToNext = GetLocalSecondsToNext(ElapsedSeconds)
            campfireNextText = (secondsToNext \ SECONDS_PER_MINUTE) & ":" & format$(secondsToNext Mod SECONDS_PER_MINUTE, "00") & _
                    " (" & readyNow & "/" & (CookTotal - CookRetrieved) & ")"
            If m_IngredientAmount > 0 Then campfireNextText = campfireNextText & " ï¿½ " & m_IngredientAmount & " " & JsonLanguage.Item("FOGATA_POR_COCINAR")
        End If
    Else
        Call InvCampfireOutput.ClearSlot(1)
        Call RefreshDepositedIngredient
        campfireCookText = JsonLanguage.Item("FOGATA_SIN_COCCION")
        campfireNextText = vbNullString
    End If
    campfireNeedsRedraw = True
    Exit Sub
RefreshPanel_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.RefreshPanel", Erl)
End Sub

Private Function GetLocalReadyCount(ByVal ElapsedSeconds As Long) As Integer
    On Error GoTo GetLocalReadyCount_Err
    Dim extraReady As Long
    Dim pending    As Integer
    pending = CookTotal - CookRetrieved - CookReady
    If pending <= 0 Or CookSecondsPerUnit <= 0 Or ElapsedSeconds < CookSecondsToNext Then
        GetLocalReadyCount = CookReady
        Exit Function
    End If
    extraReady = 1 + (ElapsedSeconds - CookSecondsToNext) \ CookSecondsPerUnit
    If extraReady > pending Then extraReady = pending
    GetLocalReadyCount = CookReady + extraReady
    Exit Function
GetLocalReadyCount_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.GetLocalReadyCount", Erl)
End Function

Private Function GetLocalSecondsToNext(ByVal ElapsedSeconds As Long) As Long
    On Error GoTo GetLocalSecondsToNext_Err
    If ElapsedSeconds < CookSecondsToNext Then
        GetLocalSecondsToNext = CookSecondsToNext - ElapsedSeconds
    ElseIf CookSecondsPerUnit > 0 Then
        GetLocalSecondsToNext = CookSecondsPerUnit - ((ElapsedSeconds - CookSecondsToNext) Mod CookSecondsPerUnit)
    End If
    Exit Function
GetLocalSecondsToNext_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.GetLocalSecondsToNext", Erl)
End Function

Private Function GetQuantity() As Integer
    On Error GoTo GetQuantity_Err
    GetQuantity = MIN_QUANTITY
    If IsNumeric(cantidad.Text) Then
        If val(cantidad.Text) >= MIN_QUANTITY And val(cantidad.Text) <= MAX_QUANTITY Then GetQuantity = val(cantidad.Text)
    End If
    Exit Function
GetQuantity_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.GetQuantity", Erl)
End Function

Private Sub DepositIngredient(ByVal inventorySlot As Byte)
    On Error GoTo DepositIngredient_Err
    If inventorySlot = 0 Then Exit Sub
    If frmMain.Inventario.ObjType(inventorySlot) <> eObjType.otRawFood Then Exit Sub
    If CookTotal > 0 And frmMain.Inventario.ObjIndex(inventorySlot) <> CookRawObj Then
        Call AddtoRichTextBox(frmMain.RecTxt, JsonLanguage.Item("FOGATA_SALIDA_OCUPADA"), 255, 128, 0, 1, 0)
        Exit Sub
    End If
    If CookTotal > 0 Then
        ' Ya hay una coccion de la misma comida: se suma directo, sin esperar a Cocinar
        Call WriteCampfireCook(inventorySlot, min(GetQuantity(), frmMain.Inventario.Amount(inventorySlot)))
        Exit Sub
    End If
    If m_IngredientSlot = inventorySlot Then
        m_IngredientAmount = min(m_IngredientAmount + GetQuantity(), frmMain.Inventario.Amount(inventorySlot))
    Else
        Call ClearDeposit
        m_IngredientSlot = inventorySlot
        m_IngredientObj = frmMain.Inventario.ObjIndex(inventorySlot)
        m_IngredientAmount = min(GetQuantity(), frmMain.Inventario.Amount(inventorySlot))
    End If
    Call RefreshPanel
    Exit Sub
DepositIngredient_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.DepositIngredient", Erl)
End Sub

Private Sub ClearDeposit()
    On Error GoTo ClearDeposit_Err
    Dim previousSlot As Byte
    previousSlot = m_IngredientSlot
    m_IngredientSlot = 0
    m_IngredientAmount = 0
    m_IngredientObj = 0
    Call InvCampfireIngredient.ClearSlot(1)
    If previousSlot > 0 Then Call ShowUserSlotWithDeposit(previousSlot, 0)
    Exit Sub
ClearDeposit_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.ClearDeposit", Erl)
End Sub

Private Sub RemoveIngredient()
    On Error GoTo RemoveIngredient_Err
    If m_IngredientSlot > 0 Then
        ' Depositado pero sin mandar a cocinar: se devuelven N unidades visualmente
        m_IngredientAmount = m_IngredientAmount - GetQuantity()
        If m_IngredientAmount <= 0 Then Call ClearDeposit
        Call RefreshPanel
    ElseIf CookTotal > 0 Then
        ' Ya se esta cocinando: se cancelan N unidades y el server devuelve los crudos
        Call WriteCampfireCancelCook(GetQuantity())
    End If
    Exit Sub
RemoveIngredient_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.RemoveIngredient", Erl)
End Sub

Private Sub ShowUserSlotWithDeposit(ByVal inventorySlot As Byte, ByVal reservedAmount As Integer)
    On Error GoTo ShowUserSlotWithDeposit_Err
    With frmMain.Inventario
        If .Amount(inventorySlot) - reservedAmount <= 0 Then
            Call InvCampfireUser.ClearSlot(inventorySlot)
        Else
            Call InvCampfireUser.SetItem(inventorySlot, .ObjIndex(inventorySlot), .Amount(inventorySlot) - reservedAmount, .Equipped(inventorySlot), .grhIndex(inventorySlot), _
                    .ObjType(inventorySlot), .MaxHit(inventorySlot), .MinHit(inventorySlot), .Def(inventorySlot), .Valor(inventorySlot), .ItemName(inventorySlot), _
                    .ElementalTags(inventorySlot), .PuedeUsar(inventorySlot))
        End If
    End With
    Exit Sub
ShowUserSlotWithDeposit_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.ShowUserSlotWithDeposit", Erl)
End Sub

Private Function ValidateDeposit() As Boolean
    On Error GoTo ValidateDeposit_Err
    If m_IngredientSlot = 0 Then Exit Function
    If frmMain.Inventario.ObjIndex(m_IngredientSlot) <> m_IngredientObj Then
        Call ClearDeposit
        Exit Function
    End If
    m_IngredientAmount = min(m_IngredientAmount, frmMain.Inventario.Amount(m_IngredientSlot))
    If m_IngredientAmount <= 0 Then
        Call ClearDeposit
        Exit Function
    End If
    Call ShowUserSlotWithDeposit(m_IngredientSlot, m_IngredientAmount)
    ValidateDeposit = True
    Exit Function
ValidateDeposit_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.ValidateDeposit", Erl)
End Function

Private Sub RefreshDepositedIngredient()
    On Error GoTo RefreshDepositedIngredient_Err
    If Not ValidateDeposit() Then
        Call InvCampfireIngredient.ClearSlot(1)
        Exit Sub
    End If
    With frmMain.Inventario
        Call InvCampfireIngredient.SetItem(1, .ObjIndex(m_IngredientSlot), m_IngredientAmount, False, .grhIndex(m_IngredientSlot), .ObjType(m_IngredientSlot), 0, 0, 0, 0, _
                .ItemName(m_IngredientSlot), 0, 1)
    End With
    Exit Sub
RefreshDepositedIngredient_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.RefreshDepositedIngredient", Erl)
End Sub

Private Sub AddCoal(ByVal inventorySlot As Byte)
    On Error GoTo AddCoal_Err
    If inventorySlot = 0 Then Exit Sub
    If InvCampfireUser.ObjIndex(inventorySlot) <> COAL_OBJ_INDEX Then Exit Sub
    Call WriteCampfireAddCoal(inventorySlot, min(GetQuantity(), InvCampfireUser.Amount(inventorySlot)))
    Exit Sub
AddCoal_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.AddCoal", Erl)
End Sub

Private Sub cmdCocinar_Click()
    On Error GoTo cmdCocinar_Click_Err
    If m_IngredientSlot = 0 Or m_IngredientAmount <= 0 Then Exit Sub
    Call WriteCampfireCook(m_IngredientSlot, m_IngredientAmount)
    Call ClearDeposit
    Exit Sub
cmdCocinar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.cmdCocinar_Click", Erl)
End Sub

Private Sub cmdApagar_Click()
    On Error GoTo cmdApagar_Click_Err
    If MsgBox(JsonLanguage.Item("FOGATA_CONFIRMAR_APAGAR"), vbYesNo + vbQuestion) = vbYes Then
        Call WriteCampfireExtinguish
    End If
    Exit Sub
cmdApagar_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.cmdApagar_Click", Erl)
End Sub

Private Sub salir_Click()
    On Error GoTo Salir_Click_Err
    Unload Me
    Exit Sub
Salir_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.Salir_Click", Erl)
End Sub

Private Sub cmdMas_Click()
    On Error GoTo cmdMas_Click_Err
    Call ApplyQuantityStep(1)
    Exit Sub
cmdMas_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.cmdMas_Click", Erl)
End Sub

Private Sub cmdMenos_Click()
    On Error GoTo cmdMenos_Click_Err
    Call ApplyQuantityStep(-1)
    Exit Sub
cmdMenos_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.cmdMenos_Click", Erl)
End Sub

Private Sub ApplyQuantityStep(ByVal quantityStep As Long)
    On Error GoTo ApplyQuantityStep_Err
    Dim newQuantity As Long
    newQuantity = val(cantidad.Text) + quantityStep
    If newQuantity < MIN_QUANTITY Then newQuantity = MIN_QUANTITY
    If newQuantity > MAX_QUANTITY Then newQuantity = MAX_QUANTITY
    cantidad.Text = format$(newQuantity)
    Exit Sub
ApplyQuantityStep_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.ApplyQuantityStep", Erl)
End Sub

Private Sub cantidad_Change()
    On Error GoTo cantidad_Change_Err
    If val(cantidad.Text) > MAX_QUANTITY Then
        cantidad.Text = format$(MAX_QUANTITY)
        cantidad.SelStart = Len(cantidad.Text)
    ElseIf val(cantidad.Text) < MIN_QUANTITY Then
        cantidad.Text = format$(MIN_QUANTITY)
        cantidad.SelStart = Len(cantidad.Text)
    End If
    Exit Sub
cantidad_Change_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.cantidad_Change", Erl)
End Sub

Private Sub cantidad_KeyPress(KeyAscii As Integer)
    On Error GoTo cantidad_KeyPress_Err
    If KeyAscii = vbKeyEscape Then
        Unload Me
        Exit Sub
    End If
    If KeyAscii <> vbKeyBack And (KeyAscii < vbKey0 Or KeyAscii > vbKey9) Then KeyAscii = 0
    Exit Sub
cantidad_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.cantidad_KeyPress", Erl)
End Sub

Private Sub tmrCampfire_Timer()
    On Error GoTo tmrCampfire_Timer_Err
    If UserMap <> m_OpenMap Or Abs(UserPos.x - CampfirePanelX) > CAMPFIRE_PANEL_MAX_DISTANCE Or Abs(UserPos.y - CampfirePanelY) > CAMPFIRE_PANEL_MAX_DISTANCE Then
        Unload Me
        Exit Sub
    End If
    Call RefreshPanel
    Exit Sub
tmrCampfire_Timer_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.tmrCampfire_Timer", Erl)
End Sub

Private Sub Form_Load()
    On Error GoTo Form_Load_Err
    Me.Picture = LoadInterface("fogata.bmp")
    Call FormParser.Parse_Form(Me)
    cantidad.BackColor = RGB(18, 19, 13)
    Call LoadButtons
    Exit Sub
Form_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.Form_Load", Erl)
End Sub

Private Sub Form_Activate()
    On Error GoTo Form_Activate_Err
    campfireNeedsRedraw = True
    Exit Sub
Form_Activate_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.Form_Activate", Erl)
End Sub

Private Sub LoadButtons()
    On Error GoTo LoadButtons_Err
    Set cBotonMas = New clsGraphicalButton
    Set cBotonMenos = New clsGraphicalButton
    Set cBotonCerrar = New clsGraphicalButton
    Set cBotonCocinar = New clsGraphicalButton
    Set cBotonApagar = New clsGraphicalButton
    Call cBotonCerrar.Initialize(salir, "boton-cerrar-default.bmp", "boton-cerrar-over.bmp", "boton-cerrar-off.bmp", Me)
    Call cBotonMas.Initialize(cmdMas, "boton-sm-mas-default.bmp", "boton-sm-mas-over.bmp", "boton-sm-mas-off.bmp", Me)
    Call cBotonMenos.Initialize(cmdMenos, "boton-sm-menos-default.bmp", "boton-sm-menos-over.bmp", "boton-sm-menos-off.bmp", Me)
    Call cBotonCocinar.Initialize(cmdCocinar, BTN_COOK_DEFAULT, BTN_COOK_OVER, BTN_COOK_OFF, Me)
    Call cBotonApagar.Initialize(cmdApagar, BTN_OFF_DEFAULT, BTN_OFF_OVER, BTN_OFF_OFF, Me)
    Exit Sub
LoadButtons_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.LoadButtons", Erl)
    Resume Next
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo Form_KeyPress_Err
    If KeyAscii = vbKeyEscape Then Unload Me
    Exit Sub
Form_KeyPress_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.Form_KeyPress", Erl)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo Form_MouseMove_Err
    Call ReleaseCapture
    Call SendMessage(Me.hWnd, WM_SYSCOMMAND, MOUSE_MOVE, 0)
    Exit Sub
Form_MouseMove_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.Form_MouseMove", Erl)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo Form_Unload_Err
    tmrCampfire.enabled = False
    m_IngredientSlot = 0
    m_IngredientAmount = 0
    m_IngredientObj = 0
    Exit Sub
Form_Unload_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.Form_Unload", Erl)
End Sub

Private Sub Form_Deactivate()
    On Error GoTo Form_Deactivate_Err
    ' Click fuera del panel (otra ventana del juego): se cierra
    Unload Me
    Exit Sub
Form_Deactivate_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.Form_Deactivate", Erl)
End Sub

Private Sub interface_Click()
    On Error GoTo interface_Click_Err
    If InvCampfireOutput.ClickedInside Then
        If InvCampfireOutput.ObjIndex(1) > 0 Then Call WriteCampfireRetrieve
    End If
    Exit Sub
interface_Click_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.interface_Click", Erl)
End Sub

Private Sub interface_DblClick()
    On Error GoTo interface_DblClick_Err
    If InvCampfireUser.ClickedInside And InvCampfireUser.IsItemSelected Then
        If InvCampfireUser.ObjIndex(InvCampfireUser.SelectedItem) = COAL_OBJ_INDEX Then
            Call AddCoal(InvCampfireUser.SelectedItem)
        Else
            Call DepositIngredient(InvCampfireUser.SelectedItem)
        End If
    ElseIf InvCampfireIngredient.ClickedInside And InvCampfireIngredient.ObjIndex(1) > 0 Then
        Call RemoveIngredient
    End If
    Exit Sub
interface_DblClick_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.interface_DblClick", Erl)
End Sub

Private Sub InvCampfireUser_ItemDropped(ByVal Drag As Integer, ByVal Drop As Integer, ByVal x As Integer, ByVal y As Integer)
    On Error GoTo InvCampfireUser_ItemDropped_Err
    ' Solo mover dentro del propio inventario; el resto lo resuelve interface_MouseUp
    If Drop > 0 And Drag <> Drop And Drop <> m_IngredientSlot And Drag <> m_IngredientSlot Then Call WriteItemMove(Drag, Drop)
    Exit Sub
InvCampfireUser_ItemDropped_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.InvCampfireUser_ItemDropped", Erl)
End Sub

Private Function IsInsideSlotArea(ByVal x As Single, ByVal y As Single, ByVal areaX As Integer, ByVal areaY As Integer) As Boolean
    IsInsideSlotArea = x >= areaX And x < areaX + CAMPFIRE_SLOT_SIZE And y >= areaY And y < areaY + CAMPFIRE_SLOT_SIZE
End Function

Private Function IsInsideUserInventory(ByVal x As Single, ByVal y As Single) As Boolean
    IsInsideUserInventory = x >= CAMPFIRE_USER_INV_X And x < CAMPFIRE_USER_INV_X + CAMPFIRE_USER_INV_WIDTH And y >= 0 And y < interface.ScaleHeight
End Function

Private Sub interface_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo interface_MouseDown_Err
    m_DragSource = DRAG_SOURCE_NONE
    m_DragSlot = 0
    If Button <> vbRightButton Then Exit Sub
    If IsInsideUserInventory(x, y) Then
        m_DragSlot = InvCampfireUser.GetSlot(x, y)
        If m_DragSlot > 0 Then
            If InvCampfireUser.ObjIndex(m_DragSlot) > 0 Then m_DragSource = DRAG_SOURCE_USER
        End If
    ElseIf IsInsideSlotArea(x, y, CAMPFIRE_INGREDIENT_X, CAMPFIRE_INGREDIENT_Y) Then
        If InvCampfireIngredient.ObjIndex(1) > 0 Then m_DragSource = DRAG_SOURCE_INGREDIENT
    ElseIf IsInsideSlotArea(x, y, CAMPFIRE_OUTPUT_X, CAMPFIRE_OUTPUT_Y) Then
        If InvCampfireOutput.ObjIndex(1) > 0 Then m_DragSource = DRAG_SOURCE_OUTPUT
    End If
    Exit Sub
interface_MouseDown_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.interface_MouseDown", Erl)
End Sub

Private Sub interface_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo interface_MouseUp_Err
    Dim dragSource As Integer
    Dim dragSlot   As Byte
    If Button <> vbRightButton Then Exit Sub
    dragSource = m_DragSource
    dragSlot = m_DragSlot
    m_DragSource = DRAG_SOURCE_NONE
    m_DragSlot = 0
    Select Case dragSource
        Case DRAG_SOURCE_USER
            If IsInsideSlotArea(x, y, CAMPFIRE_INGREDIENT_X, CAMPFIRE_INGREDIENT_Y) Then
                Call DepositIngredient(dragSlot)
            ElseIf IsInsideSlotArea(x, y, CAMPFIRE_COAL_X, CAMPFIRE_COAL_Y) Then
                Call AddCoal(dragSlot)
            End If
        Case DRAG_SOURCE_INGREDIENT
            If IsInsideUserInventory(x, y) Then Call RemoveIngredient
        Case DRAG_SOURCE_OUTPUT
            If IsInsideUserInventory(x, y) Then Call WriteCampfireRetrieve
    End Select
    Exit Sub
interface_MouseUp_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.interface_MouseUp", Erl)
End Sub

Private Sub interface_Paint()
    On Error GoTo interface_Paint_Err
    campfireNeedsRedraw = True
    Exit Sub
interface_Paint_Err:
    Call RegistrarError(Err.Number, Err.Description, "frmCampfire.interface_Paint", Erl)
End Sub
