Attribute VB_Name = "engine"
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
'
Option Explicit


Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Sub svb_run_callbacks Lib "steam_vb.dll" ()

Public RefreshRate As Integer
Private Const HORZRES As Long = 8
Private Const VERTRES As Long = 10
Private Const VREFRESH As Long = 116
Private Const TIME_MS_MOUSE As Byte = 10
Private MouseLastUpdate As Long
Private MouseTimeAcumulated As Long
Private Const MainWindowWidth = 1024
Private Const MainWindowHeight = 1024
Public Const BytesPerPixel = 4


Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Public FrameNum               As Long

'Mascotas:
Public LastOffset2X As Double
Public LastOffset2Y As Double


'Depentientes del motor grafico
Public Dialogos                 As clsDialogs
Public LucesRedondas            As clsLucesRedondas
Public LucesCuadradas           As clsLucesCuadradas

Public Cheat_X                  As Integer
Public Cheat_Y                  As Integer
''
' Maximum number of dialogs that can exist.
Public Const MAX_DIALOGS     As Byte = 100

''
' Maximum length of any dialog line without having to split it.
Public Const MAX_LENGTH      As Byte = 18

''
' Number of milliseconds to add to the lifetime per dialog character
Public Const MS_PER_CHAR     As Byte = 60

''
' Number of extra milliseconds to add to the lifetime of a new dialog
Public Const MS_ADD_EXTRA    As Integer = 5000

'Team Colors
Dim TeamColors(10) As RGBA
''
' The dialog structure
'
' @param    textLine    Array of lines of a formated chat.
' @param    x           X coord for rendering.
' @param    y           Y coord for rendering.
' @param    startTime   The time (in ms) at which the dialog was created.
' @param    lifeTime    Time (in ms) this dialog should last.
' @param    charIndex   The charIndex that created this dialog.
' @param    color       The color to be used when rendering the text.
' @param    renderable  Set to True if the chat should be rendered this frame, False otherwise
'                           (used to skip dialogs from people outside render area).
Private Type dialog
    textLine()  As String
    x           As Integer
    y           As Integer
    startTime   As Long
    lifeTime    As Long
    CharIndex   As Integer
    Color       As Long
    renderable  As Boolean
    MinChatTime As Integer
    Sube As Byte
End Type

Public scroll_dialog_pixels_per_frame As Single

''
' Array if dialogs, sorted by the charIndex.
Public dialogs(MAX_DIALOGS - 1)   As dialog

''
' The number of dialogs being used at the moment.
Public dialogCount                As Byte


Public StarGrh             As grh
Public WeatherFogX1        As Single
Public WeatherFogY1        As Single
Public WeatherFogX2        As Single
Public WeatherFogY2        As Single
Public WeatherDoFog        As Byte
Public WeatherFogCount     As Byte

Public ParticleOffsetX     As Long
Public ParticleOffsetY     As Long

Public LastOffsetX         As Integer
Public LastOffsetY         As Integer

Public EndTime             As Long

Public Const ScreenWidth   As Long = 538
Public Const ScreenHeight  As Long = 376

Public temp_rgb(3)         As RGBA

Public bRunning            As Boolean

Dim Texture      As Direct3DTexture8
Dim TransTexture As Direct3DTexture8

Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public fps                     As Long
Private FramesPerSecCounter    As Long
Public lFrameTimer            As Long
Public FrameTime               As Long

Public FadeInAlpha             As Single

Private ScrollPixelsPerFrameX  As Single
Private ScrollPixelsPerFrameY  As Single

Private TileBufferPixelOffsetX As Integer
Private TileBufferPixelOffsetY As Integer
Private TimeLast As Long

Private Const GrhFogata        As Long = 1521
Private Const GrhStar          As Long = 32472
Private Const GrhCharactersScreenUI As Long = 3839

' Colores estaticos

Public flash(3)        As RGBA
Public COLOR_EMPTY              As RGBA
Public COLOR_WHITE(3)           As RGBA
Public COLOR_RED(3)             As RGBA
Public COLOR_GREEN(3)           As RGBA
Public r As Byte
Public G As Byte
Public b As Byte
Public textcolorAsistente(3)    As RGBA

Public Sub InitEngineSprites()
    Call InitGrh(StarGrh, GrhStar, 1)
End Sub

Public Sub InitializeTeamColors()
    Call SetRGBA(TeamColors(1), 153, 217, 234)
    Call SetRGBA(TeamColors(2), 234, 133, 133)
    Call SetRGBA(TeamColors(3), 235, 87, 70)
    Call SetRGBA(TeamColors(4), 191, 173, 10)
    Call SetRGBA(TeamColors(5), 82, 173, 2)
    Call SetRGBA(TeamColors(6), 100, 45, 145)
    Call SetRGBA(TeamColors(7), 156, 33, 127)
    Call SetRGBA(TeamColors(8), 140, 27, 42)
    Call SetRGBA(TeamColors(9), 27, 136, 140)
    Call SetRGBA(TeamColors(10), 145, 75, 35)
End Sub

'Sets a Grh animation to loop indefinitely.

Public Function GetElapsedTime() As Single
    
    On Error GoTo GetElapsedTime_Err

    'Gets the time that past since the last call

    Dim Start_Time    As Currency
    Static end_time   As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then Call QueryPerformanceFrequency(timer_freq)
    
    'Get current time
    Call QueryPerformanceCounter(Start_Time)
    
    'Calculate elapsed time
    GetElapsedTime = (Start_Time - end_time) / timer_freq * 1000
    
    'Get next end time
    Call QueryPerformanceCounter(end_time)

    
    Exit Function

GetElapsedTime_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.GetElapsedTime", Erl)
    Resume Next
    
End Function



Private Sub Engine_InitExtras()
    
    On Error GoTo Engine_InitExtras_Err
    

    
    With Render_Connect_Rect
        .Top = 0
        .Left = 0
        .Right = D3DWindow.BackBufferWidth
        .Bottom = D3DWindow.BackBufferHeight
    End With
    
    With Render_Main_Rect
        .Top = 0
        .Left = 0
        .Right = frmMain.renderer.ScaleWidth
        .Bottom = frmMain.renderer.ScaleHeight
    End With
    
    Call Engine_InitColors
    
    ' Sistemas dependientes de el motor grafico
    Set LucesRedondas = New clsLucesRedondas
    Set LucesCuadradas = New clsLucesCuadradas
    Set Dialogos = New clsDialogs
    
    Call IniciarMeteorologia
    Call CargarLucesGlobales
    
    ' Fuentes graficas.
    Call Engine_Font_Initialize
    'Call Font_Create("Tahoma", 8, True, 0)
    'Call Font_Create("Verdana", 8, False, 0)
    'Call Font_Create("Verdana", 11, True, False)
    
    Call InitComposedTexture
    
    Exit Sub

Engine_InitExtras_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_InitExtras", Erl)
    Resume Next
    
End Sub

Public Sub Engine_InitColors()
    
    On Error GoTo Engine_InitColors_Err
    

    ' Colores comunes
    Call Long_2_RGBAList(COLOR_WHITE, -1)
    
    Call RGBAList(textcolorAsistente, 0, 255, 0)

    
    Exit Sub

Engine_InitColors_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_InitColors", Erl)
    Resume Next
    
End Sub

Public Sub new_engine_init(ByRef renderer As clsRenderer)
On Error GoTo NewEngineInitErr:
    Set renderer = New clsRenderer
    Call renderer.Init(frmMain.renderer.hwnd)
    frmDebug.add_text_tracebox "Renderer OK"
    Exit Sub
    
NewEngineInitErr:
    Call MsgBox(JsonLanguage.Item("MENSAJE_ERROR_MOTOR_GRAFICO") & vbNewLine & _
            JsonLanguage.Item("MENSAJE_SUGERENCIA_DRIVERS"), vbCritical, JsonLanguage.Item("MENSAJE_TITULO_ERROR"))

    
    frmDebug.add_text_tracebox "Error Number Returned: " & Err.Number

    End
  
End Sub

Public Sub engine_init()
On Error Resume Next

    Err.Clear
    If init_dx_objects() <> 0 Then
        Call MsgBox(JsonLanguage.Item("MENSAJEBOX_DIRECTX_NO_FUNCIONA"), vbCritical, App.title)
        End
    End If
#If DIRECT_PLAY = 1 Then
    Call init_direct_play(DirectX)
#End If
    
        
    
#If DEBUGGING = 1 Then
    Call list_modes(DirectD3D)
#End If
    If init_dx_device() <> 0 Then
        Call MsgBox(JsonLanguage.Item("MENSAJEBOX_ERROR_INICIO_DISPOSITIVO_DX"), vbCritical, App.title)
        End
    End If
        
    'Seteamos la matriz de proyeccion.
    Call D3DXMatrixOrthoOffCenterLH(Projection, 0, D3DWindow.BackBufferWidth, D3DWindow.BackBufferHeight, 0, -1#, 1#)
    Call D3DXMatrixIdentity(IdentityMatrix)
    Call DirectDevice.SetTransform(D3DTS_PROJECTION, Projection)
    Call DirectDevice.SetTransform(D3DTS_VIEW, IdentityMatrix)

    'Set the render states
    With DirectDevice
        Call .SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
        Call .SetVertexShader(D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1)
        Call .SetRenderState(D3DRS_LIGHTING, False)
        Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
        Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
        Call .SetRenderState(D3DRS_ALPHABLENDENABLE, True)
        Call .SetRenderState(D3DRS_ZENABLE, True)
        Call .SetRenderState(D3DRS_FILLMODE, D3DFILL_SOLID)
        Call .SetRenderState(D3DRS_CULLMODE, D3DCULL_NONE)
        Call .SetRenderState(D3DRS_ALPHAFUNC, D3DCMP_GREATER)
        Call .SetRenderState(D3DRS_ALPHAREF, 0)
    End With
    
    ' Carga de texturas
    Set SurfaceDB = New clsTexManager
    Call SurfaceDB.Init(DirectD3D8, DirectDevice, General_Get_Free_Ram_Bytes)
    
    'Sprite batching.
    Set SpriteBatch = New clsBatch
    Call SpriteBatch.Initialize(2000)

    engineBaseSpeed = 0.018
    OffsetLimitScreen = 32
    fps = 60
    FramesPerSecCounter = 60
    scroll_dialog_pixels_per_frame = 4
    ScrollPixelsPerFrameX = 8.5
    ScrollPixelsPerFrameY = 8.5
    Call Engine_InitExtras
    bRunning = True
    
    Exit Sub
    
errhandler:
    
Call MsgBox(JsonLanguage.Item("MENSAJE_ERROR_MOTOR_GRAFICO") & vbNewLine & _
            JsonLanguage.Item("MENSAJE_SUGERENCIA_DRIVERS"), vbCritical, JsonLanguage.Item("MENSAJE_TITULO_ERROR"))
    
    frmDebug.add_text_tracebox "Error Number Returned: " & Err.Number

    End

End Sub

Public Sub EngineReset()
On Error GoTo Engine_Reset_Err
    Call SpriteBatch.Release
    Set SpriteBatch = Nothing
    Call DirectDevice.Reset(D3DWindow)
    Set SpriteBatch = New clsBatch
    Call SpriteBatch.Initialize(2000)
Engine_Reset_Err:
    frmDebug.add_text_tracebox "Failed to reset error " + Err.Description
End Sub

Public Sub Engine_BeginScene(Optional ByVal Color As Long = 0)

    On Error GoTo Engine_BeginScene_Err

    If DirectDevice.TestCooperativeLevel <> D3D_OK Then
        If DirectDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
            Call engine_init
            prgRun = True
            pausa = False
            g_game_state.State = e_state_gameplay_screen
            'FIX18
            lFrameTimer = 0
            FramesPerSecCounter = 0
        ElseIf DirectDevice.TestCooperativeLevel = D3DERR_DEVICELOST Then
            'Call EngineReset https://learn.microsoft.com/en-us/windows/win32/direct3d9/lost-devices
            frmDebug.add_text_tracebox "Dx device lost, need to implement reset"
        End If
    End If
    Call DirectDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, color, 1, 0)
    Call DirectDevice.BeginScene
    Call SpriteBatch.Begin
    Exit Sub

Engine_BeginScene_Err:

    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_BeginScene", Erl)


End Sub
Public Sub Engine_EndScene(ByRef DestRect As RECT, Optional ByVal hwnd As Long = 0)

On Error GoTo ErrorHandlerDD:
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then
        Exit Sub
    End If
  
  
    Call SpriteBatch.Flush
    
    Call DirectDevice.EndScene
      
    Call DirectDevice.Present(DestRect, ByVal 0, hwnd, ByVal 0)
    
    Exit Sub
    
ErrorHandlerDD:

    'If DirectDevice.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
    '
    '    Call Engine_Init
    '
    '    prgRun = True
    '    pausa = False
    '    QueRender = 0
    '    lFrameTimer = 0
    '    FramesPerSecCounter = 0'

    'End If
        
End Sub


Public Sub Engine_Deinit()
    
    On Error GoTo Engine_Deinit_Err
    
    
    Erase MapData
    Erase charlist
    
    Set DirectDevice = Nothing
    Set DirectD3D = Nothing
    Set DirectX = Nothing
    Set SpriteBatch = Nothing
    Call RegistrarError(0, "debug log ", "clear sprite batch", 100)

    
    Exit Sub

Engine_Deinit_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_Deinit", Erl)
    Resume Next
    
End Sub

Public Sub Engine_ActFPS()
    
    On Error GoTo Engine_ActFPS_Err
    

    If FrameTime - lFrameTimer >= 1000 Then
        fps = FramesPerSecCounter
        FramesPerSecCounter = 0
        lFrameTimer = FrameTime
    End If

    
    Exit Sub

Engine_ActFPS_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_ActFPS", Erl)
    Resume Next
    
End Sub

Public Sub Draw_GrhIndex(ByVal grh_index As Long, ByVal x As Integer, ByVal y As Integer)
    
    On Error GoTo Draw_GrhIndex_Err
    

    If grh_index <= 0 Then Exit Sub
    If Not OverlapRect(RenderCullingRect, x, y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight) Then Exit Sub
    Call Batch_Textured_Box(x, y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, GrhData(grh_index).sX, GrhData(grh_index).sY, GrhData(grh_index).FileNum, COLOR_WHITE)

    
    Exit Sub

Draw_GrhIndex_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Draw_GrhIndex", Erl)
    Resume Next
    
End Sub

Public Sub Draw_GrhColor(ByVal grh_index As Long, ByVal x As Integer, ByVal y As Integer, ByRef text_color() As RGBA)
    
    On Error GoTo Draw_GrhColor_Err
    

    If grh_index <= 0 Then Exit Sub
    If Not OverlapRect(RenderCullingRect, x, y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight) Then Exit Sub
    Call Batch_Textured_Box(x, y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, GrhData(grh_index).sX, GrhData(grh_index).sY, GrhData(grh_index).FileNum, text_color)
    Exit Sub

Draw_GrhColor_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Draw_GrhColor", Erl)
    Resume Next
    
End Sub

Public Sub Draw_GrhFont(ByVal grh_index As Long, ByVal x As Integer, ByVal y As Integer, ByRef text_color() As RGBA)
    
    On Error GoTo Draw_GrhFont_Err
    

    If grh_index <= 0 Then Exit Sub
    If Not OverlapRect(RenderCullingRect, x, y, GrhData(grh_index).pixelWidth + 1, GrhData(grh_index).pixelHeight + 1) Then Exit Sub
    Call Batch_Textured_Box_Advance(x, y, GrhData(grh_index).pixelWidth + 1, GrhData(grh_index).pixelHeight + 1, GrhData(grh_index).sX, GrhData(grh_index).sY, GrhData(grh_index).FileNum, GrhData(grh_index).pixelWidth + 1, GrhData(grh_index).pixelHeight + 1, text_color)

    
    Exit Sub

Draw_GrhFont_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Draw_GrhFont", Erl)
    Resume Next
    
End Sub

Public Sub Draw_GrhIndexColor(ByVal grh_index As Long, ByVal x As Integer, ByVal y As Integer)
    
    On Error GoTo Draw_GrhIndexColor_Err
    

    If grh_index <= 0 Then Exit Sub
    If Not OverlapRect(RenderCullingRect, x, y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight) Then Exit Sub
    Call Batch_Textured_Box(x, y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, GrhData(grh_index).sX, GrhData(grh_index).sY, GrhData(grh_index).FileNum, COLOR_WHITE, True)

    
    Exit Sub

Draw_GrhIndexColor_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Draw_GrhIndexColor", Erl)
    Resume Next
    
End Sub

Public Sub Draw_Grh(ByRef grh As grh, ByVal x As Integer, ByVal y As Integer, ByVal center As Byte, ByVal animate As Byte, ByRef rgb_list() As RGBA, Optional ByVal Alpha As Boolean = False, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal angle As Single)
    
    On Error GoTo Draw_Grh_Err

    If grh.GrhIndex = 0 Or grh.GrhIndex > MaxGrh Then Exit Sub
    
    Dim CurrentFrame As Integer
    CurrentFrame = 1

    If animate Then
        If grh.started > 0 Then
            Dim ElapsedFrames As Long
            ElapsedFrames = Fix(0.5 * (FrameTime - grh.started) / grh.speed)

            If grh.Loops = INFINITE_LOOPS Or ElapsedFrames < GrhData(grh.GrhIndex).NumFrames * (grh.Loops + 1) Then
                CurrentFrame = ElapsedFrames Mod GrhData(grh.GrhIndex).NumFrames + 1

            Else
                grh.started = 0
            End If

        End If

    End If
    
    Dim CurrentGrhIndex As Long
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(grh.GrhIndex).Frames(CurrentFrame)

    'Center Grh over X,Y pos
    If center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            x = x - Int(GrhData(CurrentGrhIndex).TileWidth * TilePixelWidth \ 2) + TilePixelWidth \ 2
        End If

        If GrhData(grh.GrhIndex).TileHeight <> 1 Then
            y = y - Int(GrhData(CurrentGrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
        End If
    End If
    
    With GrhData(CurrentGrhIndex)
        If Not OverlapRect(RenderCullingRect, x, y, .pixelWidth, .pixelHeight) Then Exit Sub
        If .Tx2 = 0 And .FileNum > 0 Then
            Dim Texture As Direct3DTexture8

            Dim TextureWidth As Long, TextureHeight As Long
            Set Texture = SurfaceDB.GetTexture(.FileNum, TextureWidth, TextureHeight)
        
            .Tx1 = .sX / TextureWidth
            .Tx2 = (.sX + .pixelWidth) / TextureWidth
            .Ty1 = .sY / TextureHeight
            .Ty2 = (.sY + .pixelHeight) / TextureHeight
        End If
        
        Call Batch_Textured_Box_Pre(x, y, .pixelWidth, .pixelHeight, .Tx1, .Ty1, .Tx2, .Ty2, .FileNum, rgb_list, Alpha, angle)
    
    End With
    
    Exit Sub

Draw_Grh_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Draw_Grh", Erl)
    Resume Next
    
End Sub

Public Sub DrawSingleGrh(ByVal GrhIndex As Long, screenPos As Vector2, Alpha As Single, angle As Single, ByRef rgb_list() As RGBA)
On Error GoTo DrawSingleGrh_Err
    With GrhData(GrhIndex)
        If Not OverlapRect(RenderCullingRect, screenPos.x, screenPos.y, .pixelWidth, .pixelHeight) Then Exit Sub
        If .Tx2 = 0 And .FileNum > 0 Then
            Dim Texture As Direct3DTexture8

            Dim TextureWidth As Long, TextureHeight As Long
            Set Texture = SurfaceDB.GetTexture(.FileNum, TextureWidth, TextureHeight)

            .Tx1 = .sX / TextureWidth
            .Tx2 = (.sX + .pixelWidth) / TextureWidth
            .Ty1 = .sY / TextureHeight
            .Ty2 = (.sY + .pixelHeight) / TextureHeight
        End If

        Call Batch_Textured_Box_Pre(screenPos.x, screenPos.y, .pixelWidth, .pixelHeight, .Tx1, .Ty1, .Tx2, .Ty2, .FileNum, rgb_list, Alpha, angle)

    End With

    Exit Sub
DrawSingleGrh_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Draw_Grh", Erl)
    Resume Next
End Sub

Public Sub Draw_Grh_Breathing(ByRef grh As grh, ByVal x As Integer, ByVal y As Integer, ByVal center As Byte, ByVal animate As Byte, ByRef rgb_list() As RGBA, ByVal ease As Single, Optional ByVal Alpha As Boolean = False)
    
    On Error GoTo Draw_Grh_Breathing_Err

    If grh.GrhIndex = 0 Or grh.GrhIndex > MaxGrh Then Exit Sub
    
    Dim CurrentFrame As Integer
    CurrentFrame = 1

    If animate Then
        If grh.started > 0 Then
            Dim ElapsedFrames As Long
            ElapsedFrames = Fix(0.5 * (FrameTime - grh.started) / grh.speed)

            If grh.Loops = INFINITE_LOOPS Or ElapsedFrames < GrhData(grh.GrhIndex).NumFrames * (grh.Loops + 1) Then
                CurrentFrame = ElapsedFrames Mod GrhData(grh.GrhIndex).NumFrames + 1

            Else
                grh.started = 0
            End If

        End If

    End If
    
    Dim CurrentGrhIndex As Long
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(grh.GrhIndex).Frames(CurrentFrame)

    'Center Grh over X,Y pos
    If center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            x = x + Int(TilePixelWidth - GrhData(CurrentGrhIndex).pixelWidth) \ 2
        End If

        If GrhData(grh.GrhIndex).TileHeight <> 1 Then
            y = y - Int(GrhData(CurrentGrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
        End If
    End If

    With GrhData(CurrentGrhIndex)
        If Not OverlapRect(RenderCullingRect, x, y, .pixelWidth, .pixelHeight) Then Exit Sub
        Dim Texture As Direct3DTexture8

        Dim TextureWidth As Long, TextureHeight As Long
        Set Texture = SurfaceDB.GetTexture(.FileNum, TextureWidth, TextureHeight)

        Call SpriteBatch.SetTexture(Texture)

        Call SpriteBatch.SetAlpha(Alpha)
        
        If .Tx2 = 0 And .FileNum > 0 Then
            .Tx1 = (.sX + 0.25) / TextureWidth
            .Tx2 = (.sX + .pixelWidth) / TextureWidth
            .Ty1 = (.sY + 0.25) / TextureHeight
            .Ty2 = (.sY + .pixelHeight) / TextureHeight
        End If
        'frmdebug.add_text_tracebox ease
        'frmdebug.add_text_tracebox .Ty1
        Call SpriteBatch.DrawBreathing(x, y, .pixelWidth, .pixelHeight, ease, rgb_list, .Tx1, .Ty1, .Tx2, .Ty2)

    End With
    
    Exit Sub

Draw_Grh_Breathing_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Draw_Grh_Breathing", Erl)
    Resume Next
    
End Sub

Sub Draw_Animation(ByRef animationState As tAnimationPlaybackState, ByVal x As Integer, ByVal y As Integer, ByVal center As Byte, ByRef rgb_list() As RGBA)
On Error GoTo Draw_Animation_Err
    If animationState.PlaybackState = Stopped Then
        Exit Sub
    End If
    With FxData(GetFx(animationState))
        x = x + .OffsetX
        y = y + .OffsetY
    End With
    'Center Grh over X,Y pos
    If center Then
        If GrhData(animationState.CurrentGrh).TileWidth <> 1 Then
            x = x - Int(GrhData(animationState.CurrentGrh).TileWidth * (TilePixelWidth \ 2)) + TilePixelWidth \ 2
        End If

        If GrhData(animationState.CurrentGrh).TileHeight <> 1 Then
            y = y - Int(GrhData(animationState.CurrentGrh).TileHeight * TilePixelHeight) + TilePixelHeight
        End If
    End If
    
    
    With GrhData(animationState.CurrentGrh)
        Call RGBAList(rgb_list, 255, 255, 255, animationState.AlphaValue)
        With GrhData(.Frames(animationState.CurrentFrame))
            Dim Texture As Direct3DTexture8
    
            Dim TextureWidth As Long, TextureHeight As Long
            Set Texture = SurfaceDB.GetTexture(.FileNum, TextureWidth, TextureHeight)
    
            Call SpriteBatch.SetTexture(Texture)
    
            Call SpriteBatch.SetAlpha(animationState.Alpha)
            
            If .Tx2 = 0 And .FileNum > 0 Then
                .Tx1 = .sX / TextureWidth
                .Tx2 = (.sX + .pixelWidth) / TextureWidth
                .Ty1 = .sY / TextureHeight
                .Ty2 = (.sY + .pixelHeight) / TextureHeight
            End If
    
            Call SpriteBatch.Draw(x, y, .pixelWidth, .pixelHeight, rgb_list, .Tx1, .Ty1, .Tx2, .Ty2, 0)
        End With
    End With
    Exit Sub
Draw_Animation_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Draw_Animation", Erl)
End Sub

Sub Draw_GrhFX(ByRef grh As grh, ByVal x As Integer, ByVal y As Integer, ByVal center As Byte, ByVal animate As Byte, ByRef rgb_list() As RGBA, Optional ByVal Alpha As Boolean, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal angle As Single, Optional ByVal CharIndex As Integer)
    
    On Error GoTo Draw_GrhFX_Err
    

    

    If grh.GrhIndex = 0 Or grh.GrhIndex > MaxGrh Then Exit Sub
    
    Dim CurrentFrame As Integer
    CurrentFrame = 1

    If animate Then
        If grh.started > 0 Then
            Dim ElapsedFrames As Long
            ElapsedFrames = Fix((FrameTime - grh.started) / grh.speed)
            
            If grh.AnimacionContador > 0 Then
                grh.AnimacionContador = grh.AnimacionContador - ElapsedFrames
            End If

            If grh.Loops = INFINITE_LOOPS Or ElapsedFrames < GrhData(grh.GrhIndex).NumFrames * (grh.Loops + 1) Then
                CurrentFrame = ElapsedFrames Mod GrhData(grh.GrhIndex).NumFrames + 1

            Else
                grh.started = 0
            End If

        End If

    End If
    
    Dim CurrentGrhIndex As Long
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(grh.GrhIndex).Frames(CurrentFrame)

    If grh.AnimacionContador < grh.CantAnim * 0.1 Then
            
        grh.Alpha = grh.Alpha - 1

        Call RGBAList(rgb_list, 255, 255, 255, grh.Alpha)

    End If
    
    If grh.AnimacionContador > grh.CantAnim * 0.6 Then
        If grh.Alpha < 220 Then
            grh.Alpha = grh.Alpha + 1
        End If
        
        Call RGBAList(rgb_list, 255, 255, 255, grh.Alpha)

    End If

    'Center Grh over X,Y pos
    If center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            x = x - Int(GrhData(CurrentGrhIndex).TileWidth * (TilePixelWidth \ 2)) + TilePixelWidth \ 2
        End If

        If GrhData(grh.GrhIndex).TileHeight <> 1 Then
            y = y - Int(GrhData(CurrentGrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
        End If
    End If

    With GrhData(CurrentGrhIndex)
        If Not OverlapRect(RenderCullingRect, x, y, .pixelWidth, .pixelHeight) Then Exit Sub
        Dim Texture As Direct3DTexture8

        Dim TextureWidth As Long, TextureHeight As Long
        Set Texture = SurfaceDB.GetTexture(.FileNum, TextureWidth, TextureHeight)

        Call SpriteBatch.SetTexture(Texture)

        Call SpriteBatch.SetAlpha(Alpha)
        
        If .Tx2 = 0 And .FileNum > 0 Then
            .Tx1 = .sX / TextureWidth
            .Tx2 = (.sX + .pixelWidth) / TextureWidth
            .Ty1 = .sY / TextureHeight
            .Ty2 = (.sY + .pixelHeight) / TextureHeight
        End If

        Call SpriteBatch.Draw(x, y, .pixelWidth, .pixelHeight, rgb_list, .Tx1, .Ty1, .Tx2, .Ty2, angle)

    End With

    
    Exit Sub

Draw_GrhFX_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Draw_GrhFX", Erl)
    Resume Next
    
End Sub

Private Sub Draw_GrhSinLuz(ByRef grh As grh, ByVal x As Integer, ByVal y As Integer, ByVal center As Byte, ByVal animate As Byte, Optional ByVal Alpha As Boolean, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal angle As Single)
    
    On Error GoTo Draw_GrhSinLuz_Err
    

    

    If grh.GrhIndex = 0 Or grh.GrhIndex > MaxGrh Then Exit Sub
    
    Dim CurrentFrame As Integer
    CurrentFrame = 1

    If animate Then
        If grh.started > 0 Then
            Dim ElapsedFrames As Long
            ElapsedFrames = Fix((FrameTime - grh.started) / grh.speed)

            If grh.Loops = INFINITE_LOOPS Or ElapsedFrames < GrhData(grh.GrhIndex).NumFrames * (grh.Loops + 1) Then
                CurrentFrame = ElapsedFrames Mod GrhData(grh.GrhIndex).NumFrames + 1

            Else
                grh.started = 0
            End If

        End If

    End If
    
    Dim CurrentGrhIndex As Long
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(grh.GrhIndex).Frames(CurrentFrame)

    'Center Grh over X,Y pos
    If center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            x = x - Int(GrhData(CurrentGrhIndex).TileWidth * (TilePixelWidth \ 2)) + TilePixelWidth \ 2

        End If

        If GrhData(grh.GrhIndex).TileHeight <> 1 Then
            y = y - Int(GrhData(CurrentGrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight

        End If

    End If
    If Not OverlapRect(RenderCullingRect, x, y, GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight) Then Exit Sub
    Static light_value(3) As RGBA

    light_value(0) = global_light
    light_value(1) = light_value(0)
    light_value(2) = light_value(0)
    light_value(3) = light_value(0)

    'Device_Box_Textured_Render CurrentGrhIndex, x, y, GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, light_value, GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, Alpha, angle
    Call Batch_Textured_Box(x, y, GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, GrhData(CurrentGrhIndex).FileNum, light_value, Alpha, angle)

    
    Exit Sub

Draw_GrhSinLuz_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Draw_GrhSinLuz", Erl)
    Resume Next
    
End Sub

Public Sub render()
    
    On Error GoTo render_Err

    Rem On Error GoTo ErrorHandler:
    Dim temp_array(3) As RGBA

    Call Engine_BeginScene
    
    Call ShowNextFrame

    
    FrameTime = GetTickCount()
    FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
    
    If frmMain.Contadores.enabled Then

        Dim PosY As Integer: PosY = -10 + gameplay_render_offset.y
        Dim PosX As Integer: PosX = 640 + gameplay_render_offset.x

        If DrogaCounter > 0 Then
            Call RGBAList(temp_array, 0, 153, 0)
            If DrogaCounter > 15 Then
                Call RGBAList(temp_array, 0, 153, 0)
            ElseIf DrogaCounter > 10 Then
                Call RGBAList(temp_array, 255, 255, 0)
            Else
                Dim State As Long
                State = (FrameTime / 1000) Mod 2
                Dim Alpha As Byte
                Call RGBAList(temp_array, 230, 0, 0)
            End If
            PosY = PosY + 15
            Call Engine_Text_Render(JsonLanguage.Item("MENSAJE_542") & CLng(DrogaCounter) & "s", PosX, PosY, temp_array, 1, True, 0, 160)
        End If

    End If
    
    If FadeInAlpha > 0 Then
        Call Engine_Draw_Box(0, 0, frmMain.renderer.ScaleWidth, frmMain.renderer.ScaleHeight, RGBA_From_Comp(0, 0, 0, FadeInAlpha))
        FadeInAlpha = FadeInAlpha - 10 * timerTicksPerFrame
    End If
    
    Call Engine_EndScene(Render_Main_Rect)
    
    'TIME_MS_MOUSE
    
    MouseTimeAcumulated = MouseTimeAcumulated + timerElapsedTime
    If MouseLastUpdate + TIME_MS_MOUSE <= MouseTimeAcumulated Then
        MouseLastUpdate = MouseTimeAcumulated
    End If
    
    Engine_ActFPS

    Exit Sub

render_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.render", Erl)
    Resume Next
    
End Sub

Sub ShowNextFrame()
    
    On Error GoTo ShowNextFrame_Err
    Static OffsetCounterX As Single
    Static OffsetCounterY As Single
    
    If charlist(UserCharIndex).TranslationActive Then
        With charlist(UserCharIndex)
            Dim ElapsedTime As Long
            Dim TranslationPercent As Double
            ElapsedTime = FrameTime - .TranslationStartTime
            TranslationPercent = min(CDbl(ElapsedTime) / .TranslationTime, 1)
            OffsetCounterX = Interpolate(32 * .scrollDirectionX, 0, TranslationPercent)
            OffsetCounterY = Interpolate(32 * .scrollDirectionY, 0, TranslationPercent)
        End With
    ElseIf UserMoving Then

        '****** Move screen Left and Right if needed ******
        If AddtoUserPos.x <> 0 Then
            LastOffset2X = ScrollPixelsPerFrameX * AddtoUserPos.x * timerTicksPerFrame * charlist(UserCharIndex).Speeding
            OffsetCounterX = OffsetCounterX - LastOffset2X
            If Abs(OffsetCounterX) >= Abs(OffsetLimitScreen * AddtoUserPos.x) Then
                LastOffset2X = 0
                OffsetCounterX = 0
                AddtoUserPos.x = 0
                UserMoving = False
            End If
        End If
        '****** Move screen Up and Down if needed ******
        If AddtoUserPos.y <> 0 Then
             LastOffset2Y = ScrollPixelsPerFrameY * AddtoUserPos.y * timerTicksPerFrame * charlist(UserCharIndex).Speeding
            OffsetCounterY = OffsetCounterY - LastOffset2Y
            If Abs(OffsetCounterY) >= Abs(OffsetLimitScreen * AddtoUserPos.y) Then
                LastOffset2Y = 0
                OffsetCounterY = 0
                AddtoUserPos.y = 0
                UserMoving = False
            End If
        End If
    End If
    Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
    Call RenderScreen(UserPos.x - AddtoUserPos.x, UserPos.y - AddtoUserPos.y, OffsetCounterX, OffsetCounterY, HalfWindowTileWidth, HalfWindowTileHeight)
    Call DialogosClanes.Draw
    Call Group.RenderGroup
    Exit Sub

ShowNextFrame_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.ShowNextFrame", Erl)
    Resume Next
    
End Sub
Function ArcCos(x As Double) As Double
    ArcCos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
End Function
Function distance(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
    distance = Sqr(((y1 - y2) ^ 2 + (x1 - x2) ^ 2))
End Function
Public Function compute_vector_director(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Position
    compute_vector_director.x = x2 - x1
    compute_vector_director.y = y2 - y1
End Function

Public Function IntLerp(a As Integer, b As Integer, Factor As Single)
    Dim InvFactor As Single
    InvFactor = 1 - Factor
    IntLerp = a * InvFactor + b * Factor
End Function


Public Function calcular_direccion(ByRef dir_vector As Position) As Long
    Dim theta As Double
    Dim norma_a As Double
    Dim norma_b As Double
    
    theta = GetAngle(dir_vector.x, dir_vector.y, 1, 0) * 180 / PI
    
    ''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''
    
    Select Case Round(theta)
        Case 337 To 360, 0 To 22
            calcular_direccion = 1
        Case 23 To 67
            calcular_direccion = 2
        Case 68 To 112
            calcular_direccion = 3
        Case 113 To 157
            calcular_direccion = 4
        Case 158 To 202
            calcular_direccion = 5
        Case 203 To 247
            calcular_direccion = 6
        Case 248 To 292
            calcular_direccion = 7
        Case 293 To 336
            calcular_direccion = 8
    End Select
End Function
Public Sub Mascota_Render(ByVal CharIndex As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
    
    'calculo el pixel en el que está cada usuario y
    Dim target_x As Long
    Dim target_y As Long
    
    'target_charindex in pixels on render:
    target_x = (frmMain.renderer.ScaleWidth / 2) - ((UserPos.x - AddtoUserPos.x) - charlist(CharIndex).Pos.x) * 32 + charlist(CharIndex).MoveOffsetX
    target_y = (frmMain.renderer.ScaleHeight / 2) - ((UserPos.y - AddtoUserPos.y) - charlist(CharIndex).Pos.y) * 32 + charlist(CharIndex).MoveOffsetY
    
    Dim dir_vector As Position
    Dim dist As Long
    Dim dist_x As Long
    Dim dist_y As Long
    
    'Calculamos el vector director entre la mascota y el charindex (sin normalizar):
    dir_vector = compute_vector_director(mascota.PosX, mascota.PosY, target_x, target_y)
    dist_x = Abs(dir_vector.x)
    dist_y = Abs(dir_vector.y)
    dist = Sqr(dist_x ^ 2 + dist_y ^ 2)
    
    Dim isAnimated As Byte
    isAnimated = 1
    Static direccion As Boolean
    Static angle As Single
    'If (RandomNumber(1, 70) = 1) Then direccion = Not direccion

    angle = angle + RandomNumber(2, 10) * IIf(direccion, 1, -1) / 1500 * timerElapsedTime
    
    If dist_x > 40 Then
        mascota.PosX = mascota.PosX + (dir_vector.x / (frmMain.renderer.ScaleWidth / 2)) * timerElapsedTime / 1000 * dist * 3  ' 256 como constante no le da aceleración.
        isAnimated = 1
    End If

    mascota.PosX = mascota.PosX - LastOffset2X
    If dist_y > 40 Then
        mascota.PosY = mascota.PosY + (dir_vector.y / (frmMain.renderer.ScaleHeight / 2)) * timerElapsedTime / 1000 * dist * 3
        isAnimated = 1
    End If
    
    mascota.PosY = mascota.PosY - LastOffset2Y

    If GetTickCount() - mascota.last_time >= 200 Then
        mascota.Heading = calcular_direccion(dir_vector)
        mascota.last_time = GetTickCount()
    End If
    
    If mascota.Color(0).a < 255 Then
        Dim temp_alpha As Single

        temp_alpha = mascota.Color(0).a + 1 * timerElapsedTime / 5
        If temp_alpha > 255 Then temp_alpha = 255
        
        mascota.Color(0).a = temp_alpha
        mascota.Color(1).a = temp_alpha
        mascota.Color(2).a = temp_alpha
        mascota.Color(3).a = temp_alpha
    End If
    
    
    Call Draw_Grh(mascota.Body(mascota.Heading), mascota.PosX + Cos(angle / 2) * 5 + 150, mascota.PosY + Sin(angle) * 5 + 150, 0, isAnimated, mascota.Color)
    If mascota.Fx.started > 0 Then
        Dim colorfx(3) As RGBA
        Call RGBAList(colorfx(), 211, 153, 93, 255)
        Call Draw_Grh(mascota.Fx, mascota.PosX + Cos(angle / 2) * 5 - 7 + 150, mascota.PosY + Sin(angle) * 5 - 27 + 150, 0, isAnimated, colorfx)
    End If
End Sub


Private Sub Device_Box_Textured_Render_Advance(ByVal GrhIndex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, ByVal src_height As Integer, ByRef rgb_list() As RGBA, ByVal src_x As Integer, ByVal src_y As Integer, ByVal dest_width As Integer, Optional ByVal dest_height As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
    
    On Error GoTo Device_Box_Textured_Render_Advance_Err

    'Copies the Textures allowing resizing

    Static src_rect            As RECT

    Static dest_rect           As RECT

    Static temp_verts(3)       As TYPE_VERTEX

    Static d3dTextures         As D3D8Textures

    Static light_value(0 To 3) As RGBA
    
    If GrhIndex = 0 Then Exit Sub
    
    Set d3dTextures.Texture = SurfaceDB.GetTexture(GrhData(GrhIndex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
        
    'If Not char_current_blind Then
    '    If (light_value(0) = 0) Then light_value(0) = 0
    '    If (light_value(1) = 0) Then light_value(1) = 0
    '    If (light_value(2) = 0) Then light_value(2) = 0
    '    If (light_value(3) = 0) Then light_value(3) = 0
    'Else
    '    light_value(0) = &HFFFFFFFF 'blind_color
    '    light_value(1) = &HFFFFFFFF 'blind_color
    '    light_value(2) = &HFFFFFFFF 'blind_color
    '    light_value(3) = &HFFFFFFFF 'blind_color
    'End If
    
    'Set up the source rectangle
    With src_rect
        .Bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y

    End With
        
    'Set up the destination rectangle
    With dest_rect
        .Bottom = dest_y + dest_height
        .Left = dest_x
        .Right = dest_x + dest_width
        .Top = dest_y

    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value, d3dTextures.texwidth, d3dTextures.texheight, angle
        
    'Set Textures
    DirectDevice.SetTexture 0, d3dTextures.Texture
    
    If alpha_blend Then
        'Set Rendering for alphablending
        DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

    End If
    
    'Draw the triangles that make up our square Textures
    DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

    End If

    
    Exit Sub

Device_Box_Textured_Render_Advance_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Device_Box_Textured_Render_Advance", Erl)
    Resume Next
    
End Sub

Public Sub Batch_Textured_Box(ByVal x As Long, ByVal y As Long, _
                                ByVal Width As Integer, ByVal Height As Integer, _
                                ByVal sX As Integer, ByVal sY As Integer, _
                                ByVal tex As Long, _
                                ByRef Color() As RGBA, _
                                Optional ByVal Alpha As Boolean = False, _
                                Optional ByVal angle As Single = 0, _
                                Optional ByVal ScaleX As Single = 1!, _
                                Optional ByVal ScaleY As Single = 1!)
    
    On Error GoTo Batch_Textured_Box_Err
    

    Dim Texture As Direct3DTexture8
        
    Dim TextureWidth As Long, TextureHeight As Long
    Set Texture = SurfaceDB.GetTexture(tex, TextureWidth, TextureHeight)
    
    With SpriteBatch

        Call .SetTexture(Texture)
            
        Call .SetAlpha(Alpha)

        If TextureWidth <> 0 And TextureHeight <> 0 Then
            Call .Draw(x, y, Width * ScaleX, Height * ScaleY, Color, (sX + 0.25) / TextureWidth, (sY + 0.25) / TextureHeight, (sX + Width) / TextureWidth, (sY + Height) / TextureHeight, angle)
        Else
            Call .Draw(x, y, TextureWidth * ScaleX, TextureHeight * ScaleY, Color, , , , , angle)
        End If
            
    End With
        
    
    Exit Sub

Batch_Textured_Box_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Batch_Textured_Box", Erl)
    Resume Next
    
End Sub

Public Sub Batch_Textured_Box_Advance(ByVal x As Long, ByVal y As Long, _
                                ByVal Width As Integer, ByVal Height As Integer, _
                                ByVal sX As Integer, ByVal sY As Integer, _
                                ByVal tex As Long, _
                                ByVal dw As Integer, ByVal dH As Integer, _
                                ByRef Color() As RGBA, _
                                Optional ByVal Alpha As Boolean = False, _
                                Optional ByVal angle As Single = 0, _
                                Optional ByVal ScaleX As Single = 1!, _
                                Optional ByVal ScaleY As Single = 1!, _
                                Optional ByVal z As Long = 1)
    
    On Error GoTo Batch_Textured_Box_Advance_Err
    

    Dim Texture As Direct3DTexture8
        
    Dim TextureWidth As Long, TextureHeight As Long
    Set Texture = SurfaceDB.GetTexture(tex, TextureWidth, TextureHeight)
    
    With SpriteBatch

        Call .SetTexture(Texture)
            
        Call .SetAlpha(Alpha)
        
        If TextureWidth <> 0 And TextureHeight <> 0 Then
            Call .Draw(x, y, dw * ScaleX, dH * ScaleY, Color, (sX + 0.25) / TextureWidth, (sY + 0.25) / TextureHeight, (sX + Width) / TextureWidth, (sY + Height) / TextureHeight, angle)
        Else
            Call .Draw(x, y, TextureWidth * ScaleX, TextureHeight * ScaleY, Color, , , , , angle)
        End If
            
    End With
        
    
    Exit Sub

Batch_Textured_Box_Advance_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Batch_Textured_Box_Advance", Erl)
    Resume Next
    
End Sub

Public Sub Batch_Textured_Box_Pre(ByVal x As Long, ByVal y As Long, _
                                ByVal Width As Integer, ByVal Height As Integer, _
                                ByVal sX As Single, ByVal sY As Single, _
                                ByVal sW As Single, ByVal sH As Single, _
                                ByVal tex As Long, _
                                ByRef Color() As RGBA, _
                                Optional ByVal Alpha As Boolean = False, _
                                Optional ByVal angle As Single = 0, _
                                Optional ByVal ScaleX As Single = 1!, _
                                Optional ByVal ScaleY As Single = 1!)
    
    On Error GoTo Batch_Textured_Box_Pre_Err
    

    Dim Texture As Direct3DTexture8
        
    Dim TextureWidth As Long, TextureHeight As Long
    Set Texture = SurfaceDB.GetTexture(tex, TextureWidth, TextureHeight)
    
    With SpriteBatch

        Call .SetTexture(Texture)

        Call .SetAlpha(Alpha)

        Call .Draw(x, y, Width * ScaleX, Height * ScaleY, Color, sX, sY, sW, sH, angle)

    End With
        
    
    Exit Sub

Batch_Textured_Box_Pre_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Batch_Textured_Box_Pre", Erl)
    Resume Next
    
End Sub

Public Sub Batch_Textured_Box_Shadow(ByVal x As Long, ByVal y As Long, _
                                ByVal Width As Integer, ByVal Height As Integer, _
                                ByVal sX As Integer, ByVal sY As Integer, _
                                ByVal tex As Long, _
                                ByRef Color() As RGBA)
    
    On Error GoTo Batch_Textured_Box_Shadow_Err
    

    Dim Texture As Direct3DTexture8
        
    Dim TextureWidth As Long, TextureHeight As Long
    Set Texture = SurfaceDB.GetTexture(tex, TextureWidth, TextureHeight)
    
    With SpriteBatch

        Call .SetTexture(Texture)
            
        Call .SetAlpha(False)
        
        If TextureWidth <> 0 And TextureHeight <> 0 Then
            Call .DrawShadow(x, y, Width, Height, Color, (sX + 0.25) / TextureWidth, (sY + 0.25) / TextureHeight, (sX + Width) / TextureWidth, (sY + Height) / TextureHeight)
        Else
            Call .DrawShadow(x, y, TextureWidth, TextureHeight, Color)
        End If
            
    End With
        
    
    Exit Sub

Batch_Textured_Box_Shadow_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Batch_Textured_Box_Shadow", Erl)
    Resume Next
    
End Sub

Public Sub Device_Box_Textured_Render(ByVal GrhIndex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, ByVal src_height As Integer, ByRef Color() As RGBA, ByVal src_x As Integer, ByVal src_y As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
    
    On Error GoTo Device_Box_Textured_Render_Err

    'Just copies the Textures

    Static src_rect            As RECT

    Static dest_rect           As RECT

    Static temp_verts(3)       As TYPE_VERTEX

    Static d3dTextures         As D3D8Textures
    
    If GrhIndex = 0 Then Exit Sub
    
    Set d3dTextures.Texture = SurfaceDB.GetTexture(GrhData(GrhIndex).FileNum, d3dTextures.texwidth, d3dTextures.texheight)

    'Set up the source rectangle
    With src_rect
        .Bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y
    End With
                
    'Set up the destination rectangle
    With dest_rect
        .Bottom = dest_y + src_height
        .Left = dest_x
        .Right = dest_x + src_width
        .Top = dest_y
    End With
    
    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, Color(), d3dTextures.texwidth, d3dTextures.texheight, angle
     
    'Set Textures
    DirectDevice.SetTexture 0, d3dTextures.Texture
    
    If alpha_blend Then
        'Set Rendering for alphablending
        DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

    End If
    
    'Draw the triangles that make up our square Textures
    DirectDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
    
    If alpha_blend Then
        'Set Rendering for colokeying
        DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

    End If
    
    DirectDevice.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    DirectDevice.SetRenderState D3DRS_ALPHATESTENABLE, 1
 
    'DirectDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_DIFFUSE
    'DirectDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1

    
    Exit Sub

Device_Box_Textured_Render_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Device_Box_Textured_Render", Erl)
    Resume Next
    
End Sub

Sub Char_TextRender(ByVal CharIndex As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, ByVal x As Byte, ByVal y As Byte)
    
    On Error GoTo Char_TextRender_Err
    

    Dim moved         As Boolean

    Dim Pos           As Integer

    Dim line          As String

    Dim Color(0 To 3) As Long

    Dim i             As Long
    
    Dim screen_x      As Integer

    Dim screen_y      As Integer
    
    ' screen_x = Convert_Tile_To_View_X(PixelOffsetX) + MoveOffsetX
    ' screen_y = Convert_Tile_To_View_Y(PixelOffsetY) +

    With charlist(CharIndex)

        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        'screen_x = Convert_Tile_To_View_X(PixelOffsetX) + MoveOffsetX

        '*** Start Dialogs ***
        If .dialog <> "" Then

            'Figure out screen position
            Dim temp_array(3) As RGBA

            Dim PixelY        As Integer

            PixelY = PixelOffsetY
            
            Call Long_2_RGBAList(temp_array, .dialog_color)

            If .dialog_scroll Then
                .dialog_offset_counter_y = .dialog_offset_counter_y + (scroll_dialog_pixels_per_frame * timerTicksPerFrame * Sgn(-1))

                If Sgn(.dialog_offset_counter_y) = -1 Then
                    .dialog_offset_counter_y = 0
                    .dialog_scroll = False

                End If

                Engine_Text_Render .dialog, PixelOffsetX + 14 - CInt(Engine_Text_Width(.dialog, True) / 2), PixelY + .Body.HeadOffset.y - Engine_Text_Height(.dialog, True) + .dialog_offset_counter_y - 34, temp_array, 1, True, MapData(x, y).CharIndex
            Else
                Engine_Text_Render .dialog, PixelOffsetX + 14 - CInt(Engine_Text_Width(.dialog, True) / 2), PixelY + .Body.HeadOffset.y - 34 - Engine_Text_Height(.dialog, True), temp_array, 1, True, MapData(x, y).CharIndex

            End If

        End If
        
        If UBound(.DialogEffects) > 0 Then

            For i = 1 To UBound(.DialogEffects)
            
                With .DialogEffects(i)
                
                    If LenB(.Text) <> 0 Then
                        Dim DeltaTime As Long
                        DeltaTime = FrameTime - .Start
    
                        If DeltaTime > 1300 Then
                            .Text = vbNullString
                        Else
                            If DeltaTime > 900 Then
                                Call RGBAList(temp_array, .Color.r, .Color.G, .Color.b, .Color.a * (1300 - DeltaTime) * 0.0025)
                            Else
                                Call RGBAList(temp_array, .Color.r, .Color.G, .Color.b, .Color.a)
                            End If
                    
                            Engine_Text_Render_Efect CharIndex, .Text, PixelOffsetX + 14 - Int(Engine_Text_Width(.Text, True) * 0.5), PixelOffsetY + charlist(CharIndex).Body.HeadOffset.y - Engine_Text_Height(.Text, True) - DeltaTime * 0.025, temp_array, 1, True
            
                        End If
                    End If
                    
                End With
                
            Next

        End If
            
        '*** End Dialogs ***
    End With

    
    Exit Sub

Char_TextRender_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Char_TextRender", Erl)
    Resume Next
    
End Sub

Sub Char_Render(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, ByVal x As Byte, ByVal y As Byte)
    
    On Error GoTo Char_Render_Err

    'Draw char's to screen without offcentering them
    
    Dim Pos                 As Integer

    Dim line                As String

    Dim Color(3)            As RGBA
    
    Dim NameColor(3)        As RGBA
    Dim NameColorClan(3)    As RGBA

    Dim colorCorazon(3)     As RGBA

    Dim i                   As Long

    Dim OffsetYname         As Byte

    Dim OffsetYClan         As Byte
    
    Dim TextureX            As Integer

    Dim TextureY            As Integer
    
    Dim OffArma             As Single

    Dim OffAuras            As Integer

    Dim OffHead             As Single
    
    Dim MostrarNombre       As Boolean
    
    Dim TempGrh As grh
    
    Dim terrainHeight As Integer
    
    With charlist(CharIndex)

        If .Heading = 0 Then Exit Sub
        Dim dibujaMiembroClan As Boolean
        dibujaMiembroClan = False
        
        Dim verVidaClan As Boolean
        
        verVidaClan = False
        If .clan_index > 0 Then
            If .clan_index = charlist(UserCharIndex).clan_index And CharIndex <> UserCharIndex And .Muerto = 0 Then
                If .clan_nivel >= 6 Then
                    dibujaMiembroClan = True
                End If
                
                If .clan_nivel >= 5 Then
                    verVidaClan = True
                End If
            End If
        End If
        
        If .Moving Then
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame * .Speeding
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame * .Speeding
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
            End If
            If .scrollDirectionX = 0 And .scrollDirectionY = 0 Then
                .Moving = False
            End If
        ElseIf .TranslationActive Then
            Dim ElapsedTime As Long
            Dim TranslationPercent As Double
            ElapsedTime = FrameTime - .TranslationStartTime
            TranslationPercent = min(CDbl(ElapsedTime) / .TranslationTime, 1)
            .MoveOffsetX = Interpolate(32 * .scrollDirectionX * -1, 0, TranslationPercent)
            .MoveOffsetY = Interpolate(32 * .scrollDirectionY * -1, 0, TranslationPercent)
            If TranslationPercent >= 1 Then
                .TranslationActive = False
            End If
        ElseIf .AnimatingBody Then
            If .Body.Walk(.Heading).started = 0 Then
                .AnimatingBody = 0
                If .iBody Then
                    .Body = BodyData(.iBody)
                    .Body.Walk(.Heading).started = FrameTime
                Else
                    .Body = BodyData(0)
                End If
                .Body.Walk(.Heading).Loops = -1
                If .Idle Or .Navegando Then
                    'Start animation
                    .Body.Walk(.Heading).started = FrameTime
                End If
            End If
        ElseIf Not .Idle Then
            If .Muerto Then
                If CharIndex <> UserCharIndex Then
                    ' Si no somos nosotros, esperamos un intervalo
                    ' antes de poner la animación idle para evitar saltos
                    If FrameTime - .LastStep > TIME_CASPER_IDLE Then
                        .Body = BodyData(CASPER_BODY_IDLE)
                        .Body.Walk(.Heading).started = FrameTime
                        .Idle = True
                    End If
                Else
                    .Body = BodyData(CASPER_BODY_IDLE)
                    .Body.Walk(.Heading).started = FrameTime
                    .Idle = True
                End If
            Else
                'Stop animations
                If .Navegando = False Or UserNadandoTrajeCaucho = True Then
                    If .Body.AnimateOnIdle = 0 Then
                        .Body.Walk(.Heading).started = 0
                    ElseIf .Body.Walk(.Heading).started = 0 Then
                        .Body.Walk(.Heading).started = FrameTime
                    End If
                    If Not .MovArmaEscudo Then
                        .Arma.WeaponWalk(.Heading).started = 0
                        .Escudo.ShieldWalk(.Heading).started = 0
                    End If
                    If .Body.IdleBody > 0 Then
                        .Body = BodyData(.Body.IdleBody)
                        .Body.Walk(.Heading).started = FrameTime
                        
                    End If
                End If
            End If
            .Idle = True
        End If
        
        
        terrainHeight = TileEngine.GetTerrainHeight(x, y)
        If (.Moving) Then
            Dim prevTerrainHeight As Integer
            If (.MoveOffsetX < 0) Then
                prevTerrainHeight = TileEngine.GetTerrainHeight(x - 1, y)
            ElseIf (.MoveOffsetX > 0) Then
                prevTerrainHeight = TileEngine.GetTerrainHeight(x + 1, y)
            End If
            terrainHeight = IntLerp(terrainHeight, prevTerrainHeight, Abs(.MoveOffsetX / 32))
            
            If .HasCart Then
                If .Cart.Walk(.Heading).started = 0 Then
                    .Cart.Walk(.Heading).Loops = -1
                    .Cart.Walk(.Heading).started = FrameTime
                End If
            End If
        Else
            If .HasCart Then
                If .Cart.Walk(.Heading).started <> 0 Then
                    .Cart.Walk(.Heading).Loops = 0
                    .Cart.Walk(.Heading).started = 0
                End If
            End If
        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY - terrainHeight
        
        Dim ease As Single
        If MostrarRespiracion Then
            ease = EaseBreathing((((FrameTime - .TimeCreated) * 0.25) Mod 1000) * 0.001)
        Else
            ease = 0
        End If

        If .Body.Walk(.Heading).GrhIndex Then

            If UserCiego Then
                MostrarNombre = False
                
            ElseIf .Invisible Then

                If IsCharVisible(CharIndex) Then
                    Call RGBAList(Color, 255, 255, 255, 100)
                    
                    If .priv = 0 Then
                        
                        Select Case .status
                            ' Criminal
                            Case 0
                                Call SetRGBA(NameColor(0), ColoresPJ(23).r, ColoresPJ(23).G, ColoresPJ(23).b)
                            
                            ' Ciudadano
                            Case 1
                                Call SetRGBA(NameColor(0), ColoresPJ(20).r, ColoresPJ(20).G, ColoresPJ(20).b)
                            
                            ' Caos
                            Case 2
                                Call SetRGBA(NameColor(0), ColoresPJ(24).r, ColoresPJ(24).G, ColoresPJ(24).b)
    
                            ' Armada
                            Case 3
                                Call SetRGBA(NameColor(0), ColoresPJ(21).r, ColoresPJ(21).G, ColoresPJ(21).b)
                                
                            ' Concilio
                            Case 4
                                Call SetRGBA(NameColor(0), ColoresPJ(25).r, ColoresPJ(25).G, ColoresPJ(25).b)
                            ' Consejo
                            Case 5
                                Call SetRGBA(NameColor(0), ColoresPJ(22).r, ColoresPJ(22).G, ColoresPJ(22).b)
    
                        End Select
                    Else
                        Call SetRGBA(NameColor(0), ColoresPJ(.priv).r, ColoresPJ(.priv).G, ColoresPJ(.priv).b)
                    End If
                    
                    Call LerpRGBA(NameColor(0), NameColor(0), RGBA_From_Comp(0, 0, 0), 0.5)
                    Call RGBA_ToList(NameColor, NameColor(0))
                    Call RGBA_ToList(colorCorazon, NameColor(0))
                    
                    MostrarNombre = True
                        
                Else
                    'refactorizar bien esto, es un asco sino
                    If .Navegando Then
                        MostrarNombre = True
                        Call RGBAList(Color, 125, 125, 125, 125)
                    Else
                        MostrarNombre = False
                        Call RGBAList(Color, 0, 0, 0, 0)
                    End If
                    
                    If dibujaMiembroClan Then
                        MostrarNombre = True
                         If .priv = 0 Then
                        
                        Select Case .status
                            ' Criminal
                            Case 0
                                Call SetRGBA(NameColor(0), ColoresPJ(23).r, ColoresPJ(23).G, ColoresPJ(23).b)
                            
                            ' Ciudadano
                            Case 1
                                Call SetRGBA(NameColor(0), ColoresPJ(20).r, ColoresPJ(20).G, ColoresPJ(20).b)
                            
                            ' Caos
                            Case 2
                                Call SetRGBA(NameColor(0), ColoresPJ(24).r, ColoresPJ(24).G, ColoresPJ(24).b)
    
                            ' Armada
                            Case 3
                                Call SetRGBA(NameColor(0), ColoresPJ(21).r, ColoresPJ(21).G, ColoresPJ(21).b)
                                
                            ' Concilio
                            Case 4
                                Call SetRGBA(NameColor(0), ColoresPJ(25).r, ColoresPJ(25).G, ColoresPJ(25).b)
                            ' Consejo
                            Case 5
                                Call SetRGBA(NameColor(0), ColoresPJ(22).r, ColoresPJ(22).G, ColoresPJ(22).b)
    
                        End Select
                    Else
                        Call SetRGBA(NameColor(0), ColoresPJ(.priv).r, ColoresPJ(.priv).G, ColoresPJ(.priv).b)
                    End If
                    
                        Call LerpRGBA(NameColor(0), NameColor(0), RGBA_From_Comp(0, 0, 0), 0.5)
                        Call RGBA_ToList(NameColor, NameColor(0))
                        Call RGBA_ToList(colorCorazon, NameColor(0))
                        Call RGBAList(Color, 180, 160, 160, 160)
                    End If
                End If

            Else
                If .Muerto Then
                    Call Copy_RGBAList_WithAlpha(Color, MapData(x, y).light_value, 150)
                Else
                    Call Copy_RGBAList(Color, MapData(x, y).light_value)
                End If

                If .EsNpc Then
                    If Abs(tX - .Pos.x) < 1 And tY - .Pos.y < 1 And .Pos.y - tY < 2 Then
                        MostrarNombre = True
                        Call RGBAList(NameColor, 210, 105, 30)
                        Call InitGrh(TempGrh, 839)
                        
                        If .UserMaxHp > 0 Then
                            Dim TempColor(3) As RGBA
                            Call RGBAList(TempColor, 255, 255, 255, 200)
                            Call Draw_Grh(TempGrh, PixelOffsetX + 1 + .Body.BodyOffset.x, PixelOffsetY + 10 + .Body.BodyOffset.y, 1, 0, TempColor, False, 0, 0, 0)

                            Engine_Draw_Box PixelOffsetX + 5 + .Body.BodyOffset.x, PixelOffsetY + 37 + .Body.BodyOffset.y, .UserMinHp / .UserMaxHp * 26, 4, RGBA_From_Comp(255, 0, 0, 255)
                        End If
                    End If

                    If .simbolo <> 0 Then
                        Call Draw_GrhIndex(5257 + .simbolo, PixelOffsetX + 6 + .Body.HeadOffset.x, PixelOffsetY + .Body.HeadOffset.y - 48 - 8 * Sin((FrameTime Mod 31415) * 0.002) ^ 2)
                    End If
                    
                Else
                    MostrarNombre = True
                    
                     If .priv = 0 Then
                        
                        Select Case .status
                            ' Criminal
                            Case 0
                                Call RGBAList(NameColor, ColoresPJ(23).r, ColoresPJ(23).G, ColoresPJ(23).b)
                                Call RGBAList(colorCorazon, ColoresPJ(23).r, ColoresPJ(23).G, ColoresPJ(23).b)
                            ' Ciudadano
                            Case 1
                                Call RGBAList(NameColor, ColoresPJ(20).r, ColoresPJ(20).G, ColoresPJ(20).b)
                                Call RGBAList(colorCorazon, ColoresPJ(20).r, ColoresPJ(20).G, ColoresPJ(20).b)
                            
                            ' Caos
                            Case 2
                                Call RGBAList(NameColor, ColoresPJ(24).r, ColoresPJ(24).G, ColoresPJ(24).b)
                                Call RGBAList(colorCorazon, ColoresPJ(24).r, ColoresPJ(24).G, ColoresPJ(24).b)
    
                            ' Armada
                            Case 3
                                Call RGBAList(NameColor, ColoresPJ(21).r, ColoresPJ(21).G, ColoresPJ(21).b)
                                Call RGBAList(colorCorazon, ColoresPJ(21).r, ColoresPJ(21).G, ColoresPJ(21).b)
                                
                            ' Concilio
                            Case 4
                                Call RGBAList(NameColor, ColoresPJ(25).r, ColoresPJ(25).G, ColoresPJ(25).b)
                                Call RGBAList(colorCorazon, ColoresPJ(25).r, ColoresPJ(25).G, ColoresPJ(25).b)
                            ' Consejo
                            Case 5
                                Call RGBAList(NameColor, ColoresPJ(22).r, ColoresPJ(22).G, ColoresPJ(22).b)
                                Call RGBAList(colorCorazon, ColoresPJ(22).r, ColoresPJ(22).G, ColoresPJ(22).b)
    
                        End Select
                    Else
                        Call RGBAList(NameColor, ColoresPJ(.priv).r, ColoresPJ(.priv).G, ColoresPJ(.priv).b)
                        Call RGBAList(colorCorazon, ColoresPJ(.priv).r, ColoresPJ(.priv).G, ColoresPJ(.priv).b)
                    End If
                    
                    
                    
                                        
                    If .group_index > 0 Then
                        If charlist(CharIndex).group_index = charlist(UserCharIndex).group_index Then
                            Call Copy_RGBAList(Color, COLOR_WHITE)
                            Call SetRGBA(colorCorazon(0), 255, 255, 0)
                            Call SetRGBA(colorCorazon(1), 0, 255, 255)
                            Call SetRGBA(colorCorazon(2), 0, 255, 0)
                            Call SetRGBA(colorCorazon(3), 0, 255, 255)
                        End If
                    End If
                End If
            End If
            
            If UserCharIndex > 0 Then
                If (verVidaClan And Not .Invisible) Or dibujaMiembroClan Or charlist(UserCharIndex).priv = 5 Then
                    OffsetYname = 8
                    OffsetYClan = 8
                    Call DibujarVidaChar(CharIndex, PixelOffsetX, PixelOffsetY, OffsetYname, OffsetYClan)
                End If
            End If
            
            ' Si tiene cabeza, componemos la textura
            If .Head.Head(.Heading).GrhIndex Then
            
                If .EsEnano Then
                    OffArma = 7
                    OffAuras = 7
                End If

                OffArma = OffArma - Int(ease * 3) - .Body.BodyOffset.y
                OffHead = .Body.HeadOffset.y - Int(ease * 1.75) - 1 - .Body.BodyOffset.y
    
                BeginComposedTexture
                
                TextureX = ComposedTextureCenterX - 16
                TextureY = ComposedTextureHeight - 32
                                
                Select Case .Heading

                    Case E_Heading.EAST
    
                        If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), TextureX, TextureY + OffArma, 1, 1, COLOR_WHITE, False, x, y)
                        Call Draw_Grh_Breathing(.Body.Walk(.Heading), TextureX, TextureY, 1, 1, COLOR_WHITE, ease)
                        Call Draw_Grh(.Head.Head(.Heading), TextureX + .Body.HeadOffset.x, TextureY + OffHead, 1, 0, COLOR_WHITE, False, x, y)
                        If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), TextureX + .Body.HeadOffset.x, TextureY + OffHead, 1, 0, COLOR_WHITE, False, x, y)
                        If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), TextureX, TextureY + OffArma, 1, 1, COLOR_WHITE, False, x, y)
                        If .HasCart Then Call Draw_Grh(.Cart.Walk(.Heading), TextureX, TextureY + .Cart.HeadOffset.y, 1, 1, COLOR_WHITE, False, x, y)
                    Case E_Heading.NORTH
                        If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), TextureX, TextureY + OffArma, 1, 1, COLOR_WHITE, False, x, y)
                        If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), TextureX, TextureY + OffArma, 1, 1, COLOR_WHITE, False, x, y)
                        Call Draw_Grh_Breathing(.Body.Walk(.Heading), TextureX, TextureY, 1, 1, COLOR_WHITE, ease)
                        Call Draw_Grh(.Head.Head(.Heading), TextureX + .Body.HeadOffset.x, TextureY + OffHead, 1, 0, COLOR_WHITE, False, x, y)
                        If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), TextureX + .Body.HeadOffset.x, TextureY + OffHead, 1, 0, COLOR_WHITE, False, x, y)
                        
                    Case E_Heading.WEST
                        If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), TextureX, TextureY + OffArma, 1, 1, COLOR_WHITE, False, x, y)
                        Call Draw_Grh_Breathing(.Body.Walk(.Heading), TextureX, TextureY, 1, 1, COLOR_WHITE, ease)
                        Call Draw_Grh(.Head.Head(.Heading), TextureX + .Body.HeadOffset.x, TextureY + OffHead, 1, 0, COLOR_WHITE, False, x, y)
                        If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), TextureX + .Body.HeadOffset.x, TextureY + OffHead, 1, 0, COLOR_WHITE, False, x, y)
                        If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), TextureX, TextureY + OffArma, 1, 1, COLOR_WHITE, False, x, y)
                        If .HasCart Then Call Draw_Grh(.Cart.Walk(.Heading), TextureX, TextureY + .Cart.HeadOffset.y, 1, 1, COLOR_WHITE, False, x, y)
                    Case E_Heading.south
                        If .HasCart Then Call Draw_Grh(.Cart.Walk(.Heading), TextureX, TextureY + .Cart.HeadOffset.y, 1, 1, COLOR_WHITE, False, x, y)
                        Call Draw_Grh_Breathing(.Body.Walk(.Heading), TextureX, TextureY, 1, 1, COLOR_WHITE, ease)
                        Call Draw_Grh(.Head.Head(.Heading), TextureX + .Body.HeadOffset.x, TextureY + OffHead, 1, 0, COLOR_WHITE, False, x, y)
                        If .Casco.Head(.Heading).GrhIndex Then Call Draw_Grh(.Casco.Head(.Heading), TextureX + .Body.HeadOffset.x, TextureY + OffHead, 1, 0, COLOR_WHITE, False, x, y)
                        If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call Draw_Grh(.Escudo.ShieldWalk(.Heading), TextureX, TextureY + OffArma, 1, 1, COLOR_WHITE, False, x, y)
                        If .Arma.WeaponWalk(.Heading).GrhIndex Then Call Draw_Grh(.Arma.WeaponWalk(.Heading), TextureX, TextureY + OffArma, 1, 1, COLOR_WHITE, False, x, y)
                End Select
                     
                EndComposedTexture
                        
                If Not .Invisible Or dibujaMiembroClan Or .Navegando Then
                    ' Reflejo
                    PresentComposedTexture PixelOffsetX + .Body.BodyOffset.x, PixelOffsetY + .Body.BodyOffset.y, Color, 0, , True

                    ' Sombra
                    PresentComposedTexture PixelOffsetX + .Body.BodyOffset.x, PixelOffsetY + .Body.BodyOffset.y, Color, 0, True
                    If LenB(.Body_Aura) <> 0 And .Body_Aura <> "0" Then Call Renderizar_Aura(.Body_Aura, PixelOffsetX + .Body.BodyOffset.x, PixelOffsetY + OffArma + .Body.BodyOffset.y, x, y, CharIndex)
                    If LenB(.Head_Aura) <> 0 And .Head_Aura <> "0" Then Call Renderizar_Aura(.Head_Aura, PixelOffsetX + .Body.BodyOffset.x, PixelOffsetY + OffArma + .Body.BodyOffset.y, x, y, CharIndex)
                    If LenB(.Arma_Aura) <> 0 And .Arma_Aura <> "0" Then Call Renderizar_Aura(.Arma_Aura, PixelOffsetX + .Body.BodyOffset.x, PixelOffsetY + OffArma + .Body.BodyOffset.y, x, y, CharIndex)
                    If LenB(.Otra_Aura) <> 0 And .Otra_Aura <> "0" Then Call Renderizar_Aura(.Otra_Aura, PixelOffsetX + .Body.BodyOffset.x, PixelOffsetY + OffArma + .Body.BodyOffset.y, x, y, CharIndex)
                    If LenB(.Escudo_Aura) <> 0 And .Escudo_Aura <> "0" Then Call Renderizar_Aura(.Escudo_Aura, PixelOffsetX + .Body.BodyOffset.x, PixelOffsetY + OffArma + .Body.BodyOffset.y, x, y, CharIndex)
                    If LenB(.DM_Aura) <> 0 And .DM_Aura <> "0" Then Call Renderizar_Aura(.DM_Aura, PixelOffsetX + .Body.BodyOffset.x, PixelOffsetY + OffArma + .Body.BodyOffset.y, x, y, CharIndex)
                    If LenB(.RM_Aura) <> 0 And .RM_Aura <> "0" Then Call Renderizar_Aura(.RM_Aura, PixelOffsetX + .Body.BodyOffset.x, PixelOffsetY + OffArma + .Body.BodyOffset.y, x, y, CharIndex)
                End If

                ' Char
                If CharindexSeguido > 0 And .Invisible And CharindexSeguido <> CharIndex Then
                    Call SetRGBA(Color(0), 180, 180, 180, 255)
                    Call SetRGBA(Color(1), 180, 180, 180, 255)
                    Call SetRGBA(Color(2), 180, 180, 180, 255)
                    Call SetRGBA(Color(3), 180, 180, 180, 255)
                    Call SetRGBA(NameColor(0), 180, 180, 180, 255)
                    Call SetRGBA(NameColor(1), 180, 180, 180, 255)
                    Call SetRGBA(NameColor(2), 180, 180, 180, 255)
                    Call SetRGBA(NameColor(3), 180, 180, 180, 255)
                    line = .nombre
                    Engine_Text_Render line, PixelOffsetX + 16 - CInt(Engine_Text_Width(line, True) / 2) + .Body.BodyOffset.x, PixelOffsetY + .Body.BodyOffset.y + 30 + OffsetYname - Engine_Text_Height(line, True), NameColor, 1, False, 0, 255
                End If
                
                PresentComposedTexture PixelOffsetX + .Body.BodyOffset.x, PixelOffsetY + .Body.BodyOffset.y, Color, False
                ' we need to draw this outside the composed texture or it will be cut out
                If .Heading = E_Heading.NORTH Then
                    If .HasCart Then Call Draw_Grh(.Cart.Walk(.Heading), PixelOffsetX + .Body.BodyOffset.x, PixelOffsetY + .Body.BodyOffset.y + .Cart.HeadOffset.y, 1, 1, Color, False, x, y)
                End If
            ' Si no, solo dibujamos body
            Else
                If Not .Invisible Then
                    Call Draw_Sombra(.Body.Walk(.Heading), PixelOffsetX + .Body.BodyOffset.x, PixelOffsetY + .Body.BodyOffset.y, 1, 1, False, x, y)
                End If
                Call Draw_Grh(.Body.Walk(.Heading), PixelOffsetX + .Body.BodyOffset.x, PixelOffsetY + .Body.BodyOffset.y, 1, 1, Color, False, x, y)
            End If
    
            'Draw name over head
            Nombres = Not MapData(charlist(CharIndex).Pos.x, charlist(CharIndex).Pos.y).zone.OcultarNombre
     
                
            If UserCharIndex > 0 Then
             With charlist(UserCharIndex)
                 Dim new_music As Integer
                 new_music = MapData(.Pos.x, .Pos.y).zone.Musica
                 
                 If new_music > 0 Then
                     Call ao20audio.PlayMidi(new_music, True)
                 Else
                     
                     Call ao20audio.PlayMidi(MapDat.music_numberLow, True)
                 End If
             End With
            End If
    
            If Nombres And Len(.nombre) > 0 And MostrarNombre Then
                Pos = InStr(.nombre, "<")
                If Pos = 0 Then Pos = InStr(.nombre, "[")
                If Pos = 0 Then Pos = Len(.nombre) + 2
                'Nick
                
                line = Left$(.nombre, Pos - 2)
                Dim Factor As Double
                Factor = MapData(x, y).light_value(0).r / 255
                
                If .Navegando Then
                   
                     If .priv = 0 Then
                        
                        Select Case .status
                            ' Criminal
                            Case 0
                                Call RGBAList(NameColor, ColoresPJ(23).r, ColoresPJ(23).G, ColoresPJ(23).b)
                                Call RGBAList(colorCorazon, ColoresPJ(23).r, ColoresPJ(23).G, ColoresPJ(23).b)
                            ' Ciudadano
                            Case 1
                                Call RGBAList(NameColor, ColoresPJ(20).r, ColoresPJ(20).G, ColoresPJ(20).b)
                                Call RGBAList(colorCorazon, ColoresPJ(20).r, ColoresPJ(20).G, ColoresPJ(20).b)
                            
                            ' Caos
                            Case 2
                                Call RGBAList(NameColor, ColoresPJ(24).r, ColoresPJ(24).G, ColoresPJ(24).b)
                                Call RGBAList(colorCorazon, ColoresPJ(24).r, ColoresPJ(24).G, ColoresPJ(24).b)
    
                            ' Armada
                            Case 3
                                Call RGBAList(NameColor, ColoresPJ(21).r, ColoresPJ(21).G, ColoresPJ(21).b)
                                Call RGBAList(colorCorazon, ColoresPJ(21).r, ColoresPJ(21).G, ColoresPJ(21).b)
                                
                            ' Concilio
                            Case 4
                                Call RGBAList(NameColor, ColoresPJ(25).r, ColoresPJ(25).G, ColoresPJ(25).b)
                                Call RGBAList(colorCorazon, ColoresPJ(25).r, ColoresPJ(25).G, ColoresPJ(25).b)
                            ' Consejo
                            Case 5
                                Call RGBAList(NameColor, ColoresPJ(22).r, ColoresPJ(22).G, ColoresPJ(22).b)
                                Call RGBAList(colorCorazon, ColoresPJ(22).r, ColoresPJ(22).G, ColoresPJ(22).b)
    
                        End Select
                    Else
                        Call RGBAList(NameColor, ColoresPJ(.priv).r, ColoresPJ(.priv).G, ColoresPJ(.priv).b)
                        Call RGBAList(colorCorazon, ColoresPJ(.priv).r, ColoresPJ(.priv).G, ColoresPJ(.priv).b)
                    End If
                Else
                    NameColor(0).r = NameColor(0).r * Factor
                    NameColor(0).G = NameColor(0).G * Factor
                    NameColor(0).b = NameColor(0).b * Factor
                    NameColor(1).r = NameColor(1).r * Factor
                    NameColor(1).G = NameColor(1).G * Factor
                    NameColor(1).b = NameColor(1).b * Factor
                    NameColor(2).r = NameColor(2).r * Factor
                    NameColor(2).G = NameColor(2).G * Factor
                    NameColor(2).b = NameColor(2).b * Factor
                    NameColor(3).r = NameColor(3).r * Factor
                    NameColor(3).G = NameColor(3).G * Factor
                    NameColor(3).b = NameColor(3).b * Factor
                
                End If
                
                If .Team > 0 Then
                    Call RGBAList(NameColor, TeamColors(.Team).r, TeamColors(.Team).G, TeamColors(.Team).b, TeamColors(.Team).a)
                End If
                Engine_Text_Render line, PixelOffsetX + 16 - CInt(Engine_Text_Width(line, True) / 2) + .Body.BodyOffset.x, PixelOffsetY + .Body.BodyOffset.y + 30 + OffsetYname - Engine_Text_Height(line, True), NameColor, 1, False, 0, IIf(.Invisible, 160, 255)

                'Clan
                
                If .priv = 2 Or .priv = 3 Or .priv = 4 Then
                    line = "<Game Master>"
                ElseIf .priv = 5 Then
                    line = JsonLanguage.Item("MENSAJE_543")
                Else
                    line = .clan
                End If
                
                If .Team > 0 Then
                    line = JsonLanguage.Item("MENSAJE_544") & .Team & ">"
                End If
                
                If .banderaIndex > 0 And .Team > 0 Then
                    Dim flag As grh
                    If .banderaIndex = 1 Then
                        Call InitGrh(flag, 58712)
                    ElseIf .banderaIndex = 2 Then
                        Call InitGrh(flag, 60298)
                    End If
                    Call Draw_Grh(flag, PixelOffsetX + 1 + .Body.BodyOffset.x, PixelOffsetY - 45 + .Body.BodyOffset.y, 1, 0, Color, True, 0, 0, 0)
                End If
                
                If (UserCharIndex > LBound(charlist) And UserCharIndex < UBound(charlist)) Then
                    
                    If (.clan_index = charlist(UserCharIndex).clan_index And CharIndex <> UserCharIndex And .EsNpc = False And .Team <= 0) Or (CharIndex = UserCharIndex And .Invisible) Then
                        'Seteo color de nombre del clan solo si es de mi clan
                        Call SetRGBA(NameColorClan(0), 255, 255, 0, 255)
                        Call SetRGBA(NameColorClan(1), 255, 255, 0, 255)
                        Call SetRGBA(NameColorClan(2), 255, 255, 0, 255)
                        Call SetRGBA(NameColorClan(3), 255, 255, 0, 255)
                        Engine_Text_Render line, PixelOffsetX + 16 - CInt(Engine_Text_Width(line, True) / 2) + .Body.BodyOffset.x, PixelOffsetY + .Body.BodyOffset.y + 42 + OffsetYClan - Engine_Text_Height(line, True), NameColorClan, 1, False, 0, IIf(.Invisible, 160, 255)
                    Else
                        Engine_Text_Render line, PixelOffsetX + 16 - CInt(Engine_Text_Width(line, True) / 2) + .Body.BodyOffset.x, PixelOffsetY + .Body.BodyOffset.y + 42 + OffsetYClan - Engine_Text_Height(line, True), NameColor, 1, False, 0, IIf(.Invisible, 160, 255)
                    End If
                End If
            ElseIf Nombres And .Team > 0 Then
                line = JsonLanguage.Item("MENSAJE_544") & .Team & ">"
                Call RGBAList(NameColor, TeamColors(.Team).r, TeamColors(.Team).G, TeamColors(.Team).b, TeamColors(.Team).a)
                Engine_Text_Render line, PixelOffsetX + 16 - CInt(Engine_Text_Width(line, True) / 2) + .Body.BodyOffset.x, PixelOffsetY + .Body.BodyOffset.y + 30 + OffsetYname - Engine_Text_Height(line, True), NameColor, 1, False, 0, IIf(.Invisible, 160, 255)
            End If
        End If

        If .particle_count > 0 Then

            For i = 1 To .particle_count

                If .particle_group(i) > 0 Then
                    Particle_Group_Render .particle_group(i), PixelOffsetX + .Body.BodyOffset.x + (TilePixelWidth / 2), PixelOffsetY + .Body.BodyOffset.y
                End If

            Next i

        End If
        If Nombres And Len(.nombre) > 0 And MostrarNombre And .tipoUsuario > 0 Then
            Select Case .tipoUsuario
                Case eTipoUsuario.aventurero
                    Call RGBAList(color, 0, 255, 0, IIf(.Invisible, 120, 255))
                Case eTipoUsuario.heroe
                    Call RGBAList(color, 255, 0, 0, IIf(.Invisible, 160, 255))
                Case eTipoUsuario.Legend
                    Call RGBAList(color, 255, 255, 0, IIf(.Invisible, 120, 255))
            End Select

            Dim txt_width As Long
            txt_width = Engine_Text_Width(.nombre, True)
            Call Draw_Grh(StarGrh, PixelOffsetX + 1 + .Body.BodyOffset.x + (txt_width / 2) + 8, PixelOffsetY + 20 + .Body.BodyOffset.y, 1, 1, color, False, 0, 0, 0)
        End If
        
        'Barra de tiempo
        If .BarTime < .MaxBarTime And Not .Invisible Then
            Call InitGrh(TempGrh, 839)
            Call RGBAList(Color, 255, 255, 255, 200)

            Call Draw_Grh(TempGrh, PixelOffsetX + 1 + .Body.BodyOffset.x, PixelOffsetY - 55 + .Body.BodyOffset.y, 1, 0, Color, False, 0, 0, 0)
            
            Engine_Draw_Box PixelOffsetX + 5 + .Body.BodyOffset.x, PixelOffsetY - 28 + .Body.BodyOffset.y, .BarTime / .MaxBarTime * 26, 4, RGBA_From_Comp(3, 214, 166, 120) ', RGBA_From_Comp(0, 0, 0, 255)

            .BarTime = .BarTime + (timerElapsedTime / 1000)
            'frmdebug.add_text_tracebox .BarTime
            If .BarTime >= .MaxBarTime Then
                charlist(CharIndex).BarTime = 0
                charlist(CharIndex).BarAccion = 99
                charlist(CharIndex).MaxBarTime = 0
            End If

        End If
                   
        ' Meditación
        If .ActiveAnimation.PlaybackState <> Stopped Then
            Call UpdateAnimation(.ActiveAnimation)
            Call Draw_Animation(.ActiveAnimation, PixelOffsetX + .Body.BodyOffset.x, PixelOffsetY + 4 + .Body.BodyOffset.y, 1, Color)
        End If

        If .FxCount > 0 Then

            For i = 1 To .FxCount

                If .FxList(i).FxIndex > 0 And .FxList(i).started <> 0 Then
                    Call RGBAList(Color, 255, 255, 255, 220)

                    If FxData(.FxList(i).FxIndex).IsPNG = 1 Then
                        Call Draw_GrhFX(.FxList(i), PixelOffsetX + FxData(.FxList(i).FxIndex).OffsetX + .Body.BodyOffset.x, PixelOffsetY + FxData(.FxList(i).FxIndex).OffsetY + 20 + .Body.BodyOffset.y, 1, 1, Color, False, , , , CharIndex)
                    Else
                        Call Draw_GrhFX(.FxList(i), PixelOffsetX + FxData(.FxList(i).FxIndex).OffsetX + .Body.BodyOffset.x, PixelOffsetY + FxData(.FxList(i).FxIndex).OffsetY + 20 + .Body.BodyOffset.y, 1, 1, Color, True, , , , CharIndex)
                    End If

                End If

                If .FxList(i).started = 0 Then
                    .FxList(i).FxIndex = 0

                End If

            Next i

            If .FxList(.FxCount).started = 0 Then
                .FxCount = .FxCount - 1

            End If

        End If

    End With

    
    Exit Sub

Char_Render_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Char_Render", Erl)
    Resume Next
    
End Sub
Public Sub DibujarVidaChar(ByVal CharIndex As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer, ByRef OffsetYname As Byte, ByRef OffsetYClan As Byte)
    With charlist(CharIndex)
        Engine_Draw_Box PixelOffsetX + .Body.BodyOffset.x, PixelOffsetY + 33 + .Body.BodyOffset.y, 33, 5, RGBA_From_Comp(10, 10, 10)
        If .UserMaxHp <> 0 Then
            Dim FullSize As Long
            FullSize = .UserMaxHp + .Shield
            Engine_Draw_Box PixelOffsetX + 1 + .Body.BodyOffset.x, PixelOffsetY + 34 + .Body.BodyOffset.y, .UserMinHp / FullSize * 31, 3, RGBA_From_Comp(255, 0, 0)
            Engine_Draw_Box PixelOffsetX + 1 + .Body.BodyOffset.x + .UserMinHp / FullSize * 31, PixelOffsetY + 34 + .Body.BodyOffset.y, .Shield / FullSize * 31, 3, RGBA_From_Comp(162, 108, 16)
        Else
            Engine_Draw_Box PixelOffsetX + 1 + .Body.BodyOffset.x, PixelOffsetY + 34 + .Body.BodyOffset.y, 31, 4, RGBA_From_Comp(255, 0, 0)
        End If
    
        If .UserMaxMAN <> 0 Then
            OffsetYname = 12
            OffsetYClan = 12
    
            Engine_Draw_Box PixelOffsetX + .Body.BodyOffset.x, PixelOffsetY + 38 + .Body.BodyOffset.y, 33, 4, RGBA_From_Comp(10, 10, 10)
            Engine_Draw_Box PixelOffsetX + 1 + .Body.BodyOffset.y, PixelOffsetY + 38 + .Body.BodyOffset.y, .UserMinMAN / .UserMaxMAN * 31, 3, RGBA_From_Comp(0, 100, 255)
        End If
    End With
End Sub

Public Function IsCharVisible(ByVal CharIndex As Integer) As Boolean

    
        If CharIndex = UserCharIndex Then
            IsCharVisible = True
            Exit Function
        End If
        
        If charlist(UserCharIndex).priv > 0 And charlist(CharIndex).priv <= charlist(UserCharIndex).priv Then
            IsCharVisible = True
            Exit Function
        End If


End Function

#If REMOTE_CLOSE = 1 Then
Public Sub bot_main_loop()
On Error GoTo Start_Err
    Call LoadFonts
    prgRun = True
    InitiateShutdownProcess = False
    Dim countdown As Integer
    countdown = 60
    Do While prgRun
        DoEvents
        Call modNetwork.Poll
        DoEvents
        If InitiateShutdownProcess Then
            If countdown > 0 And ShutdownProcessTimer.ElapsedSeconds > 5 Then
                ShutdownProcessTimer.Start
                Call WriteGlobalMessage("WARNING: The server is going to close in " & countdown & " seconds for scheduled maintainance. Please disconnect from the game!!!")
                frmDebug.add_text_tracebox "WARNING: The server is going to close in " & countdown & " seconds for scheduled maintainance. Please disconnect from the game!!!"
                Call SaveStringInFile("WARNING: The server is going to close in " & countdown & " seconds for scheduled maintainance. Please disconnect from the game!!!", "remote_debug.txt")
                countdown = countdown - 5
            ElseIf countdown = 0 Then
                Call WriteFeatureEnable("SGRACEFULLY", 1)
                Call SaveStringInFile("Sent request to shutdown to the server", "remote_debug.txt")
                countdown = -1
            ElseIf countdown < 0 Then

            End If
            
        End If
    Loop

    Exit Sub

Start_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Start", Erl)
    Resume Next
    
End Sub
#End If

Public Sub Start()
On Error GoTo Start_Err
    DoEvents
    Do While prgRun
        If GetGameplayForm().WindowState <> vbMinimized Then
             Select Case g_game_state.State()
                Case e_state_gameplay_screen
                    render
                    Check_Keys
                    Moviendose = False
                    DrawMainInventory


                    If frmComerciar.visible Then
                        DrawInterfaceComerciar

                    ElseIf frmBancoObj.visible Then
                        DrawInterfaceBoveda

                    ElseIf frmBancoCuenta.visible Then
                        DrawInterfaceBovedaCuenta

                    ElseIf frmCrafteo.visible Then
                        DrawInterfaceCrafting
                    End If

                    If FrmKeyInv.visible Then
                        DrawInterfaceKeys
                    End If
                    
                    If frmComerciarUsu.visible Then
                        DrawInventoryComercio
                        DrawInventoryUserComercio
                        DrawInventoryOtherComercio
                    End If

                Case e_state_connect_screen
#If DXUI Then
                    If Not frmConnect.visible Then
                            Call ShowLogin
                            FrmLogear.Hide
                    End If
                    Debug.Assert Not g_connectScreen Is Nothing
                    g_connectScreen.render DirectDevice
                   
                    ' Update mouse coords and button state
                    UpdateMouse frmConnect.render.hWnd
                    g_MouseButtons = GetAsyncKeyState(VK_LBUTTON) And &H8000 ' left button state

                    ' Pass movement and clicks to UI
                    g_connectScreen.HandleInput g_MouseX, g_MouseY, g_MouseButtons
                    If g_connectScreen.WasConnectClicked Then
                        Debug.Print "g_connectScreen.WasConnectClicked"
                        'Call SetActiveServer(txtIp.text, txtPort.text)
                        'Call DoLogin(NameTxt.text, PasswordTxt.text, chkRecordar.Tag = "1")
                    End If
                    
                    
                    
#Else
                    If Not frmConnect.visible Then
                            Call ShowLogin
                    End If
                    RenderConnect 57, 45, 0, 0
#End If
                Case e_state_account_screen
                    rendercuenta 42, 43, 0, 0

                Case e_state_createchar_screen
                    RenderCrearPJ 76, 82, 0, 0
            End Select

        Else
        #If DIRECT_PLAY = 0 Then
            Sleep 60&
        #End If
            Call frmMain.Inventario.ReDraw
        End If

        DoEvents

        Call modNetwork.Poll
        Call svb_run_callbacks
        Call UpdateAntiCheat
    Loop

    EngineRun = False

    Call CloseClient

    
    Exit Sub

Start_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Start", Erl)
    Resume Next
    
End Sub

Public Sub SetMapFx(ByVal x As Byte, ByVal y As Byte, ByVal Fx As Integer, ByVal Loops As Integer)

    'Sets an FX to the character.
    
    On Error GoTo SetMapFx_Err
    
    Dim indice As Byte

    With MapData(x, y)
    
        indice = Map_FX_Group_Next_Open(x, y)
    
        .FxList(indice).FxIndex = Fx
        Call InitGrh(.FxList(indice), FxData(Fx).Animacion)
        .FxList(indice).Loops = Loops
    
    End With

    
    Exit Sub

SetMapFx_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.SetMapFx", Erl)
    Resume Next
    
End Sub

Private Function Engine_FToDW(f As Single) As Long
    
    On Error GoTo Engine_FToDW_Err
    

    ' single > long
    Dim buf As D3DXBuffer

    Set buf = DirectD3D8.CreateBuffer(4)
    DirectD3D8.BufferSetData buf, 0, 4, 1, f
    DirectD3D8.BufferGetData buf, 0, 4, 1, Engine_FToDW

    
    Exit Function

Engine_FToDW_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_FToDW", Erl)
    Resume Next
    
End Function

Private Function VectorToRGBA(vec As D3DVECTOR, fHeight As Single) As Long
    
    On Error GoTo VectorToRGBA_Err
    

    Dim r As Integer, G As Integer, b As Integer, a As Integer

    r = 127 * vec.x + 128
    G = 127 * vec.y + 128
    b = 127 * vec.z + 128
    a = 255 * fHeight
    VectorToRGBA = D3DColorARGB(a, r, G, b)

    
    Exit Function

VectorToRGBA_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.VectorToRGBA", Erl)
    Resume Next
    
End Function

Public Sub DrawMainInventory()
    
    On Error GoTo DrawMainInventory_Err
    

    ' Sólo dibujamos cuando es necesario
    'If Not frmMain.Inventario.NeedsRedraw Then Exit Sub

    Dim InvRect As RECT

    InvRect.Left = 0
    InvRect.Top = 0
    InvRect.Right = frmMain.picInv.ScaleWidth
    InvRect.Bottom = frmMain.picInv.ScaleHeight

    ' Comenzamos la escena
    Call Engine_BeginScene

    ' Dibujamos el fondo del inventario principal
    'Call Draw_GrhIndex(6, 0, 0)

    ' Dibujamos items
    Call frmMain.Inventario.DrawInventory
    
    ' Dibujamos item arrastrado
    Call frmMain.Inventario.DrawDraggedItem

    ' Presentamos la escena
    Call Engine_EndScene(InvRect, frmMain.picInv.hwnd)

    
    Exit Sub

DrawMainInventory_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.DrawMainInventory", Erl)
    Resume Next
    
End Sub

Public Sub DrawInterfaceComerciar()
    
    On Error GoTo DrawInterfaceComerciar_Err
    

    ' Sólo dibujamos cuando es necesario
    If Not frmComerciar.InvComNpc.NeedsRedraw And Not frmComerciar.InvComUsu.NeedsRedraw Then Exit Sub

    Dim InvRect As RECT

    InvRect.Left = 0
    InvRect.Top = 0
    InvRect.Right = frmComerciar.interface.ScaleWidth
    InvRect.Bottom = frmComerciar.interface.ScaleHeight
    RenderCullingRect = InvRect
    ' Comenzamos la escena
    Call Engine_BeginScene

    ' Dibujamos el fondo del inventario de comercio
    Call Draw_GrhIndex(837, 0, 0)

    ' Dibujamos items del NPC
    Call frmComerciar.InvComNpc.DrawInventory
    
    ' Dibujamos items del usuario
    Call frmComerciar.InvComUsu.DrawInventory

    ' Dibujamos "ambos" items arrastrados (aunque sólo puede estar uno activo a la vez)
    Call frmComerciar.InvComNpc.DrawDraggedItem
    Call frmComerciar.InvComUsu.DrawDraggedItem
    
    ' Me fijo qué inventario está seleccionado
    Dim CurrentInventory As clsGrapchicalInventory
    
    Dim cantidad         As Integer

    If frmComerciar.InvComNpc.SelectedItem > 0 Then
        Set CurrentInventory = frmComerciar.InvComNpc
        ' Al comprar, calculamos el valor según la cantidad exacta que ingresó
        cantidad = Val(frmComerciar.cantidad.Text)
    ElseIf frmComerciar.InvComUsu.SelectedItem > 0 Then
        Set CurrentInventory = frmComerciar.InvComUsu
        ' Al vender, calculamos el valor según el min(cantidad_ingresada, cantidad_items)
        cantidad = min(Val(frmComerciar.cantidad.Text), CurrentInventory.Amount(CurrentInventory.SelectedItem))

    End If
    
    ' Si hay alguno seleccionado
    If Not CurrentInventory Is Nothing Then
        ' Dibujo el item seleccionado
        'Call Draw_GrhColor(CurrentInventory.GrhIndex(CurrentInventory.SelectedItem), 282, 251, COLOR_WHITE)
    
        ' Muestro info del item
        Dim str As String

        str = JsonLanguage.Item("MENSAJE_545")
        
        Select Case CurrentInventory.PuedeUsar(CurrentInventory.SelectedItem)

            Case 1
                str = str & JsonLanguage.Item("MENSAJE_546")

            Case 2
                str = str & JsonLanguage.Item("MENSAJE_547")

            Case 3
                str = str & JsonLanguage.Item("MENSAJE_548")

            Case 4
                str = str & JsonLanguage.Item("MENSAJE_549")

            Case 5
                str = str & JsonLanguage.Item("MENSAJE_550")

            Case 6
                str = str & JsonLanguage.Item("MENSAJE_551")

            Case 0
                str = JsonLanguage.Item("MENSAJE_552")

        End Select
                           
        frmComerciar.lblnombre = CurrentInventory.ItemName(CurrentInventory.SelectedItem) & str
        frmComerciar.lbldesc = CurrentInventory.GetInfo(CurrentInventory.ObjIndex(CurrentInventory.SelectedItem))
        frmComerciar.lblcosto = PonerPuntos(Fix(CurrentInventory.Valor(CurrentInventory.SelectedItem) * cantidad))
        
        Set CurrentInventory = Nothing

    End If

    ' Presentamos la escena
    Call Engine_EndScene(InvRect, frmComerciar.interface.hwnd)
    RenderCullingRect = GameplayDrawAreaRect
    
    Exit Sub

DrawInterfaceComerciar_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.DrawInterfaceComerciar", Erl)
    Resume Next
    
End Sub

Public Sub DrawInterfaceBovedaCuenta()
    
    On Error GoTo DrawInterfaceBoveda_Err
    

    ' Sólo dibujamos cuando es necesario
    If Not frmBancoCuenta.InvBovedaCuenta.NeedsRedraw And Not frmBancoCuenta.InvBankUsuCuenta.NeedsRedraw Then Exit Sub

    Dim InvRect As RECT

    InvRect.Left = 0
    InvRect.Top = 0
    InvRect.Right = frmBancoCuenta.interface.ScaleWidth
    InvRect.Bottom = frmBancoCuenta.interface.ScaleHeight
    RenderCullingRect = InvRect
    ' Comenzamos la escena
    Call Engine_BeginScene

    ' Dibujamos el fondo de la bóveda
    Call Draw_GrhIndex(838, 0, 0)

    ' Dibujamos items de la bóveda
    Call frmBancoCuenta.InvBovedaCuenta.DrawInventory
    
    ' Dibujamos items del usuario
    Call frmBancoCuenta.InvBankUsuCuenta.DrawInventory

    ' Dibujamos "ambos" items arrastrados (aunque sólo puede estar uno activo a la vez)
    Call frmBancoCuenta.InvBovedaCuenta.DrawDraggedItem
    Call frmBancoCuenta.InvBankUsuCuenta.DrawDraggedItem
    
    ' Me fijo qué inventario está seleccionado
    Dim CurrentInventory As clsGrapchicalInventory
    
    If frmBancoCuenta.InvBovedaCuenta.SelectedItem > 0 Then
        Set CurrentInventory = frmBancoCuenta.InvBovedaCuenta
    ElseIf frmBancoCuenta.InvBankUsuCuenta.SelectedItem > 0 Then
        Set CurrentInventory = frmBancoCuenta.InvBankUsuCuenta

    End If
    
    ' Si hay alguno seleccionado
    If Not CurrentInventory Is Nothing Then

        ' Muestro info del item
        Dim str As String

        str = JsonLanguage.Item("MENSAJE_545")
        
        Select Case CurrentInventory.PuedeUsar(CurrentInventory.SelectedItem)

            Case 1
                str = str & JsonLanguage.Item("MENSAJE_546")

            Case 2
                str = str & JsonLanguage.Item("MENSAJE_547")

            Case 3
                str = str & JsonLanguage.Item("MENSAJE_548")

            Case 4
                str = str & JsonLanguage.Item("MENSAJE_549")

            Case 5
                str = str & JsonLanguage.Item("MENSAJE_550")

            Case 6
                str = str & JsonLanguage.Item("MENSAJE_551")

            Case 0
                str = JsonLanguage.Item("MENSAJE_552")

        End Select
        
        frmBancoCuenta.lblnombre.Caption = CurrentInventory.ItemName(CurrentInventory.SelectedItem) & str
        frmBancoCuenta.lbldesc.Caption = CurrentInventory.GetInfo(CurrentInventory.ObjIndex(CurrentInventory.SelectedItem))
        
        Set CurrentInventory = Nothing

    End If

    ' Presentamos la escena
    Call Engine_EndScene(InvRect, frmBancoCuenta.interface.hwnd)
    RenderCullingRect = GameplayDrawAreaRect
    
    Exit Sub

DrawInterfaceBoveda_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.DrawInterfaceBoveda", Erl)
    Resume Next
    
End Sub

Public Sub DrawInterfaceBoveda()
    
    On Error GoTo DrawInterfaceBoveda_Err
    

    ' Sólo dibujamos cuando es necesario
    If Not frmBancoObj.InvBoveda.NeedsRedraw And Not frmBancoObj.InvBankUsu.NeedsRedraw Then Exit Sub

    Dim InvRect As RECT

    InvRect.Left = 0
    InvRect.Top = 0
    InvRect.Right = frmBancoObj.interface.ScaleWidth
    InvRect.Bottom = frmBancoObj.interface.ScaleHeight
    RenderCullingRect = InvRect
    ' Comenzamos la escena
    Call Engine_BeginScene

    ' Dibujamos el fondo de la bóveda
    Call Draw_GrhIndex(838, 0, 0)

    ' Dibujamos items de la bóveda
    Call frmBancoObj.InvBoveda.DrawInventory
    
    ' Dibujamos items del usuario
    Call frmBancoObj.InvBankUsu.DrawInventory

    ' Dibujamos "ambos" items arrastrados (aunque sólo puede estar uno activo a la vez)
    Call frmBancoObj.InvBoveda.DrawDraggedItem
    Call frmBancoObj.InvBankUsu.DrawDraggedItem
    
    ' Me fijo qué inventario está seleccionado
    Dim CurrentInventory As clsGrapchicalInventory
    
    If frmBancoObj.InvBoveda.SelectedItem > 0 Then
        Set CurrentInventory = frmBancoObj.InvBoveda
    ElseIf frmBancoObj.InvBankUsu.SelectedItem > 0 Then
        Set CurrentInventory = frmBancoObj.InvBankUsu

    End If
    
    ' Si hay alguno seleccionado
    If Not CurrentInventory Is Nothing Then

        ' Muestro info del item
        Dim str As String

        str = JsonLanguage.Item("MENSAJE_545")
        
        Select Case CurrentInventory.PuedeUsar(CurrentInventory.SelectedItem)

            Case 1
                str = str & JsonLanguage.Item("MENSAJE_546")

            Case 2
                str = str & JsonLanguage.Item("MENSAJE_547")

            Case 3
                str = str & JsonLanguage.Item("MENSAJE_548")

            Case 4
                str = str & JsonLanguage.Item("MENSAJE_549")

            Case 5
                str = str & JsonLanguage.Item("MENSAJE_550")

            Case 6
                str = str & JsonLanguage.Item("MENSAJE_551")

            Case 0
                str = JsonLanguage.Item("MENSAJE_552")

        End Select
        
        frmBancoObj.lblnombre.Caption = CurrentInventory.ItemName(CurrentInventory.SelectedItem) & str
        frmBancoObj.lbldesc.Caption = CurrentInventory.GetInfo(CurrentInventory.ObjIndex(CurrentInventory.SelectedItem))
        
        Set CurrentInventory = Nothing

    End If

    ' Presentamos la escena
    Call Engine_EndScene(InvRect, frmBancoObj.interface.hwnd)
    RenderCullingRect = GameplayDrawAreaRect
    
    Exit Sub

DrawInterfaceBoveda_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.DrawInterfaceBoveda", Erl)
    Resume Next
    
End Sub

Public Sub DrawInterfaceKeys()
    
    On Error GoTo DrawInterfaceKeys_Err
    

    ' Sólo dibujamos cuando es necesario
    If Not FrmKeyInv.InvKeys.NeedsRedraw Then Exit Sub

    Dim InvRect As RECT

    InvRect.Left = 0
    InvRect.Top = 0
    InvRect.Right = FrmKeyInv.interface.ScaleWidth
    InvRect.Bottom = FrmKeyInv.interface.ScaleHeight
    RenderCullingRect = InvRect
    ' Comenzamos la escena
    Call Engine_BeginScene

    ' Dibujamos el fondo de la bóveda
    'Call Draw_GrhIndex(838, 0, 0)
    
    ' Dibujamos llaves
    Call FrmKeyInv.InvKeys.DrawInventory

    ' Presentamos la escena
    Call Engine_EndScene(InvRect, FrmKeyInv.interface.hwnd)
    RenderCullingRect = GameplayDrawAreaRect
    
    Exit Sub

DrawInterfaceKeys_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.DrawInterfaceKeys", Erl)
    Resume Next
    
End Sub

Public Sub DrawInventoryComercio()
    
    On Error GoTo DrawInventorysComercio_Err
    

    ' Sólo dibujamos cuando es necesario
    If Not frmComerciarUsu.InvUser.NeedsRedraw Then Exit Sub

    Dim InvRect As RECT

    InvRect.Left = 0
    InvRect.Top = 0
    InvRect.Right = frmComerciarUsu.picInv.ScaleWidth
    InvRect.Bottom = frmComerciarUsu.picInv.ScaleHeight
    RenderCullingRect = InvRect
    ' Comenzamos la escena
    Call Engine_BeginScene
    
    ' Dibujamos llaves
    Call frmComerciarUsu.InvUser.DrawInventory
    
    ' Presentamos la escena
    Call Engine_EndScene(InvRect, frmComerciarUsu.picInv.hwnd)

    RenderCullingRect = GameplayDrawAreaRect
    Exit Sub

DrawInventorysComercio_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.DrawInventorysComercio", Erl)
    Resume Next
    
End Sub

Public Sub DrawInventoryUserComercio()
    
    On Error GoTo DrawInventoryUserComercio_Err
    

    ' Sólo dibujamos cuando es necesario
    If Not frmComerciarUsu.InvUserSell.NeedsRedraw Then Exit Sub

    Dim InvRect As RECT

    InvRect.Left = 0
    InvRect.Top = 0
    InvRect.Right = frmComerciarUsu.picInvUserSell.ScaleWidth
    InvRect.Bottom = frmComerciarUsu.picInvUserSell.ScaleHeight
    RenderCullingRect = InvRect
    ' Comenzamos la escena
    Call Engine_BeginScene
    
    ' Dibujamos llaves
    Call frmComerciarUsu.InvUserSell.DrawInventory
    
    ' Presentamos la escena
    Call Engine_EndScene(InvRect, frmComerciarUsu.picInvUserSell.hwnd)

    RenderCullingRect = GameplayDrawAreaRect
    Exit Sub

DrawInventoryUserComercio_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.DrawInventoryUserComercio", Erl)
    Resume Next
    
End Sub

Public Sub DrawInventoryOtherComercio()
    
    On Error GoTo DrawInventoryOtherComercio_Err
    

    ' Sólo dibujamos cuando es necesario
    If Not frmComerciarUsu.InvOtherSell.NeedsRedraw Then Exit Sub

    Dim InvRect As RECT

    InvRect.Left = 0
    InvRect.Top = 0
    InvRect.Right = frmComerciarUsu.picInvOtherSell.ScaleWidth
    InvRect.Bottom = frmComerciarUsu.picInvOtherSell.ScaleHeight
    RenderCullingRect = InvRect
    ' Comenzamos la escena
    Call Engine_BeginScene
    ' Dibujamos llaves
    Call frmComerciarUsu.InvOtherSell.DrawInventory
    ' Presentamos la escena
    Call Engine_EndScene(InvRect, frmComerciarUsu.picInvOtherSell.hwnd)
    RenderCullingRect = GameplayDrawAreaRect
    
    Exit Sub

DrawInventoryOtherComercio_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.DrawInventoryOtherComercio", Erl)
    Resume Next
    
End Sub

Public Sub DrawInterfaceCrafting()
    
    On Error GoTo DrawInterfaceBoveda_Err

    ' Sólo dibujamos cuando es necesario
    If Not frmCrafteo.InvCraftUser.NeedsRedraw And Not frmCrafteo.InvCraftItems.NeedsRedraw And Not frmCrafteo.InvCraftCatalyst.NeedsRedraw Then Exit Sub

    Dim InvRect As RECT
    InvRect.Left = 0
    InvRect.Top = 0
    InvRect.Right = frmCrafteo.PicInven.ScaleWidth
    InvRect.Bottom = frmCrafteo.PicInven.ScaleHeight
    RenderCullingRect = InvRect
    ' Comenzamos la escena
    Call Engine_BeginScene

    ' Dibujamos el fondo
    Call Draw_GrhIndex(frmCrafteo.InventoryGrhIndex, 0, 0)

    ' Dibujamos los inventarios
    Call frmCrafteo.InvCraftUser.DrawInventory
    Call frmCrafteo.InvCraftItems.DrawInventory
    Call frmCrafteo.InvCraftCatalyst.DrawInventory
    
    ' Dibujamos el resultado o, si no hay ninguno, el tipo de crafteo
    If frmCrafteo.ResultGrhIndex Then
        Call Draw_GrhIndex(frmCrafteo.ResultGrhIndex, 100, 15)
        Call Engine_Text_Render(JsonLanguage.Item("MENSAJE_553") & frmCrafteo.PorcentajeAcierto & "%", 25, 60, COLOR_WHITE)

        Dim Color(3) As RGBA
        Call RGBAList(Color, 255, 255, 0)
        Call Engine_Text_Render(JsonLanguage.Item("MENSAJE_554") & PonerPuntos(frmCrafteo.PrecioCrafteo) & JsonLanguage.Item("MENSAJE_555"), 25, 140, color)
    Else
        Call Draw_GrhIndex(frmCrafteo.TipoGrhIndex, 100, 15)
    End If

    ' Dibujamos los items arrastrados (aunque sólo puede estar uno activo a la vez)
    Call frmCrafteo.InvCraftUser.DrawDraggedItem
    Call frmCrafteo.InvCraftItems.DrawDraggedItem
    Call frmCrafteo.InvCraftCatalyst.DrawDraggedItem

    ' Presentamos la escena
    Call Engine_EndScene(InvRect, frmCrafteo.PicInven.hwnd)
    RenderCullingRect = GameplayDrawAreaRect
    Exit Sub

DrawInterfaceBoveda_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.DrawInterfaceCrafting", Erl)
    Resume Next
    
End Sub

Public Sub Grh_Render_Advance(ByRef grh As grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByVal Height As Integer, ByVal Width As Integer, ByRef rgb_list() As RGBA, Optional ByVal h_center As Boolean, Optional ByVal v_center As Boolean, Optional ByVal alpha_blend As Boolean = False)
    
    On Error GoTo Grh_Render_Advance_Err

    'Similar to Grh_Render, but let´s you resize the Grh

    Dim tile_width  As Integer

    Dim tile_height As Integer

    Dim grh_index   As Long
    
    If grh.GrhIndex = 0 Or grh.GrhIndex > MaxGrh Then Exit Sub
    
    Dim CurrentFrame As Integer
    CurrentFrame = 1

    If grh.started > 0 Then
        Dim ElapsedFrames As Long
        ElapsedFrames = Fix((FrameTime - grh.started) / grh.speed)

        If grh.Loops = INFINITE_LOOPS Or ElapsedFrames < GrhData(grh.GrhIndex).NumFrames * (grh.Loops + 1) Then
            CurrentFrame = ElapsedFrames Mod GrhData(grh.GrhIndex).NumFrames + 1

        Else
            grh.started = 0
        End If

    End If

    'Figure out what frame to draw (always 1 if not animated)
    grh_index = GrhData(grh.GrhIndex).Frames(CurrentFrame)
    
    'Center Grh over X, Y pos
    If GrhData(grh_index).TileWidth <> 1 Then
        screen_x = screen_x - Int(GrhData(grh_index).TileWidth * (32 \ 2)) + 32 \ 2

    End If
    
    If GrhData(grh_index).TileHeight <> 1 Then
        screen_y = screen_y - Int(GrhData(grh_index).TileHeight * 32) + 32

    End If
    If Not OverlapRect(RenderCullingRect, screen_x, screen_y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight) Then Exit Sub
    'Draw it to device
    'Device_Box_Textured_Render_Advance grh_index, screen_x, screen_y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, rgb_list, GrhData(grh_index).sX, GrhData(grh_index).sY, Width, Height, alpha_blend, grh.angle
    Call Batch_Textured_Box_Advance(screen_x, screen_y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, GrhData(grh_index).sX, GrhData(grh_index).sY, GrhData(grh_index).FileNum, Width, Height, rgb_list, alpha_blend, grh.angle)

    
    Exit Sub

Grh_Render_Advance_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Grh_Render_Advance", Erl)
    Resume Next
    
End Sub

Public Sub Grh_Render(ByRef grh As grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef rgb_list() As RGBA, Optional ByVal h_centered As Boolean = True, Optional ByVal v_centered As Boolean = True, Optional ByVal alpha_blend As Boolean = False)
    
    On Error GoTo Grh_Render_Err

    'Added centering

    Dim tile_width  As Integer

    Dim tile_height As Integer

    Dim grh_index   As Long
    
    If grh.GrhIndex = 0 Or grh.GrhIndex > MaxGrh Then Exit Sub
    
    Dim CurrentFrame As Integer
    CurrentFrame = 1

    If grh.started > 0 Then
        Dim ElapsedFrames As Long
        ElapsedFrames = Fix((FrameTime - grh.started) / grh.speed)

        If grh.Loops = INFINITE_LOOPS Or ElapsedFrames < GrhData(grh.GrhIndex).NumFrames * (grh.Loops + 1) Then
            CurrentFrame = ElapsedFrames Mod GrhData(grh.GrhIndex).NumFrames + 1

        Else
            grh.started = 0
        End If

    End If

    'Figure out what frame to draw (always 1 if not animated)
    grh_index = GrhData(grh.GrhIndex).Frames(CurrentFrame)
    
    'Center Grh over X, Y pos
    If GrhData(grh_index).TileWidth <> 1 Then
        screen_x = screen_x - Int(GrhData(grh_index).TileWidth * (32 \ 2)) + 32 \ 2

    End If
    
    If GrhData(grh_index).TileHeight <> 1 Then
        screen_y = screen_y - Int(GrhData(grh_index).TileHeight * 32) + 32

    End If
    If Not OverlapRect(RenderCullingRect, screen_x, screen_y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight) Then Exit Sub
    'Draw it to device
    'Device_Box_Textured_Render grh_index, screen_x, screen_y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, rgb_list(), GrhData(grh_index).sX, GrhData(grh_index).sY, alpha_blend, grh.angle
    Call Batch_Textured_Box(screen_x, screen_y, GrhData(grh_index).pixelWidth, GrhData(grh_index).pixelHeight, GrhData(grh_index).sX, GrhData(grh_index).sY, GrhData(grh_index).FileNum, rgb_list, alpha_blend, grh.angle)

    
    Exit Sub

Grh_Render_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Grh_Render", Erl)
    Resume Next
    
End Sub

Private Function Grh_Check(ByVal grh_index As Long) As Boolean
    
    On Error GoTo Grh_Check_Err

    'check grh_index
    If grh_index > 0 And grh_index <= MaxGrh Then
        Grh_Check = GrhData(grh_index).NumFrames

    End If

    
    Exit Function

Grh_Check_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Grh_Check", Erl)
    Resume Next
    
End Function

Function Engine_PixelPosX(ByVal x As Integer) As Integer

    'Converts a tile position to a screen position
    
    On Error GoTo Engine_PixelPosX_Err
    
    Engine_PixelPosX = (x - 1) * 32

    
    Exit Function

Engine_PixelPosX_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_PixelPosX", Erl)
    Resume Next
    
End Function

Function Engine_PixelPosY(ByVal y As Integer) As Integer

    'Converts a tile position to a screen position

    On Error GoTo Engine_PixelPosY_Err
    
    Engine_PixelPosY = (y - 1) * 32

    
    Exit Function

Engine_PixelPosY_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_PixelPosY", Erl)
    Resume Next
    
End Function

Function Engine_ElapsedTime() As Long
    
    On Error GoTo Engine_ElapsedTime_Err

    'Gets the time that past since the last call

    Dim Start_Time As Long

    Start_Time = FrameTime
    Engine_ElapsedTime = Start_Time - EndTime

    If Engine_ElapsedTime > 1000 Then Engine_ElapsedTime = 1000
    EndTime = Start_Time

    
    Exit Function

Engine_ElapsedTime_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_ElapsedTime", Erl)
    Resume Next
    
End Function

Private Sub Renderizar_Aura(ByVal aura_index As String, ByVal x As Integer, ByVal y As Integer, ByVal map_x As Byte, ByVal map_y As Byte, Optional ByVal userIndex As Long = 0)
    
    On Error GoTo Renderizar_Aura_Err
    

    Dim rgb_list(3)      As RGBA

    Dim i                As Byte

    Dim Index            As Long

    Dim Color            As Long

    Dim aura_grh         As grh

    Dim giro             As Single

    Dim lado             As Byte

    Index = Val(ReadField(1, aura_index, Asc(":")))
    Color = Val(ReadField(2, aura_index, Asc(":")))
    giro = Val(ReadField(3, aura_index, Asc(":")))
    lado = Val(ReadField(4, aura_index, Asc(":")))

    'frmdebug.add_text_tracebox charlist(userindex).AuraAngle
    If giro > 0 And userIndex > 0 Then
        'If lado = 0 Then
        charlist(userIndex).AuraAngle = charlist(userIndex).AuraAngle + (timerTicksPerFrame * giro)
        'Else
        'charlist(userindex).AuraAngle = charlist(userindex).AuraAngle - (timerTicksPerFrame * giro)
        ' End If
    
        If charlist(userIndex).AuraAngle >= 360 Then charlist(userIndex).AuraAngle = 0

    End If

    Call Long_2_RGBAList(rgb_list, Color)

    'Convertimos el Aura en un GRH
    Call InitGrh(aura_grh, Index)

    'Y por ultimo renderizamos esta capa con Draw_Grh
    If giro > 0 And userIndex > 0 Then
        Call Draw_Grh(aura_grh, x, y + 30, 1, 0, rgb_list(), True, map_x, map_y, charlist(userIndex).AuraAngle)
    Else
        Call Draw_Grh(aura_grh, x, y + 30, 1, 0, rgb_list(), True, map_x, map_y, 0)

    End If
    
    
    Exit Sub

Renderizar_Aura_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Renderizar_Aura", Erl)
    Resume Next
    
End Sub

Public Sub RenderConnect(ByVal TileX As Integer, ByVal TileY As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
    
    On Error GoTo RenderConnect_Err
    
#If DXUI = 0 Then
    Call Engine_BeginScene
#End If

     Select Case UserMap
        Case 1 ' ulla 45-43
            TileX = 45
            TileY = 43
        Case 34 ' nix 22-75
            TileX = 22
            TileY = 75
        Case 59 ' bander 49-43
            TileX = 49
            TileY = 43
        Case 151 ' Arghal 38-41
            TileX = 41
            TileY = 50
        Case 62 ' Lindos 63-40
            TileX = 64
            TileY = 44
        Case 195 ' Arkhein 64-32
            TileX = 76
            TileY = 26
        Case 112 ' Esperanza 50-45
            TileX = 62
            TileY = 51
        Case 354 ' Polo 78-66
            TileX = 33
            TileY = 38
        Case 559 ' Penthar 33-50
            TileX = 34
            TileY = 50
        Case 188 ' Penthar 48-36
            TileX = 48
            TileY = 36
    End Select
    
    
    Call RenderScreen(TileX, TileY, PixelOffsetX, PixelOffsetY, HalfConnectTileWidth, HalfConnectTileHeight)
        
    Dim DefaultColor(3) As Long

    Dim Color           As Long
    Dim ColorGM(3) As RGBA
    ColorGM(0) = RGBA_From_Comp(248, 107, 3)
    ColorGM(1) = ColorGM(0)
    ColorGM(2) = ColorGM(0)
    ColorGM(3) = ColorGM(0)
    intro = 1

    If intro = 1 Then
        Draw_Grh BodyData(123).Walk(3), 490, 333, 1, 0, COLOR_WHITE
        Draw_Grh HeadData(118).Head(3), 490, 328, 1, 0, COLOR_WHITE
            
        Draw_Grh CascoAnimData(13).Head(3), 490, 326, 1, 0, COLOR_WHITE
        Draw_Grh WeaponAnimData(6).WeaponWalk(3), 490, 333, 1, 0, COLOR_WHITE
        Engine_Text_Render "Gulfas Morgolock", 454, 367, ColorGM, 1
        Engine_Text_Render JsonLanguage.Item("MENSAJE_556"), 443, 382, ColorGM, 1

        RenderText "v" & App.Major & "." & App.Minor & " Build: " & App.Revision, 40, 20, COLOR_WHITE, 4, False
    End If

    LastOffsetX = ParticleOffsetX
    LastOffsetY = ParticleOffsetY
    
    TextEfectAsistente = TextEfectAsistente + (15 * timerTicksPerFrame * Sgn(-1))

    If TextEfectAsistente <= 1 Then
        TextEfectAsistente = 0
    End If

    Engine_Text_Render TextAsistente, 510 - Engine_Text_Width(TextAsistente, True, 1) / 2, 287 - Engine_Text_Height(TextAsistente, True) + TextEfectAsistente, textcolorAsistente, 1, True, , 200

    'Logo viejo
    Dim TempGrh As grh, cc(3) As RGBA
    
    Call InitGrh(TempGrh, 1172)

    Call RGBAList(cc, 255, 255, 255, 255)

    Draw_Grh TempGrh, (D3DWindow.BackBufferWidth - GrhData(TempGrh.GrhIndex).pixelWidth) \ 2 + 6, 10, 0, 1, cc(), False

    'Logo nuevo
    'Marco
    Call InitGrh(TempGrh, 1169)

    Draw_Grh TempGrh, 0, 0, 0, 0, COLOR_WHITE, False

    Call InitGrh(TempGrh, 16577)

    Draw_Grh TempGrh, 810, 655, 0, 1, cc(), False

    If FadeInAlpha > 0 Then
        Call Engine_Draw_Box(0, 0, D3DWindow.BackBufferWidth, D3DWindow.BackBufferHeight, RGBA_From_Comp(0, 0, 0, FadeInAlpha))
        FadeInAlpha = FadeInAlpha - 10 * timerTicksPerFrame
    End If

#If DXUI = 0 Then
    Call Engine_EndScene(Render_Connect_Rect, frmConnect.render.hwnd)
#End If
    
    FrameTime = GetTickCount()
    'FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
    
    Exit Sub

RenderConnect_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.RenderConnect", Erl)
    Resume Next
    
End Sub




Public Sub RenderCrearPJ(ByVal TileX As Integer, ByVal TileY As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
    
    On Error GoTo RenderCrearPJ_Err
    

    Call Engine_BeginScene

    Call RenderScreen(TileX, TileY, PixelOffsetX, PixelOffsetY, HalfConnectTileWidth, HalfConnectTileHeight)

    RenderUICrearPJ

    Dim TempGrh As grh
    Call InitGrh(TempGrh, 1171)

    Draw_Grh TempGrh, 494, 190, 1, 1, COLOR_WHITE, False
    'Logo viejo
    
    'Marco
    Call InitGrh(TempGrh, 1169)

    Draw_Grh TempGrh, 0, 0, 0, 0, COLOR_WHITE, False

    Call Engine_EndScene(Render_Connect_Rect, frmConnect.render.hwnd)

    FrameTime = GetTickCount()
    FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed

    'RenderAccountCharacters

    
    Exit Sub

RenderCrearPJ_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.RenderCrearPJ", Erl)
    Resume Next
    
End Sub

Public Sub rendercuenta(ByVal TileX As Integer, ByVal TileY As Integer, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
    
    On Error GoTo rendercuenta_Err
    

    Call Engine_BeginScene

    FrameTime = GetTickCount()
    FramesPerSecCounter = FramesPerSecCounter + 1
    timerElapsedTime = GetElapsedTime()
    timerTicksPerFrame = timerElapsedTime * engineBaseSpeed

    RenderAccountCharacters
    
    Call Engine_EndScene(Render_Connect_Rect, frmConnect.render.hwnd)
    
    Exit Sub

rendercuenta_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.rendercuenta", Erl)
    Resume Next
    
End Sub

Public Sub RenderUICrearPJ()
    
    On Error GoTo RenderUICrearPJ_Err
    

    Dim TempGrh As grh
    
    Dim ColorGray(3) As RGBA
    Call RGBAList(ColorGray, 200, 200, 200)

    Call InitGrh(TempGrh, 727)
    
    Draw_Grh TempGrh, 475, 545, 1, 1, COLOR_WHITE, False

    Engine_Text_Render JsonLanguage.Item("MENSAJE_557"), 280, 125, ColorGray, 5, False

    Dim OffsetX As Integer
    Dim OffsetY As Integer

    RenderText JsonLanguage.Item("MENSAJE_558"), 460, 205, COLOR_WHITE, 6, False

    OffsetX = 240
    OffsetY = 15
    RenderText JsonLanguage.Item("MENSAJE_559"), 345 + OffsetX, 240 + OffsetY, COLOR_WHITE, 6, False

    Engine_Draw_Box 317 + OffsetX, 260 + OffsetY, 95, 21, RGBA_From_Comp(1, 1, 1, 100)
    Engine_Text_Render "<", 300 + OffsetX, 260 + OffsetY, COLOR_WHITE, 1, False
        
    Engine_Text_Render ">", 418 + OffsetX, 261 + OffsetY, COLOR_WHITE, 1, False

    Engine_Text_Render frmCrearPersonaje.lstProfesion.List(frmCrearPersonaje.lstProfesion.ListIndex), 365 + OffsetX - Engine_Text_Width(frmCrearPersonaje.lstProfesion.List(frmCrearPersonaje.lstProfesion.ListIndex), True, 1) / 2, 262 + OffsetY, ColorGray, 1, True

    RenderText JsonLanguage.Item("MENSAJE_560"), 347 + OffsetX, 290 + OffsetY, COLOR_WHITE, 6, False
    Engine_Draw_Box 317 + OffsetX, 305 + OffsetY, 95, 21, RGBA_From_Comp(1, 1, 1, 100)

    'Engine_Text_Render "Humano", 470 - Engine_Text_Height("Humano", False), 304, DefaultColor, 1, False
    Engine_Text_Render frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex), 360 + OffsetX - Engine_Text_Width(frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex), True, 1) / 2, 308 + OffsetY, ColorGray, 1, True
    
    
    
    Engine_Text_Render "<", 300 + OffsetX, 305 + OffsetY, ColorGray, 1, False
    Engine_Text_Render ">", 418 + OffsetX, 305 + OffsetY, ColorGray, 1, False

    OffsetX = 5
    OffsetY = 5

    RenderText JsonLanguage.Item("MENSAJE_561"), 340 + OffsetX, 255, COLOR_WHITE, 6, False
    
    
    Engine_Draw_Box 317 + OffsetX, 275, 95, 21, RGBA_From_Comp(1, 1, 1, 100)

    Engine_Text_Render frmCrearPersonaje.lstGenero.List(frmCrearPersonaje.lstGenero.ListIndex), 360 + OffsetX - Engine_Text_Width(frmCrearPersonaje.lstGenero.List(frmCrearPersonaje.lstGenero.ListIndex), True, 1) / 2, 277, ColorGray, 1, True
    
    Engine_Text_Render "<", 300 + OffsetX, 275, ColorGray, 1, False
    
    Engine_Text_Render ">", 418 + OffsetX, 275, ColorGray, 1, False
    
    'NACIMIENTO
    

    OffsetY = 30
    RenderText JsonLanguage.Item("MENSAJE_562"), 340 + OffsetX, 305, ColorGray, 6, False
    Engine_Draw_Box 317 + OffsetX, 320, 95, 21, RGBA_From_Comp(1, 1, 1, 100)
    
    Engine_Text_Render frmCrearPersonaje.lstHogar.List(frmCrearPersonaje.lstHogar.ListIndex), 360 + OffsetX - Engine_Text_Width(frmCrearPersonaje.lstHogar.List(frmCrearPersonaje.lstHogar.ListIndex), True, 1) / 2, 322, ColorGray, 1, True

    Engine_Text_Render "<", 300 + OffsetX, 320, ColorGray, 1, False
    
    Engine_Text_Render ">", 418 + OffsetX, 320, ColorGray, 1, False

    
    'NACIMIENTO
     Dim Offy As Long
     Offy = -38

     Dim OffX As Long
     OffX = 340
    
    'Atributos
    RenderText JsonLanguage.Item("MENSAJE_563"), 235 + OffX, 385 + Offy, COLOR_WHITE, 6, True
    
    Dim atributeValue As Long
    
    atributeValue = Val(frmCrearPersonaje.lbFuerza.Caption) + Val(frmCrearPersonaje.modfuerza.Caption)
    RenderText JsonLanguage.Item("MENSAJE_564"), 185 + OffX, 410 + Offy, COLOR_WHITE, 1, True
    Call renderAttributesColors(atributeValue, 305 + OffX, 413 + Offy) 'Atributo Fuerza
    
    atributeValue = Val(frmCrearPersonaje.lbAgilidad.Caption) + Val(frmCrearPersonaje.modAgilidad.Caption)
    Engine_Text_Render JsonLanguage.Item("MENSAJE_565"), 185 + OffX, 440 + Offy, COLOR_WHITE, 1, True
    Call renderAttributesColors(atributeValue, 305 + OffX, 443 + Offy) ' Atributo Agilidad
    
    
    atributeValue = Val(frmCrearPersonaje.lbInteligencia.Caption) + Val(frmCrearPersonaje.modInteligencia.Caption)
    Engine_Text_Render JsonLanguage.Item("MENSAJE_566"), 185 + OffX, 470 + Offy, COLOR_WHITE, 1, True
    Call renderAttributesColors(atributeValue, 305 + OffX, 473 + Offy) ' Atributo Inteligencia
    
    
    atributeValue = Val(frmCrearPersonaje.lbConstitucion.Caption) + Val(frmCrearPersonaje.modConstitucion.Caption)
    Engine_Text_Render JsonLanguage.Item("MENSAJE_567"), 185 + OffX, 500 + Offy, COLOR_WHITE, , True
    Call renderAttributesColors(atributeValue, 305 + OffX, 503 + Offy) ' Atributo Constitución
    
    
    atributeValue = Val(frmCrearPersonaje.lbCarisma.Caption) + Val(frmCrearPersonaje.modCarisma.Caption)
    Engine_Text_Render JsonLanguage.Item("MENSAJE_568"), 185 + OffX, 530 + Offy, COLOR_WHITE, , True
    Call renderAttributesColors(atributeValue, 305 + OffX, 533 + Offy) ' Atributo Carisma
      
    
    

     OffX = -340
     Offy = -100




    Dim OffAspectoX As Integer
    Dim OffAspectoY As Integer
    
    
    OffAspectoX = -5
    OffAspectoY = -40
     
     
         
     Engine_Draw_Box 280 + OffAspectoX, 407 + OffAspectoY, 185, 148, RGBA_From_Comp(0, 0, 0, 80)
     
     
     RenderText "Aspecto", 345 + OffAspectoX, 385 + OffAspectoY, COLOR_WHITE, 6, False
     
     
     
    ' CPHeading = 3

    If CPHead <> 0 And CPArma <> 0 Then
    
         
    RenderText "Cabeza", 350 + OffAspectoX, 410 + OffAspectoY, COLOR_WHITE, 1, False
    Engine_Text_Render "<", 335 + OffAspectoX, 412 + OffAspectoY, COLOR_WHITE, 1, True
    Engine_Text_Render ">", 403 + OffAspectoX, 412 + OffAspectoY, COLOR_WHITE, 1, True
        
    Engine_Text_Render ">", 423 + OffAspectoX, 448 + OffAspectoY, COLOR_WHITE, 3, True
    Engine_Text_Render "<", 293 + OffAspectoX, 448 + OffAspectoY, COLOR_WHITE, 3, True
        
    
    
    Dim Raza As Byte
    If frmCrearPersonaje.lstRaza.ListIndex < 0 Then
    frmCrearPersonaje.lstRaza.ListIndex = 0
    End If
    Raza = frmCrearPersonaje.lstRaza.ListIndex
    Dim enanooff As Byte
    If Raza = 0 Or Raza = 1 Or Raza = 2 Or Raza = 5 Then
    enanooff = 0
    
    Else
    enanooff = 10
    End If

            
            If enanooff > 0 Then
                Draw_Grh BodyData(CPBodyE).Walk(CPHeading), 689 + OffX, 346 - Offy, 1, 0, COLOR_WHITE
            Else
                Draw_Grh BodyData(CPBody).Walk(CPHeading), 689 + OffX, 346 - Offy, 1, 0, COLOR_WHITE
            End If
            
            Draw_Grh HeadData(CPHead).Head(CPHeading), 689 + OffX, 346 - Offy + BodyData(CPBody).HeadOffset.y + enanooff, 1, 0, COLOR_WHITE
            
            'If CPEquipado Then
            'Draw_Grh CascoAnimData(CPGorro).Head(CPHeading), 700 + OffX, 366 - Offy + BodyData(CPBody).HeadOffset.y + enanooff, 1, 0, DefaultColor()
            'Draw_Grh WeaponAnimData(CPArma).WeaponWalk(CPHeading), 685 + 15 + OffX, 365 - Offy + enanooff, 1, 0, DefaultColor()
            'Call Renderizar_Aura(CPAura, 686 + 15 + offx, 360 - offy, 0, 0)
            'End If
            
            Dim Color(3) As RGBA
            
            Color(0) = RGBA_From_Comp(0, 128, 190)
            Color(1) = Color(0)
            Color(2) = Color(0)
            Color(3) = Color(0)
            Engine_Text_Render CPName, 365 - Engine_Text_Width(CPName, True) / 2, 478, Color, 1, True
    Else
        Engine_Text_Render "X", 360, 428, COLOR_WHITE, 3, True
    End If
    
    
    'DefaultColor(0) = D3DColorXRGB(255, 255, 255)
    'DefaultColor(1) = DefaultColor(0)
    'DefaultColor(2) = DefaultColor(0)
    'DefaultColor(3) = DefaultColor(0)

    'Boton Atras
    'Engine_Draw_Box 147, 628, 100, 40, D3DColorARGB(80, 0, 0, 0)
    'Engine_Text_Render "< Volver", 170, 640, DefaultColor, 1, True
    
    'Boton Crear
    'If StopCreandoCuenta Then
    '    Engine_Draw_Box 730, 630, 100, 40, D3DColorARGB(120, 100, 180, 100)
    '    Engine_Text_Render "Creando...", 750, 640, DefaultColor, 1, True
    'Else
    '    Engine_Draw_Box 730, 630, 100, 40, D3DColorARGB(80, 0, 0, 0)
    '    Engine_Text_Render "Crear PJ >", 750, 640, DefaultColor, 1, True
    'End If

       
    'Engine_Text_Render "DADO", 670, 390, DefaultColor()
    'Draw_GrhIndex 1123, 655, 345

    
    Exit Sub

RenderUICrearPJ_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.RenderUICrearPJ", Erl)
    Resume Next
    
End Sub
Private Function renderAttributesColors(ByVal Value As Integer, ByVal x As Integer, ByVal y As Integer)
        If Value > 18 Then
        RenderText str(Value), x, y, COLOR_GREEN, 1, True
    ElseIf Value < 18 Then
        RenderText str(Value), x, y, COLOR_RED, 1, True
    Else
        RenderText str(Value), x, y, COLOR_WHITE, 1, True
    End If
End Function


Public Sub RenderAccountCharacters()
On Error GoTo RenderAccountCharacters_Err
    Dim i               As Long: Dim sumax As Long
    Dim x               As Integer: Dim y As Integer
    Dim notY            As Integer
    Dim Color           As RGBA
    Dim Texto           As String
    Dim temp_array(3) As RGBA
    Dim TempColor(3) As RGBA
    Dim grh As grh
    
    Texto = CuentaEmail
    sumax = 84
    
    'Dibujo la escena debajo del mapa
    Call RenderScreen(RenderCuenta_PosX, RenderCuenta_PosY, 0, 0, HalfConnectTileWidth, HalfConnectTileHeight)
    
    If LastPJSeleccionado <> PJSeleccionado Then
        If AlphaRenderCuenta < MAX_ALPHA_RENDER_CUENTA Then
            AlphaRenderCuenta = min(AlphaRenderCuenta + timerTicksPerFrame * 10, MAX_ALPHA_RENDER_CUENTA)
        Else
            LastPJSeleccionado = PJSeleccionado
            If PJSeleccionado <> 0 Then
                Call SwitchMap(Pjs(PJSeleccionado).Mapa)
                RenderCuenta_PosX = Pjs(PJSeleccionado).PosX
                RenderCuenta_PosY = Pjs(PJSeleccionado).PosY
            End If
        End If
    ElseIf PJSeleccionado <> 0 And AlphaRenderCuenta > 0 Then
        If Pjs(PJSeleccionado).Mapa <> 0 Then
            AlphaRenderCuenta = max(AlphaRenderCuenta - timerTicksPerFrame * 10, 0)
        End If
    End If

    Call RGBAList(TempColor, 255, 255, 255, 100 + AlphaRenderCuenta)
    
    Call InitGrh(grh, 4531)
    Call Draw_Grh(grh, 0, 0, 0, 0, TempColor, False, 0, 0, 0)
    
    Call Draw_GrhIndex(GrhCharactersScreenUI, 0, 0)
      
    For i = 1 To MAX_PERSONAJES_EN_CUENTA
            
        If (i > 5) Then
            x = ((i * 132) - (5 * 132))
            y = 440
        Else
            x = (i * 132)
            y = 283

        End If

        x = x + sumax

        temp_array(0) = Pjs(i).LetraColor
        temp_array(1) = Pjs(i).LetraColor
        temp_array(2) = Pjs(i).LetraColor
        temp_array(3) = Pjs(i).LetraColor
        
        Dim Body As Integer
        Dim enBarca As Boolean
        Body = Pjs(i).Body
        
        If (Body <> 0) Then
            If PJSeleccionado = i Then
                Call Particle_Group_Render(Select_part, x + 32, y + 5)
            End If

            If (Body <> 0) Then
                Draw_Grh BodyData(Body).Walk(3), x + 15, y + 10, 1, 1, COLOR_WHITE
            End If
            enBarca = Body = 84 Or Body = 85 Or Body = 86 Or Body = 87 Or Body = 1263 Or Body = 1264 Or Body = 1265 Or Body = 1266 Or Body = 1267 Or Body = 1268 Or Body = 1269 Or Body = 1270 Or Body = 1271 Or Body = 1272 Or Body = 1273 Or Body = 1274
            
            If (Pjs(i).Head <> 0) And Not enBarca Then
                Draw_Grh HeadData(Pjs(i).Head).Head(3), x + 15, y - notY + BodyData(Pjs(i).Body).HeadOffset.y + 10, 1, 0, COLOR_WHITE
            End If
            
            If (Pjs(i).Casco <> 0) Then
                If Pjs(i).Casco <= UBound(CascoAnimData) And Pjs(i).Casco >= LBound(CascoAnimData) Then
                    Draw_Grh CascoAnimData(Pjs(i).Casco).Head(3), x + 15, y - notY + BodyData(Pjs(i).Body).HeadOffset.y + 10, 1, 0, COLOR_WHITE
                End If
            End If
            
            If (Pjs(i).Escudo <> 0) Then
                If Pjs(i).Escudo <= UBound(ShieldAnimData) And Pjs(i).Escudo >= LBound(ShieldAnimData) Then
                    Draw_Grh ShieldAnimData(Pjs(i).Escudo).ShieldWalk(3), x + 14, y - notY + 10, 1, 0, COLOR_WHITE
                End If
            End If
                        
            If (Pjs(i).Arma <> 0) Then
                If Pjs(i).Arma <= UBound(WeaponAnimData) And Pjs(i).Arma >= LBound(WeaponAnimData) Then
                    Draw_Grh WeaponAnimData(Pjs(i).Arma).WeaponWalk(3), x + 14, y - notY + 10, 1, 0, COLOR_WHITE
                End If
            End If

            Engine_Text_Render Pjs(i).nombre, x + 30 - Engine_Text_Width(Pjs(i).nombre, True) / 2, y + 56 - Engine_Text_Height(Pjs(i).nombre, True), temp_array(), 1, True
            
            If PJSeleccionado = i Then
                Dim Offy As Byte
                Offy = 0
                Engine_Text_Render Pjs(i).nombre, 511 - Engine_Text_Width(Pjs(i).nombre, True) / 2, 565 - Engine_Text_Height(Pjs(i).nombre, True), temp_array(), 1, True
                If Pjs(i).ClanName <> "<>" Then
                    Engine_Text_Render Pjs(i).ClanName, 511 - Engine_Text_Width(Pjs(i).ClanName, True) / 2, 565 + 15 - Engine_Text_Height(Pjs(i).ClanName, True), temp_array(), 1, True
                    Offy = 15
                Else
                    Offy = 0
                End If
                Engine_Text_Render JsonLanguage.Item("MENSAJE_569") & ListaClases(Pjs(i).Clase), 511 - Engine_Text_Width(JsonLanguage.Item("MENSAJE_569") & ListaClases(Pjs(i).Clase), True) / 2, Offy + 570 - Engine_Text_Height(JsonLanguage.Item("MENSAJE_569") & ListaClases(Pjs(i).Clase), True), COLOR_WHITE, 1, True
                Engine_Text_Render JsonLanguage.Item("MENSAJE_570") & Pjs(i).nivel, 511 - Engine_Text_Width(JsonLanguage.Item("MENSAJE_570") & Pjs(i).nivel, True) / 2, Offy + 585 - Engine_Text_Height(JsonLanguage.Item("MENSAJE_570") & Pjs(i).nivel, True), COLOR_WHITE, 1, True
                Engine_Text_Render CStr(Pjs(i).NameMapa), 511 - Engine_Text_Width(CStr(Pjs(i).NameMapa), True) / 2, Offy + 615 - Engine_Text_Height(CStr(Pjs(i).NameMapa), True), COLOR_WHITE, 1, True
            End If
        End If
    Next i

    Exit Sub

RenderAccountCharacters_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.RenderAccountCharacters", Erl)
    Resume Next
    
End Sub

Sub EfectoEnPantalla(ByVal Color As Long, ByVal Time As Long)
    
    On Error GoTo EfectoEnPantalla_Err
    
    frmMain.Efecto.Interval = Time
    frmMain.Efecto.enabled = True
    EfectoEnproceso = True
    Call SetGlobalLight(Color)

    
    Exit Sub

EfectoEnPantalla_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.EfectoEnPantalla", Erl)
    Resume Next
    
End Sub

Public Sub SetBarFx(ByVal CharIndex As Integer, ByVal BarTime As Integer)

    'Sets an FX to the character.
    
    On Error GoTo SetBarFx_Err
    
    With charlist(CharIndex)
        .BarTime = BarTime

    End With

    Exit Sub

SetBarFx_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.SetBarFx", Erl)
    Resume Next
    
End Sub

Public Function Engine_Get_2_Points_Angle(ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Double

    On Error GoTo Engine_Get_2_Points_Angle_Err
    

    Engine_Get_2_Points_Angle = Engine_Get_X_Y_Angle((x2 - x1), (y2 - y1))
   
    
    Exit Function

Engine_Get_2_Points_Angle_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_Get_2_Points_Angle", Erl)
    Resume Next
    
End Function

Public Function Engine_Get_X_Y_Angle(ByVal x As Double, ByVal y As Double) As Double

    On Error GoTo Engine_Get_X_Y_Angle_Err
    

    Dim dblres As Double
 
    dblres = 0
   
    If (y <> 0) Then
        dblres = Engine_Convert_Radians_To_Degrees(Atn(x / y))

        If (x <= 0 And y < 0) Then
            dblres = dblres + 180
        ElseIf (x > 0 And y < 0) Then
            dblres = dblres + 180
        ElseIf (x < 0 And y > 0) Then
            dblres = dblres + 360

        End If

    Else

        If (x > 0) Then
            dblres = 90
        ElseIf (x < 0) Then
            dblres = 270

        End If

    End If
   
    Engine_Get_X_Y_Angle = dblres
   
    
    Exit Function

Engine_Get_X_Y_Angle_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_Get_X_Y_Angle", Erl)
    Resume Next
    
End Function

Public Function Engine_Convert_Radians_To_Degrees(ByVal s_radians As Double) As Integer

    'Converts a radian to degrees

    On Error GoTo Engine_Convert_Radians_To_Degrees_Err
    

    Engine_Convert_Radians_To_Degrees = (s_radians * 180) / 3.14159265358979
 
    
    Exit Function

Engine_Convert_Radians_To_Degrees_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_Convert_Radians_To_Degrees", Erl)
    Resume Next
    
End Function

' programado por maTih.-
 
Public Sub InitializeInventory()
    '
    ' @ Inicializa el array de efectos.
    
    On Error GoTo Initialize_Err
    
     
    ReDim Effect(1 To 255) As Effect_Type
    
    ' Inicializo inventarios
    Set frmMain.Inventario = New clsGrapchicalInventory
    Set frmComerciar.InvComUsu = New clsGrapchicalInventory
    Set frmComerciar.InvComNpc = New clsGrapchicalInventory
    Set frmBancoObj.InvBankUsu = New clsGrapchicalInventory
    Set frmBancoObj.InvBoveda = New clsGrapchicalInventory
    Set frmComerciarUsu.InvUser = New clsGrapchicalInventory
    Set frmComerciarUsu.InvUserSell = New clsGrapchicalInventory
    Set frmComerciarUsu.InvOtherSell = New clsGrapchicalInventory
    
    Set frmBancoCuenta.InvBankUsuCuenta = New clsGrapchicalInventory
    Set frmBancoCuenta.InvBovedaCuenta = New clsGrapchicalInventory
    
    Set FrmKeyInv.InvKeys = New clsGrapchicalInventory
    
    Call frmMain.Inventario.Initialize(frmMain.picInv, MAX_INVENTORY_SLOTS, , , 0, 0, 3, 3, True, 9)
    Call frmComerciar.InvComUsu.Initialize(frmComerciar.interface, MAX_INVENTORY_SLOTS, 210, 0, 252, 0, 3, 3, True)
    Call frmComerciar.InvComNpc.Initialize(frmComerciar.interface, MAX_INVENTORY_SLOTS, 210, , 1, 0, 3, 3)
    Call frmComerciarUsu.InvUser.Initialize(frmComerciarUsu.picInv, MAX_INVENTORY_SLOTS, , , 0, 0, 3, 3, True)
    Call frmComerciarUsu.InvUserSell.Initialize(frmComerciarUsu.picInvUserSell, 6, , , 0, 0, 3, 3, True)
    Call frmComerciarUsu.InvOtherSell.Initialize(frmComerciarUsu.picInvOtherSell, 6, , , 0, 0, 3, 3, True)
   
    Call frmBancoObj.InvBankUsu.Initialize(frmBancoObj.interface, MAX_INVENTORY_SLOTS, 210, 0, 252, 0, 3, 3, True)
    Call frmBancoObj.InvBoveda.Initialize(frmBancoObj.interface, MAX_BANCOINVENTORY_SLOTS, 210, 0, 0, 0, 3, 3)
    
    Call frmBancoCuenta.InvBankUsuCuenta.Initialize(frmBancoCuenta.interface, MAX_INVENTORY_SLOTS, 210, 0, 252, 0, 3, 3, True)
    Call frmBancoCuenta.InvBovedaCuenta.Initialize(frmBancoCuenta.interface, MAX_BANCOINVENTORY_SLOTS, 210, 0, 0, 0, 3, 3)

    Call FrmKeyInv.InvKeys.Initialize(FrmKeyInv.interface, MAX_KEYS, , , 0, 0, 3, 3, True) 'Inventario de llaves
    FrmKeyInv.InvKeys.MostrarCantidades = False
    
    Set frmCrafteo.InvCraftUser = New clsGrapchicalInventory
    Set frmCrafteo.InvCraftItems = New clsGrapchicalInventory
    Set frmCrafteo.InvCraftCatalyst = New clsGrapchicalInventory

    Call frmCrafteo.InvCraftUser.Initialize(frmCrafteo.PicInven, MAX_INVENTORY_SLOTS, 210, , 250, 0, 3, 3, True)
    Call frmCrafteo.InvCraftItems.Initialize(frmCrafteo.PicInven, MAX_SLOTS_CRAFTEO, 175, , 25, 180, 3, 3, True)
    Call frmCrafteo.InvCraftCatalyst.Initialize(frmCrafteo.PicInven, 1, 35, 35, 100, 90, 3, 3, True)

    Exit Sub

Initialize_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Initialize", Erl)
    Resume Next
    
End Sub

Public Sub Terminate_Index(ByVal effect_Index As Integer)
    
    On Error GoTo Terminate_Index_Err
    
 
    '
    ' @ Destruye un indice del array
 
    Dim clear_Index As Effect_Type
 
    'Si es un slot válido
    If (effect_Index <> 0) And (effect_Index <= UBound(Effect())) Then
        Effect(effect_Index) = clear_Index

    End If
 
    
    Exit Sub

Terminate_Index_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Terminate_Index", Erl)
    Resume Next
    
End Sub
 
Public Function Effect_Begin(ByVal Fx_Index As Integer, ByVal Bind_Speed As Single, ByVal x As Single, ByVal y As Single, Optional ByVal explosion_FX_Index As Integer = -1, Optional ByVal explosion_FX_Loops As Integer = -1, Optional ByVal receptor As Integer = 1, Optional ByVal Emisor As Integer = 1, Optional ByVal wav As Integer = 1, Optional ByVal Fx As Integer = -1) As Integer
    
    On Error GoTo Effect_Begin_Err
    
 
    '
    ' @ Inicia un nuevo efecto y devuelve el index.
 
    Effect_Begin = GetFreeIndex()
 
    ' frmdebug.add_text_tracebox "fx =" & fX
    'Si hay efecto
    If (Effect_Begin <> 0) Then
    
        With Effect(Effect_Begin)
            .Now_X = CInt(x) - 16
            .Now_Y = CInt(y) - 20
         
            .Fx_Index = Fx_Index
         
            .ViajeSpeed = Bind_Speed
            .ViajeChar = Emisor
            .DestinoChar = receptor
         
            .wav = wav
         
            'Explosión?
            If (explosion_FX_Index <> 0) And (Fx = 0) Then
                .End_Effect = explosion_FX_Index
                .End_Loops = explosion_FX_Loops
                .FxEnd_Effect = 0
                
            End If
         
            If (Fx = 1) Then
                .End_Effect = 0
                .FxEnd_Effect = explosion_FX_Index
                .End_Loops = explosion_FX_Loops
            
            End If
         
            .Slot_Used = True
         
        End With
    
    End If
 
    
    Exit Function

Effect_Begin_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Effect_Begin", Erl)
    Resume Next
    
End Function

Public Function Effect_BeginXY(ByVal Fx_Index As Integer, ByVal Bind_Speed As Single, ByVal x As Single, ByVal y As Single, ByVal DestinoX As Byte, ByVal Destinoy As Byte, Optional ByVal explosion_FX_Index As Integer = -1, Optional ByVal explosion_FX_Loops As Integer = -1, Optional ByVal Emisor As Integer = 1, Optional ByVal wav As Integer = 1, Optional ByVal Fx As Integer = 0) As Integer
    '
    ' @ Inicia un nuevo efecto y devuelve el index.
    
    On Error GoTo Effect_BeginXY_Err
    
 
    ' frmdebug.add_text_tracebox "fx =" & fX
    Effect_BeginXY = GetFreeIndex()
 
    'Si hay efecto
    If (Effect_BeginXY <> 0) Then
    
        With Effect(Effect_BeginXY)
            .Now_X = CInt(x) - 16
            .Now_Y = CInt(y) - 20
         
            .Fx_Index = Fx_Index
         
            .ViajeSpeed = Bind_Speed
            .ViajeChar = Emisor
            .DestinoChar = 0
            .DestX = DestinoX
            .DesyY = Destinoy
         
            .wav = wav
         
            'Explosión?
            If (explosion_FX_Index <> 0) And (Fx = 0) Then
                .End_Effect = explosion_FX_Index
                .End_Loops = explosion_FX_Loops
                .FxEnd_Effect = 0

            End If
         
            If (Fx = 1) Then
                .End_Effect = 0
                .FxEnd_Effect = explosion_FX_Index
                .End_Loops = explosion_FX_Loops

            End If
         
            .Slot_Used = True
         
        End With
    
    End If
 
    
    Exit Function

Effect_BeginXY_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Effect_BeginXY", Erl)
    Resume Next
    
End Function
 
Public Sub Effect_Render_All()
    
    On Error GoTo Effect_Render_All_Err
    
 
    '
    ' @ Dibuja todos los efectos
 
    Dim i As Long

    For i = 1 To UBound(Effect())

        With Effect(i)
         
            If .Slot_Used Then
    
                Effect_Render_Slot CInt(i)

            End If
         
        End With

    Next i
 
    
    Exit Sub

Effect_Render_All_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Effect_Render_All", Erl)
    Resume Next
    
End Sub
 
Public Sub Effect_Render_Slot(ByVal effect_Index As Integer)
    
    On Error GoTo Effect_Render_Slot_Err
    
 
    '
    ' @ Renderiza un efecto.

    Dim colornpcs(3) As Long

    colornpcs(0) = D3DColorXRGB(255, 255, 255)
    colornpcs(1) = colornpcs(0)
    colornpcs(2) = colornpcs(0)
    colornpcs(3) = colornpcs(0)
    
    With Effect(effect_Index)
 
        Dim target_Angle As Single
     
        .Now_Moved = GetTickCount()
     
        'Controla el intervalo de vuelo
        If (.Last_Move + 10) < .Now_Moved Then
            .Last_Move = GetTickCount()
        
            'Si tiene char de destino.
            If (.DestinoChar <> 0) Then
     
                'Actualiza la pos de destino.
                '.Viaje_X = charlist(.ViajeChar).NowPosX
                '.Viaje_Y = charlist(.ViajeChar).NowPosY
            
                .Viaje_X = Get_Pixelx_Of_Char(.DestinoChar) - 0
                .Viaje_Y = Get_PixelY_Of_Char(.DestinoChar) - 32

            Else
                .Viaje_X = Get_Pixelx_Of_XY(.DestX) - 0
                .Viaje_Y = Get_PixelY_Of_XY(.DesyY) - 32

            End If
        
        End If
      
        'Actualiza el ángulo.
        target_Angle = Engine_GetAngle(.Now_X, .Now_Y, CInt(.Viaje_X), CInt(.Viaje_Y))
    
        'Actualiza la posición del efecto.
        .Now_X = (.Now_X + Sin(target_Angle * DegreeToRadian) * .ViajeSpeed * timerTicksPerFrame * 9)
        .Now_Y = (.Now_Y - Cos(target_Angle * DegreeToRadian) * .ViajeSpeed * timerTicksPerFrame * 9)
        'Si hay posición dibuja.
        If (.Now_X <> 0) And (.Now_Y <> 0) Then
            ' Call DDrawTransGrhtoSurface(.FX_Grh, .Now_X, .Now_Y, 1, 1)

            Call Particle_Group_Render(spell_particle, .Now_X, .Now_Y)
        
            'Check si terminó.
            ' If (.FX_Grh.Started = 0) Then .Fx_Index = 0: .Slot_Used = False
        
            If Abs(CInt(.Viaje_X) - CInt(.Now_X)) < 5 Then
                .Now_X = .Viaje_X

            End If

            If Abs(CInt(.Viaje_Y) - CInt(.Now_Y)) < 5 Then
        
                .Now_Y = .Viaje_Y

            End If
        
            If (.Now_X = .Viaje_X) And (.Now_Y = .Viaje_Y) Then
       
                'Inicializa la explosión : p
                If (.End_Effect <> 0) And .DestinoChar <> 0 Then
                    If .DestinoChar <> 0 Then
                        Call General_Char_Particle_Create(.End_Effect, .DestinoChar, .End_Loops)
                        If EstaPCarea(.DestinoChar) Then
                            Call ao20audio.PlayWav(.wav, False, ao20audio.ComputeCharFxVolume(charlist(.DestinoChar).Pos), ao20audio.ComputeCharFxPan(charlist(.DestinoChar).Pos))
                        End If
                        .Slot_Used = False
                        Exit Sub

                    End If

                End If
            
                If (.End_Effect <> 0) And .DestinoChar = 0 Then
                    MapData(.DestX, .DesyY).particle_group = 0
                    General_Particle_Create .End_Effect, .DestX, .DesyY, .End_Loops
                    Call ao20audio.PlayWav(.wav)
                    Dim dest_pos As Position
                    dest_pos.x = .DestX
                    dest_pos.y = .DesyY
                    If EstaEnArea(.DestX, .DesyY) Then
                        Call ao20audio.PlayWav(.wav, False, ao20audio.ComputeCharFxVolume(dest_pos), ao20audio.ComputeCharFxPan(dest_pos))
                    End If
                    .Slot_Used = False
                    Exit Sub

                End If
            
                If (.FxEnd_Effect > 0) And .DestinoChar <> 0 Then
                    If EstaPCarea(.DestinoChar) Then
                        Call ao20audio.PlayWav(.wav, False, ao20audio.ComputeCharFxVolume(charlist(.DestinoChar).Pos), ao20audio.ComputeCharFxPan(charlist(.DestinoChar).Pos))
                    End If
                    Call SetCharacterFx(.DestinoChar, .FxEnd_Effect, .End_Loops)
                    .Slot_Used = False
                    Exit Sub

                End If
            
                If (.FxEnd_Effect > 0) And (.DestinoChar = 0) Then
                    Dim p As Position
                    p.x = .DestX
                    p.y = .DesyY
                    If EstaEnArea(.DestX, .DesyY) Then
                        Call ao20audio.PlayWav(.wav, False, ao20audio.ComputeCharFxVolume(p), ao20audio.ComputeCharFxPan(p))
                    End If
                    Call SetMapFx(.DestX, .DesyY, .FxEnd_Effect, 0)
                    .Slot_Used = False
                    Exit Sub

                End If
          
            End If
        
        End If

    End With
 
    
    Exit Sub

Effect_Render_Slot_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Effect_Render_Slot", Erl)
    Resume Next
    
End Sub
 
Public Function Engine_GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Single

    'Gets the angle between two points in a 2d plane
    'More info: [url=http://www.vbgore.com/GameClient.TileEn]http://www.vbgore.com/GameClient.TileEn[/url] ... e_GetAngle" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;

    Dim SideA As Single

    Dim SideC As Single
 
    On Error GoTo ErrOut
 
    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = TargetY Then
 
        'Check for going right (90 degrees)
        If CenterX < TargetX Then
            Engine_GetAngle = 90
 
            'Check for going left (270 degrees)
        Else
            Engine_GetAngle = 270

        End If
 
        'Exit the function
        Exit Function
 
    End If
 
    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = TargetX Then
 
        'Check for going up (360 degrees)
        If CenterY > TargetY Then
            Engine_GetAngle = 360
 
            'Check for going down (180 degrees)
        Else
            Engine_GetAngle = 180

        End If
 
        'Exit the function
        Exit Function
 
    End If
 
    'Calculate Side C
    SideC = Sqr(Abs(TargetX - CenterX) ^ 2 + Abs(TargetY - CenterY) ^ 2)
 
    'Side B = CenterY
 
    'Calculate Side A
    SideA = Sqr(Abs(TargetX - CenterX) ^ 2 + TargetY ^ 2)
 
    'Calculate the angle
    If CenterY = 0 Then
        Engine_GetAngle = 90
    Else
        Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
        Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583
    End If
 
    'If the angle is >180, subtract from 360
    If TargetX < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle
 
    'Exit function
 
    Exit Function
 
    'Check for error
ErrOut:
 
    'Return a 0 saying there was an error
    Engine_GetAngle = 0
 
    Exit Function
 
End Function
 
Public Function GetFreeIndex() As Integer
    
    On Error GoTo GetFreeIndex_Err
    
 
    '
    ' @ Devuelve un índice para un nuevo FX.
 
    Dim i As Long
 
    For i = 1 To UBound(Effect())

        'No está usado.
        If Not Effect(i).Slot_Used Then
            GetFreeIndex = CInt(i)
            Exit Function

        End If

    Next i
 
    GetFreeIndex = NO_INDEX
 
    
    Exit Function

GetFreeIndex_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.GetFreeIndex", Erl)
    Resume Next
    
End Function

Public Sub Draw_Grh_ItemInWater(ByRef grh As grh, ByVal x As Integer, ByVal y As Integer, ByVal center As Byte, ByVal animate As Byte, ByRef rgb_list() As RGBA, Optional ByVal Alpha As Boolean = False, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal angle As Single)
    
    On Error GoTo Draw_Grh_Err

    If grh.GrhIndex = 0 Or grh.GrhIndex > MaxGrh Then Exit Sub
    
    Dim CurrentFrame As Integer
    CurrentFrame = 1

    If animate Then
        If grh.started > 0 Then
            Dim ElapsedFrames As Long
            ElapsedFrames = Fix(0.5 * (FrameTime - grh.started) / grh.speed)

            If grh.Loops = INFINITE_LOOPS Or ElapsedFrames < GrhData(grh.GrhIndex).NumFrames * (grh.Loops + 1) Then
                CurrentFrame = ElapsedFrames Mod GrhData(grh.GrhIndex).NumFrames + 1

            Else
                grh.started = 0
            End If

        End If

    End If
    
    Dim CurrentGrhIndex As Long
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(grh.GrhIndex).Frames(CurrentFrame)

    'Center Grh over X,Y pos
    If center Then
        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            x = x - Int(GrhData(CurrentGrhIndex).TileWidth * TilePixelWidth \ 2) + TilePixelWidth \ 2
        End If

        If GrhData(grh.GrhIndex).TileHeight <> 1 Then
            y = y - Int(GrhData(CurrentGrhIndex).TileHeight * TilePixelHeight) + TilePixelHeight
        End If
    End If

    With GrhData(CurrentGrhIndex)
        If Not OverlapRect(RenderCullingRect, x, y, .pixelWidth, .pixelHeight) Then Exit Sub
        If .FileNum > 0 Then
            Dim Texture As Direct3DTexture8

            Dim TextureWidth As Long, TextureHeight As Long
            Set Texture = SurfaceDB.GetTexture(.FileNum, TextureWidth, TextureHeight)
            
        
            .Tx1 = .sX / TextureWidth
            .Tx2 = (.sX + .pixelWidth) / TextureWidth
            .Ty1 = .sY / TextureHeight
            .Ty2 = (.sY + .pixelHeight) / TextureHeight
            
            If Alpha Then
                DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
                DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
            End If
            
            Call SpriteBatch.SetTexture(Texture)
            Call SpriteBatch.DrawItemInWater(x, y, .pixelWidth, .pixelHeight, rgb_list, .Tx1, .Ty1, .Tx2, .Ty2, angle Mod 360) ' Angle
            
            If Alpha Then
                DirectDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
                DirectDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
            End If
        End If
        
        'Call Batch_Textured_Box_Pre(x, y, .pixelWidth, .pixelHeight, .Tx1, .Ty1, .Tx2, .Ty2, .FileNum, rgb_list, Alpha, Angle)
    
    End With
    
    Exit Sub

Draw_Grh_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Draw_Grh", Erl)
    Resume Next
    
End Sub

Public Sub Draw_Grh_Precalculated(ByRef grh As grh, ByRef rgb_list() As RGBA, ByVal EsAgua As Boolean, ByVal EsLava As Boolean, ByVal MapX As Integer, ByVal MapY As Integer, ByVal MinX As Integer, ByVal MaxX As Integer, ByVal MinY As Integer, ByVal MaxY As Integer)
    
    On Error GoTo Draw_Grh_Precalculated_Err

    If Not OverlapRect(RenderCullingRect, grh.x - MinX * TilePixelWidth, grh.y - TilePixelHeight * MinY + gameplay_render_offset.y, TilePixelWidth, TilePixelHeight) Then Exit Sub

    If grh.GrhIndex = 0 Or grh.GrhIndex > MaxGrh Then Exit Sub
    
    Dim CurrentFrame As Integer
    CurrentFrame = 1

    If grh.started > 0 Then
        Dim ElapsedFrames As Long
        ElapsedFrames = Fix((FrameTime - grh.started) / grh.speed)

        If grh.Loops = INFINITE_LOOPS Or ElapsedFrames < GrhData(grh.GrhIndex).NumFrames * (grh.Loops + 1) Then
            CurrentFrame = ElapsedFrames Mod GrhData(grh.GrhIndex).NumFrames + 1

        Else
            grh.started = 0
        End If

    End If
    
    Dim CurrentGrhIndex As Long
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(grh.GrhIndex).Frames(CurrentFrame)

    Dim Texture As Direct3DTexture8

    Dim TextureWidth As Long, TextureHeight As Long
    Set Texture = SurfaceDB.GetTexture(GrhData(CurrentGrhIndex).FileNum, TextureWidth, TextureHeight)
    
    With GrhData(CurrentGrhIndex)

        Call SpriteBatch.SetTexture(Texture)
        
        If .Tx2 = 0 And .FileNum > 0 Then
            .Tx1 = .sX / TextureWidth
            .Tx2 = (.sX + .pixelWidth) / TextureWidth
            .Ty1 = .sY / TextureHeight
            .Ty2 = (.sY + .pixelHeight) / TextureHeight
        End If
        
        Dim Top As Byte, Right As Byte, Bottom As Byte, Left As Byte
    
        If EsAgua Then
            If MapY > MinY Then Top = (MapData(MapX, MapY - 1).Blocked And FLAG_AGUA) * INV_FLAG_AGUA
            If MapX < MaxX Then Right = (MapData(MapX + 1, MapY).Blocked And FLAG_AGUA) * INV_FLAG_AGUA
            If MapY < MaxY Then Bottom = (MapData(MapX, MapY + 1).Blocked And FLAG_AGUA) * INV_FLAG_AGUA
            If MapX > MinX Then Left = (MapData(MapX - 1, MapY).Blocked And FLAG_AGUA) * INV_FLAG_AGUA
        
            Call SpriteBatch.DrawWater(grh.x, grh.y, TilePixelWidth, TilePixelHeight, rgb_list, .Tx1, .Ty1, .Tx2, .Ty2, MapX, MapY, Top, Right, Bottom, Left)
        
        ElseIf EsLava Then
            If MapY > MinY Then Top = (MapData(MapX, MapY - 1).Blocked And FLAG_LAVA) * INV_FLAG_LAVA
            If MapX < MaxX Then Right = (MapData(MapX + 1, MapY).Blocked And FLAG_LAVA) * INV_FLAG_LAVA
            If MapY < MaxY Then Bottom = (MapData(MapX, MapY + 1).Blocked And FLAG_LAVA) * INV_FLAG_LAVA
            If MapX > MinX Then Left = (MapData(MapX - 1, MapY).Blocked And FLAG_LAVA) * INV_FLAG_LAVA
        
            Call SpriteBatch.DrawLava(grh.x, grh.y, TilePixelWidth, TilePixelHeight, rgb_list, .Tx1, .Ty1, .Tx2, .Ty2, MapX, MapY, Top, Right, Bottom, Left)
        
        Else
             Call SpriteBatch.Draw(grh.x, grh.y, TilePixelWidth, TilePixelHeight, rgb_list, .Tx1, .Ty1, .Tx2, .Ty2)
             
             
        End If
    
    End With

    
    Exit Sub

Draw_Grh_Precalculated_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Draw_Grh_Precalculated", Erl)
    Resume Next
    
End Sub

Public Sub Engine_Draw_Box(ByVal x As Integer, ByVal y As Integer, ByVal Width As Integer, ByVal Height As Integer, Color As RGBA)
    
    On Error GoTo Engine_Draw_Box_Err
    
    If Not OverlapRect(RenderCullingRect, x, y, Width, Height) Then Exit Sub
    Call RGBAList(temp_rgb, Color.r, Color.G, Color.b, Color.a)
    
    Call SpriteBatch.SetTexture(Nothing)
    Call SpriteBatch.SetAlpha(False)
    Call SpriteBatch.Draw(x, y, Width, Height, temp_rgb())

    
    Exit Sub

Engine_Draw_Box_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_Draw_Box", Erl)
    Resume Next
    
End Sub

Public Sub Engine_Draw_Load(ByVal x As Integer, ByVal y As Integer, ByVal Width As Integer, ByVal Height As Integer, Color As RGBA, angle As Single)
    
    On Error GoTo Engine_Draw_Load_Err
    

    Call RGBAList(temp_rgb, Color.r, Color.G, Color.b, Color.a)


    If angle >= 360 Then angle = 0

    Call SpriteBatch.SetTexture(Nothing)
    Call SpriteBatch.SetAlpha(False)
    Call SpriteBatch.DrawLoad(x, y, Width, Height, temp_rgb(), angle)

    
    Exit Sub

Engine_Draw_Load_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_Draw_Load", Erl)
    Resume Next
    
End Sub

Public Sub Engine_Draw_Box_Border(ByVal x As Integer, ByVal y As Integer, ByVal Width As Integer, ByVal Height As Integer, Color As RGBA, ColorLine As RGBA)
    
    On Error GoTo Engine_Draw_Box_Border_Err
    If Not OverlapRect(RenderCullingRect, x, y, Width, Height) Then Exit Sub

    Call Engine_Draw_Box(x, y, Width, Height, Color)

    Call Engine_Draw_Box(x, y, Width, 1, ColorLine)
    Call Engine_Draw_Box(x, y + Height, Width, 1, ColorLine)
    Call Engine_Draw_Box(x, y, 1, Height, ColorLine)
    Call Engine_Draw_Box(x + Width, y, 1, Height, ColorLine)

    
    Exit Sub

Engine_Draw_Box_Border_Err:
    Call RegistrarError(Err.Number, Err.Description, "engine.Engine_Draw_Box_Border", Erl)
    Resume Next
    
End Sub

Public Sub DibujarNPC(PicBox As PictureBox, ByVal Head As Integer, ByVal Body As Integer, Optional ByVal Heading As Byte = 3)

    On Error GoTo DibujarNPC_Err

    Dim x As Integer
    Dim y As Integer

    Dim bodyGrh As Long, headGrh As Long

    If Body Then
        bodyGrh = BodyData(Body).Walk(Heading).GrhIndex
    End If

    If Head Then
        headGrh = HeadData(Head).Head(Heading).GrhIndex
    End If
    
    If bodyGrh Then
        bodyGrh = GrhData(bodyGrh).Frames(1)
    
        x = (PicBox.ScaleWidth - GrhData(bodyGrh).pixelWidth) \ 2
        y = min(PicBox.ScaleHeight - GrhData(bodyGrh).pixelHeight + BodyData(Body).HeadOffset.y \ 2, (PicBox.ScaleHeight - GrhData(bodyGrh).pixelHeight) \ 2)
        
        Call Grh_Render_To_Hdc(PicBox, bodyGrh, x, y, False, RGB(11, 11, 11))

        If headGrh Then
            headGrh = GrhData(headGrh).Frames(1)
            
            x = (PicBox.ScaleWidth - GrhData(headGrh).pixelWidth) \ 2 + 1
            y = y + GrhData(bodyGrh).pixelHeight - GrhData(headGrh).pixelHeight + BodyData(Body).HeadOffset.y

            Call Grh_Render_To_HdcSinBorrar(PicBox, headGrh, x, y, False)
        End If
        
    End If
    
    Exit Sub

DibujarNPC_Err:
    Call RegistrarError(Err.Number, Err.Description, "FrmQuestInfo.DibujarNPC", Erl)
    Resume Next
    
End Sub

Public Function GetTickCount() As Long
    ' Devolvemos el valor absoluto de la cantidad de ticks que paso desde que prendimos la PC
    
    GetTickCount = (timeGetTime And &H7FFFFFFF)
    
End Function
