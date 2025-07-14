VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cfVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : cfVolume
' Author    : beededea
' Date      : 18/10/2024
' Purpose   : This is the class for the gauge form and all the elements upon it.
'---------------------------------------------------------------------------------------

'@IgnoreModule IntegerDataType, ModuleWithoutFolder
Option Explicit

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Private Declare Function SafeArrayGetDim Lib "oleaut32.dll" (ByRef saArray() As Any) As Long

' Declare volumeForm as an RC form to be available throughout by calling fVolume as defined as this class in modMain

Public WithEvents volumeForm As cWidgetForm
Attribute volumeForm.VB_VarHelpID = -1

Public collPSDNonUIElements As cCollection, FX As Long, FY As Long, FZ As Single  'available on the outside, even after the Form was closed

Private pvtPSDWidth As Long
Private pvtPSDHeight As Long

Private mMute As Integer ' property
Private mSystemAudioLevel As Currency ' property
Private mResetAudio As Boolean

Private RenderDev As cMMDevice
Private WithEvents SystemVolume As cSimpleAudioVolume 'use cSimpleAudioVolume with events
Attribute SystemVolume.VB_VarHelpID = -1

'Private lblPercent As cwLabel
'Private lblPercentShadow As cwLabel

'Slider-related Private-Vars
Private WithEvents SliderSet As cWidgetBase
Attribute SliderSet.VB_VarHelpID = -1
Private sldFixedTop As Double
Private sldXMax As Double
Private sldXMin As Double
Private sldX As Double
Private mVolumePerc As Currency

Private mOpacity As Long

Public WithEvents tmrSampleAudioVolume As cTimer
Attribute tmrSampleAudioVolume.VB_VarHelpID = -1
Public WithEvents tmrSampleAudioMute As cTimer
Attribute tmrSampleAudioMute.VB_VarHelpID = -1

'---------------------------------------------------------------------------------------
' Procedure : initialisePrivateVars
' Author    : beededea
' Date      : 12/02/2025
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub initialisePrivateVars()
   On Error GoTo initialisePrivateVars_Error

    pvtPSDWidth = 0
    pvtPSDHeight = 0
    
   On Error GoTo 0
   Exit Sub

initialisePrivateVars_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure initialisePrivateVars of Class Module cfVolume"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : Class_Initialize
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   : Constructor (Initialise)
'---------------------------------------------------------------------------------------
'
Private Sub Class_Initialize()
   
    On Error GoTo Class_Initialize_Error
    
    Call initialisePrivateVars
    
    Set collPSDNonUIElements = New_c.Collection(False)
    
    Call obtainAudioDevice
    
    Set tmrSampleAudioVolume = Nothing
    Set tmrSampleAudioVolume = New_c.Timer(1000, True) ' interval is one second.
    
    Set tmrSampleAudioMute = Nothing
    Set tmrSampleAudioMute = New_c.Timer(1000, True) ' interval is one second.
    
   On Error GoTo 0
   Exit Sub

Class_Initialize_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Class_Initialize of Class Module cfVolume"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : obtainAudioDevice
' Author    : beededea
' Date      : 22/02/2025
' Purpose   : obtain the audio device and set the SystemVolume object as an RC6 cSimpleAudioVolume object
'---------------------------------------------------------------------------------------
'
Private Sub obtainAudioDevice()
        
   On Error GoTo obtainAudioDevice_Error

    'retrieve (and Set) the Default-Renderdevice
    Set RenderDev = New_c.MMDeviceEnumerator.GetDefaultAudioEndpoint(eRender, eMultimedia)
    
    'retrieve (and Set) a cSimpleAudioVolume instance, to get access to Volume-Handling
    Set SystemVolume = RenderDev.Activate(AudioEndpointVolume)

   On Error GoTo 0
   Exit Sub

obtainAudioDevice_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure obtainAudioDevice of Class Module cfVolume"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : InitFromPSD
' Author    : olaf schmidt and me
' Date      : 31/07/2023
' Purpose   :  Using a previously populated Cairo imageList with keyed entries  it creates a surface from each and every PSD layer
'              in the PSD file.
'              It adds path X,Y and alpha properties to the excluded collection
'              It then creates widgets with corresponding keynames, locations and sizes as per the original PSD for each layer
'              It creates an instance of the cwOverlay class and populates it with the excluded items that will be rendered in the overlay
'              The overlay comprises items that are non-clickable and will not generate events, ie. animated gauge hands, pendulum &c.
'---------------------------------------------------------------------------------------
'
Public Sub InitFromPSD(ByRef PSD_FileNameOrByteArray As Variant)
    
    Const AlphaWithTaskbarEntry As Integer = 6
    Const AlphaNoTaskbarEntry As Integer = 7
  
    Dim I As Long: I = 0
    Dim W As cWidgetBase
    Dim someOpacity As Double: someOpacity = 0
    
    On Error GoTo InitFromPSD_Error
    
    someOpacity = Val(gblOpacity) / 100

    'create the Top-Level-Form
    Set volumeForm = Cairo.WidgetForms.Create(IIf(App.LogMode, AlphaNoTaskbarEntry, AlphaWithTaskbarEntry), gblWidgetName, True, 1, 1)
        volumeForm.WidgetRoot.BackColor = -1 ' transparent
 
    With New_c.SimplePSD(PSD_FileNameOrByteArray)  'create a new PSD-Parser.instance (and load the passed content)
        pvtPSDWidth = .Width
        pvtPSDHeight = .Height       'store the original Psd-Pixel-Width/Height in Private vars (as the base from which we calculate the zoomed Form-Width/Height)
    
        For I = 0 To .LayersCount - 1 ' loop through each of the Layers in the PSD
            If .LayerByteSize(I) Then  ' check this is a true Alpha-Surface-Layer and not just a PSD layer 'group'
                If .LayerAlphaPercent(I) > 0 Then ' only handles layers that have an opacity greater than 0 - need to note this for the future, this will cause a problem!
                    
                    'add each current Layer path and surface object into the global ImageList collection (using LayerPath as the ImageKey)
                    Cairo.ImageList.AddSurface .LayerPath(I), .LayerSurface(I)
                    
                    ' check if each layer is in the layer exclude list, if it IS then we add it to a collection for non UI elements (ie. do not create Widgets)
                    If collPSDNonUIElements.Exists(.LayerPath(I)) Then
                        
                        'we add layer info. (used later in cwOverlay) to the excluded layers that will form the overlay.
                        collPSDNonUIElements(.LayerPath(I)) = Array(.LayerX(I), .LayerY(I), someOpacity)  'here we update the so far empty slots with the PSD-offsets
                                    
                    Else
                    
                        'create a widget instance for all layers in the PSD, excluding any layers entered into the exclude-list
                        
                        Set W = volumeForm.Widgets.Add(New cwAlphaImg, LCase$(.LayerPath(I)), .LayerX(I), .LayerY(I), .LayerWidth(I), .LayerHeight(I)).Widget
                        
                        W.ImageKey = W.Key 'W.Key equals ImageList-Key, set above - and LayerPath(i) at this point ... set it also as the ImageKey of our new created Widget

                        W.Alpha = 0
                    
                        ' note: the clickable layers characteristics are set in adjustMainControls
                        
                        ' all non-clickable Layer-Widgets will be -1 or "non-hoverable" and "fully click-through"
                        W.HoverColor = -1 ' task: might change this later when new ver or RC6 arrives
                        If gblGaugeTooltips = "1" Then W.ToolTip = "Ctrl + mouse scrollwheel up/down to resize, you can also drag me to a new position."
                        W.MousePointer = IDC_SIZEALL
                        
                    End If
                End If
            End If
        Next I
    End With '<-- the Parser-instance will be destroyed here (freeing the Memory, the internal PSD-Layers have occupied)
        
    ' create a overlay here just for a single custom widget that contains just the animated elements.
    Set overlayWidget = volumeForm.Widgets.Add(New cwOverlay, "Overlay", 0, 0, pvtPSDWidth, pvtPSDHeight)
    Set overlayWidget.cCollOverlayImages = collPSDNonUIElements 'make the excluded PSD-Surface-Paths known to the Overlay-Widget
        overlayWidget.cxPerc = 0.4940476 'define the relative center (within cwOverlay, which has the same size as the Form and the PSD...) this is used for clock-hand-positioning
        overlayWidget.cyPerc = 0.4727272
    
    'the next 5 lines will init a "moving Slider" (which was defined as "sliderset"-layer in the PSD)
    Set SliderSet = volumeForm.Widgets("sliderset").Widget 'store the already (in PSD-Loop created) Widget-reference in an explicit "WithEvents-Variable"
        
    With SliderSet
        .Moveable = True 'make it "moveable"
        .HoverColor = 0 ' set the hover colour to grey - this may change later with new RC6
        .MousePointer = IDC_HAND
        .Tag = 0.01
        .Alpha = someOpacity
    End With
    
    sldFixedTop = SliderSet.Top   'and store the initial position-data for y-direction (which will not be changed)
    sldXMax = SliderSet.Left      'as well as the x-direction related pos (which is initially at the "100%-Volume, right-endswitch" )
    sldXMin = sldXMax - 300       'the Min-Value is simply the Max-Value minus Slider-Range (the 300pxl-range was determined by "probing")

    On Error GoTo 0
   Exit Sub

InitFromPSD_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InitFromPSD of Class Module cfVolume"
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : volumeForm_BubblingEvent
' Author    : beededea
' Date      : 31/07/2023
' Purpose   : central handling of all Click-Events for all pre-defined PSD layer widgets
'---------------------------------------------------------------------------------------
'
Private Sub volumeForm_BubblingEvent(ByRef Sender As Object, ByRef EventName As String, ByRef p1 As Variant, ByRef p2 As Variant, ByRef P3 As Variant, ByRef P4 As Variant, ByRef P5 As Variant, ByRef P6 As Variant, ByRef P7 As Variant)
    Dim stdMessage As String: stdMessage = vbNullString
    Dim stdSubject As String: stdSubject = vbNullString
    
    On Error GoTo volumeForm_BubblingEvent_Error

    If gblIgnoreMouse = "1" Then Exit Sub
    
    stdMessage = "Right Click to open the menu and the preferences. CTRL+ Mouse scrollwheel UP/DOWN to resize. "
    stdSubject = "You can turn off the balloon tooltips in the preferences."
    
    If EventName = "W_MouseMove" Then
        Select Case LCase$(Sender.Widget.Key)
            Case "helppin"
                If gblPrefsTooltips = "0" Then CreateToolTip volumeForm.hWnd, "This button opens the one-page help file for this widget. There is a further detailed CHM help file available from the right-click menu that contains all the information that you need to run this widget.", _
                     TTIconInfo, gblWidgetName & " Help Pin Help", , , , True
            Case "lockingpin"
                If gblPrefsTooltips = "0" Then CreateToolTip volumeForm.hWnd, "Click me to lock the widget in place. This lock can be released by clicking again on the locking pin, or in the preferences. A locking sound will denote the button's operation.", _
                     TTIconInfo, gblWidgetName & " Locking pin Help", , , , True
            Case "bell"
                If gblPrefsTooltips = "0" Then CreateToolTip volumeForm.hWnd, "This is the mute button. Click me to toggle sound on/off.", _
                     TTIconInfo, gblWidgetName & " Sound Toggle Help", , , , True
            Case "lockingpinunlocked"
                If gblPrefsTooltips = "0" Then CreateToolTip volumeForm.hWnd, "Click me to lock the widget in place. This lock can be released by clicking again on the locking pin, or in the preferences. A locking sound will denote the button's operation.", _
                     TTIconInfo, gblWidgetName & " Locking pin Help", , , , True
            Case "indicatorred"
                If gblPrefsTooltips = "0" Then CreateToolTip volumeForm.hWnd, "This is the mute indicator. A green lamp means that the sound is unmuted. Red lamp means that mute is ON.", _
                     TTIconInfo, gblWidgetName & " Locking pin Help", , , , True
            Case "indicatorgreen"
                If gblPrefsTooltips = "0" Then CreateToolTip volumeForm.hWnd, "This is the mute indicator. Sound is unmuted.", _
                     TTIconInfo, gblWidgetName & " Locking pin Help", , , , True
            Case "sliderset"
                If gblPrefsTooltips = "0" Then CreateToolTip volumeForm.hWnd, "This is the Volume Slider. Click and drag me left/right to change the volume.", _
                     TTIconInfo, gblWidgetName & " Locking pin Help", , , , True

            Case "speaker", "cable", "bar", "pipes", "cablewheelset"
                If gblPrefsTooltips = "0" Then CreateToolTip volumeForm.hWnd, stdMessage _
                    & vbCrLf & stdSubject, _
                    TTIconInfo, gblWidgetName & " Configuration Help", , , , True
        End Select
    End If
    
    If EventName = "W_MouseDown" Then
        Select Case LCase$(Sender.Widget.Key)
            Case "lockingpin"
                Call lockingpin_mouseDown(Sender)
            Case "helppin"
                Call helppin_mouseDown(Sender)
            Case "lockingpinunlocked"
                Call lockingpin_mouseDown(Sender)
            Case "bell"
                Call bell_mouseDown(Sender)
            Case "sliderset"
                overlayWidget.Locked = True
                fVolume.tmrSampleAudioVolume.Enabled = False
            Case Else
                Debug.Print Sender.Widget.Key
        End Select
    End If
    
    If EventName = "W_MouseUp" Then
        Select Case LCase$(Sender.Widget.Key)
            Case "bell"
                Call bell_mouseUp(Sender)
            Case "sliderset"
                overlayWidget.Locked = False
                fVolume.tmrSampleAudioVolume.Enabled = True
            Case Else
                Debug.Print Sender.Widget.Key
        End Select
    End If
            
   On Error GoTo 0
   Exit Sub

volumeForm_BubblingEvent_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure volumeForm_BubblingEvent of Class Module cfVolume"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : lockingpin_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub lockingpin_mouseDown(ByRef Sender As Object)
   On Error GoTo lockingpin_mouseDown_Error
   
    Dim fileToPlay As String: fileToPlay = vbNullString

    fileToPlay = "lock.wav"
    If gblEnableSounds = "1" And fFExists(App.Path & "\resources\sounds\" & fileToPlay) Then
        playSound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    If overlayWidget.Locked = False Then
        Sender.Widget.Alpha = 0
        fVolume.volumeForm.Widgets("lockingpinunlocked").Widget.Alpha = Val(gblOpacity) / 100
    Else
        Sender.Widget.Alpha = Val(gblOpacity) / 100
        fVolume.volumeForm.Widgets("lockingpinunlocked").Widget.Alpha = 0
    End If
    Sender.Widget.Refresh ' this removes the delay in the lockingpin depressing
    Call toggleWidgetLock
    
   On Error GoTo 0
   Exit Sub

lockingpin_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lockingpin_mouseDown of Class Module cfVolume"

End Sub
'---------------------------------------------------------------------------------------
' Procedure : helppin_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub helppin_mouseDown(ByRef Sender As Object)
   On Error GoTo helppin_mouseDown_Error
   
    Dim fileToPlay As String: fileToPlay = vbNullString

    fileToPlay = "lock.wav"
    If gblEnableSounds = "1" And fFExists(App.Path & "\resources\sounds\" & fileToPlay) Then
        playSound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
'    If overlayWidget.Locked = False Then
'        Sender.Widget.Alpha = 0
'        fVolume.volumeForm.Widgets("helppinunlocked").Widget.Alpha = Val(gblOpacity) / 100
'    Else
'        Sender.Widget.Alpha = Val(gblOpacity) / 100
'        fVolume.volumeForm.Widgets("helppinunlocked").Widget.Alpha = 0
'    End If
    Sender.Widget.Refresh ' this removes the delay in the helppin depressing
    
    Call helpSplash
    
   On Error GoTo 0
   Exit Sub

helppin_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure helppin_mouseDown of Class Module cfVolume"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : bell_mouseDown
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub bell_mouseDown(ByRef Sender As Object)
   On Error GoTo bell_mouseDown_Error
   
    Dim fileToPlay As String: fileToPlay = vbNullString
    
    tmrSampleAudioMute.Enabled = False
    
    fileToPlay = "buzzer.wav"

    If Mute = 1 Then
        Mute = 0
    Else
        Mute = 1
    End If
    
    If gblEnableSounds = "1" And fFExists(App.Path & "\resources\sounds\" & fileToPlay) Then
        playSound App.Path & "\resources\sounds\" & fileToPlay, ByVal 0&, SND_FILENAME Or SND_ASYNC
    End If
    
    Sender.Widget.Alpha = 0
    Sender.Widget.Refresh ' this removes the delay in the bell depressing
    
    tmrSampleAudioMute.Enabled = True
    
   On Error GoTo 0
   Exit Sub

bell_mouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure bell_mouseDown of Class Module cfVolume"

End Sub


'---------------------------------------------------------------------------------------
' Procedure : bell_mouseUp
' Author    : beededea
' Date      : 03/08/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub bell_mouseUp(ByRef Sender As Object)
   On Error GoTo bell_mouseUp_Error
    
    Sender.Widget.Alpha = Val(gblOpacity) / 100
    Sender.Widget.Refresh ' this removes the delay in the bell depressing
    
   On Error GoTo 0
   Exit Sub

bell_mouseUp_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure bell_mouseUp of Class Module cfVolume"

End Sub



  



'---------------------------------------------------------------------------------------
' Procedure : volumeForm_DblClick
' Author    : beededea
' Date      : 05/05/2023
' Purpose   : If a double-click command has been selected in the prefs then this will carry it out
'---------------------------------------------------------------------------------------
'
Private Sub volumeForm_DblClick()
    Dim userprof As String: userprof = vbNullString
    Dim thisCommand As String: thisCommand = vbNullString
    
    On Error GoTo volumeForm_DblClick_Error

    '    If gblIgnoreMouse = "1" Then Exit Sub

    If gblWidgetFunctions = "0" Or gblIgnoreMouse = "1" Then Exit Sub
    
    If LTrim$(gblDblClickCommand) = vbNullString Then Exit Sub
    
    thisCommand = gblDblClickCommand
        
    If InStr(thisCommand, "%userprofile%") Then
        userprof = Environ$("USERPROFILE")
        thisCommand = Replace(thisCommand, "%userprofile%", userprof)
    End If
    
    ' .91 DAEB 08/12/2022 frmMain.frm SteamyDock responds to %systemroot% environment variables during runCommand
    If InStr(thisCommand, "%systemroot%") Then
        userprof = Environ$("SYSTEMROOT")
        thisCommand = Replace(thisCommand, "%systemroot%", userprof)
    End If
    
    If gblSHIFT_1 = True Then
        gblSHIFT_1 = False
        Call ShellExecute(volumeForm.hWnd, "Open", gblOpenFile, vbNullString, App.Path, 1)
    Else
        Call ShellExecute(volumeForm.hWnd, "runas", thisCommand, vbNullString, App.Path, 1)
    End If

    On Error GoTo 0
    Exit Sub

volumeForm_DblClick_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure volumeForm_DblClick of Class Module cfVolume"
            Resume Next
          End If
    End With
End Sub


'---------------------------------------------------------------------------------------
' Procedure : volumeForm_KeyDown
' Author    : beededea
' Date      : 01/06/2019
' Purpose   : get F5 and SHIFT keypresses and key downs for the main toggles H,A,S,M,P,T,D,W,B & 1-5
'---------------------------------------------------------------------------------------
'
Private Sub volumeForm_KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
    On Error GoTo volumeForm_KeyDown_Error

    Call getKeyPress(KeyCode, Shift) ' will not catch Shift or CTRL keypresses on their own, must be combination with another key.

    On Error GoTo 0
    Exit Sub

volumeForm_KeyDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure volumeForm_KeyDown of Class Module cfVolume"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : volumeForm_MouseUp
' Author    : beededea
' Date      : 05/05/2023
' Purpose   :    if this is a multi monitor system then:
'                we note the monitor ID at volumeForm form_load and store as the volumeFormMonitorID, for widgetPrefs widgetPrefsMonitorID
'                on gaugeFromMouseUp (drag complete) we sample the monitor ID again -only works, for the prefs form we have to have a 200ms timer that tests for form .left and .top x,y change (no MOVED event and mouse UP does not work on the form itself when the titlebar dragged)
'                if the monitor has changed then sample the physical monitor resolution
'                if the resolution is different then calculate new size proportion
'                alter the size of the form to correspond to the monitor
'---------------------------------------------------------------------------------------
'
Private Sub volumeForm_MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)

    On Error GoTo volumeForm_MouseUp_Error
    
    frmTimer.tmrScreenResolution.Interval = 4500

    ' save the form x,y position whenever the globe/form is dragged
    Call saveMainRCFormPosition

    Call positionRCFormByMonitorSize
    
    On Error GoTo 0
    Exit Sub

volumeForm_MouseUp_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure volumeForm_MouseUp of Class Module cfVolume"
            Resume Next
          End If
    End With
End Sub





'---------------------------------------------------------------------------------------
' Procedure : volumeForm_MouseMove
' Author    : beededea
' Date      : 31/07/2023
' Purpose   : move the whole form on a mousDown and drag
'---------------------------------------------------------------------------------------
'
Private Sub volumeForm_MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
  Static x0 As Single, y0 As Single: If Button = 0 Then x0 = x: y0 = y 'just store the offsets when no button is down
  
   On Error GoTo volumeForm_MouseMove_Error
    
    If overlayWidget.Locked = True Then Exit Sub
    If gblIgnoreMouse = "1" Then Exit Sub
    If gblMenuOccurred = True Then
        gblMenuOccurred = False
        Exit Sub
    End If

    If Button = vbLeftButton And Not volumeForm.ActiveWidget Is Nothing Then 'Form-Dragging (under certain conditions)

      FX = (x - x0) * volumeForm.WidgetRoot.Zoom + volumeForm.Left
      FY = (y - y0) * volumeForm.WidgetRoot.Zoom + volumeForm.Top
      volumeForm.Move FX, FY
      
      Debug.Print "FX " & FX
      Debug.Print "FY " & FY
      
    End If


   On Error GoTo 0
   Exit Sub

volumeForm_MouseMove_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure volumeForm_MouseMove of Class Module cfVolume"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : volumeForm_MouseWheel
' Author    : beededea
' Date      : 31/07/2023
' Purpose   : Handles rotation of the mousewheel, with CTRL key pressed, making the widget smaller or larger. By default, rotate forward positive = smaller
'---------------------------------------------------------------------------------------
'
Private Sub volumeForm_MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Single, ByVal Ypos As Single)
   Dim RotationDirection As Long: RotationDirection = 0
   
   On Error GoTo volumeForm_MouseWheel_Error

    Rotation = Rotation / 2

    If gblScrollWheelDirection = "1" Then
        If Rotation > 0 Then
            RotationDirection = 0 - Abs(Rotation)
        Else
            RotationDirection = Abs(Rotation)
        End If
    Else
        RotationDirection = Rotation
    End If
    
   'If MouseKeys = 8 Then Call AdjustZoom(FZ + RotationDirection / 2400)
   If MouseKeys = 8 Then widgetPrefs.GaugeSize = (FZ + RotationDirection / 2400) 'change the Zoom only, when the Ctrl-Key is down

   On Error GoTo 0
   Exit Sub

volumeForm_MouseWheel_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure volumeForm_MouseWheel of Class Module cfVolume"
End Sub


'---------------------------------------------------------------------------------------
' Procedure : AdjustZoom
' Author    : olaf schmidt
' Date      : 31/07/2023
' Purpose   : Adjust the size of the whole gauge, anchoring the gauge top left, storing the values for later use
'---------------------------------------------------------------------------------------
'
Public Sub AdjustZoom(ByVal NewZoom As Single)
   On Error GoTo AdjustZoom_Error

    FZ = NewZoom
    If FZ < 0.05 Then FZ = 0.05 Else If FZ > 3 Then FZ = 3 'keep FZ within a sane ZoomInterval
    volumeForm.WidgetRoot.Zoom = FZ '* volumeForm.WidgetRoot.CurrentMonitor.Zoom '<-- if the multiplicator is activated, we'd have DPI-awareness on high-res monitors
    volumeForm.Move volumeForm.Left, volumeForm.Top, pvtPSDWidth * volumeForm.WidgetRoot.Zoom, pvtPSDHeight * volumeForm.WidgetRoot.Zoom
    
    ' when resized from the gauge scroll up/down it needs to write it back so the size is correct when prefs opened.
    gblGaugeSize = CStr(NewZoom * 100)
    
    If widgetPrefs.IsLoaded Then widgetPrefs.sliGaugeSize.Value = Val(gblGaugeSize) ' ensure that the prefs size slider matches when zoom is modified without using the slider
    
    If gaugeMonitorStruct.IsPrimary = True Then
        gblGaugePrimaryHeightRatio = CStr(NewZoom)
    Else
        gblGaugeSecondaryHeightRatio = CStr(NewZoom)
    End If

   On Error GoTo 0
   Exit Sub

AdjustZoom_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure AdjustZoom of Class Module cfVolume"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : volumeForm_MouseDown
' Author    :
' Date      : 09/05/2023
' Purpose   : A click on the gauge form, test screen resolution and handle right mouse click for the menu
'---------------------------------------------------------------------------------------
'
Private Sub volumeForm_MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef x As Single, ByRef y As Single)
   On Error GoTo volumeForm_MouseDown_Error

    'essential code to test for monitor resolution before and after a mouse-down pick up and drag to another monitor
    frmTimer.tmrScreenResolution.Enabled = False
    frmTimer.tmrScreenResolution.Interval = 200
    frmTimer.tmrScreenResolution.Enabled = True
    
    If Button = vbRightButton Then
        gblMenuOccurred = True
        Call menuForm.PopupMenu(menuForm.mnuMainMenu)
    Else
        If gblIgnoreMouse = "1" Then Exit Sub
        ' so we pass it through
    End If

   On Error GoTo 0
   Exit Sub

volumeForm_MouseDown_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure volumeForm_MouseDown of Class Module cfVolume"
End Sub



'---------------------------------------------------------------------------------------
' Procedure : IsVisible
' Author    : beededea
' Date      : 08/05/2023
'
' Purpose   : Useful property to determine by value as to whether this form is loaded,
'             allowing external checks to the form to determine whether it is loaded,
'             WITHOUT activating the form automatically.
'---------------------------------------------------------------------------------------
'
Public Property Get IsVisible() As Boolean
    On Error GoTo IsVisible_Error

    If volumeForm.WindowState = vbNormal Then
        IsVisible = volumeForm.Visible
    Else
        IsVisible = False
    End If

    On Error GoTo 0
    Exit Property

IsVisible_Error:

    With Err
         If .Number <> 0 Then
            MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsVisible of Class Module cfVolume"
            Resume Next
          End If
    End With
End Property



'---------------------------------------------------------------------------------------
' Property  : Opacity
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property to determine (by value) the opacity of the elements within this class
'---------------------------------------------------------------------------------------
'
Public Property Get opacity() As String
   On Error GoTo opacityGet_Error

   opacity = mOpacity

   On Error GoTo 0
   Exit Property

opacityGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Opacity of Class Module cfVolume"
End Property
'---------------------------------------------------------------------------------------
' Property  : Opacity
' Author    : beededea
' Date      : 17/05/2023
' Purpose   : property to determine (by value) the opacity of the elements within this class
'---------------------------------------------------------------------------------------
'
Public Property Let opacity(ByVal newValue As String)
   On Error GoTo opacityLet_Error

   If mOpacity <> newValue Then mOpacity = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

opacityLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Opacity of Class Module cfVolume"
End Property


' ----------------------------------------------------------------
' Procedure Name: SliderSet_Moving
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 26/05/2024
' ----------------------------------------------------------------
Private Sub SliderSet_Moving() 'this is an explcit Event-Handler on the single Widget-instance (derived in the routine above)
    On Error GoTo SliderSet_Moving_Error
     sldX = SliderSet.Left
     If sldX < sldXMin Then sldX = sldXMin Else If sldX > sldXMax Then sldX = sldXMax 'keep sldX "in range"
    
     VolumePerc = (sldX - sldXMin) / (sldXMax - sldXMin) 'assign the new Value to the Public Prop, to trigger all the necessary refreshs
    
    On Error GoTo 0
    Exit Sub

SliderSet_Moving_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SliderSet_Moving, line " & Erl & "."

End Sub
'---------------------------------------------------------------------------------------
' Property  : SystemAudioLevel
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let SystemAudioLevel(ByVal newValue As Currency)
   On Error GoTo SystemAudioLevelLet_Error

   If mSystemAudioLevel <> newValue Then mSystemAudioLevel = newValue Else Exit Property

   On Error GoTo 0
   Exit Property

SystemAudioLevelLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property SystemAudioLevel of Class Module cfVolume"
End Property

' ----------------------------------------------------------------
' Procedure Name: SystemAudioLevel
' Purpose:
' Procedure Kind: Property (Get)
' Procedure Access: Public
' Return Type: Currency
' Author: beededea
' Date: 26/05/2024
' ----------------------------------------------------------------
Public Property Get SystemAudioLevel() As Currency
    
    On Error GoTo SystemAudioLevel_Error
    
    mSystemAudioLevel = SystemVolume.GetMasterVolume
    SystemAudioLevel = mSystemAudioLevel

    
    On Error GoTo 0
    Exit Property

SystemAudioLevel_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SystemAudioLevel, line " & Erl & "."

End Property

'---------------------------------------------------------------------------------------
' Property  : Mute
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let Mute(ByVal newValue As Integer)
   On Error GoTo MuteLet_Error
   
    If mMute <> newValue Then mMute = newValue Else Exit Property
    
    Mute = mMute
   
    SystemVolume.SetMute mMute
    
    Call setLamps
   
   On Error GoTo 0
   Exit Property

MuteLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Mute of Class Module cfVolume"
End Property

'---------------------------------------------------------------------------------------
' Property  : Mute
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get Mute() As Integer
   On Error GoTo MuteGet_Error

    Mute = mMute
    
    SystemVolume.SetMute mMute
    
    Call setLamps

   On Error GoTo 0
   Exit Property

MuteGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property Mute of Class Module cfVolume"
End Property


' ----------------------------------------------------------------
' Procedure Name: VolumePerc
' Purpose:
' Procedure Kind: Property (Get)
' Procedure Access: Public
' Return Type: Currency
' Author: beededea
' Date: 26/05/2024
' ----------------------------------------------------------------
Public Property Get VolumePerc() As Currency
    
    On Error GoTo VolumePerc_Error
    
    VolumePerc = mVolumePerc
    
    On Error GoTo 0
    Exit Property

VolumePerc_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure VolumePerc, line " & Erl & "."

End Property

' ----------------------------------------------------------------
' Procedure Name: VolumePerc
' Purpose:
' Procedure Kind: Property (Let)
' Procedure Access: Public
' Parameter RHS (Currency):
' Author: beededea
' Date: 26/05/2024
' ----------------------------------------------------------------
Public Property Let VolumePerc(ByVal RHS As Currency)
    On Error GoTo VolumePerc_Error
    
    If RHS < 0 Then RHS = 0 Else If RHS > 1 Then RHS = 1 'keep RHS "in range"
    sldX = RHS * (sldXMax - sldXMin) + sldXMin
    overlayWidget.sldCurWidth = sldXMax - sldX 'inform the Overlay about the new current "Cable-Width"
    SliderSet.Move sldX, sldFixedTop 'move the SliderSet-widget accordingly
    If mVolumePerc <> RHS Then mVolumePerc = RHS Else Exit Property  'only refresh, when there was a change to the previous value
    
    If Not Cairo.ToolTipForm Is Nothing Then Cairo.ToolTipForm.Visible = False
    volumeForm.Refresh
    
    Call HandleVolumeChangeForHardware
    
    On Error GoTo 0
    Exit Property

VolumePerc_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure VolumePerc, line " & Erl & "."

End Property

'---------------------------------------------------------------------------------------
' Property  : resetAudio
' Author    : beededea
' Date      : 10/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Let resetAudio(ByVal newValue As Boolean)
   On Error GoTo resetAudioLet_Error
   
    If mResetAudio <> newValue Then mResetAudio = newValue Else Exit Property
    
    resetAudio = mResetAudio
    
    If mResetAudio = True Then
        Call obtainAudioDevice
        mResetAudio = False
    End If
   
   On Error GoTo 0
   Exit Property

resetAudioLet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property resetAudio of Class Module cfVolume"
End Property

'---------------------------------------------------------------------------------------
' Property  : resetAudio
' Author    : beededea
' Date      : 17/05/2023
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Property Get resetAudio() As Boolean
   On Error GoTo resetAudioGet_Error

    resetAudio = mResetAudio

   On Error GoTo 0
   Exit Property

resetAudioGet_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in Property resetAudio of Class Module cfVolume"
End Property
' ----------------------------------------------------------------
' Procedure Name: HandleVolumeChangeForHardware
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 27/05/2024
' ----------------------------------------------------------------
Private Sub HandleVolumeChangeForHardware()

    On Error GoTo HandleVolumeChangeForHardware_Error
   
    SystemVolume.SetMasterVolume VolumePerc
    
    On Error GoTo 0
    Exit Sub

HandleVolumeChangeForHardware_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure HandleVolumeChangeForHardware, line " & Erl & "."

End Sub


    
'---------------------------------------------------------------------------------------
' Procedure : tmrSampleAudioVolume_Timer
' Author    :
' Date      : 30/07/2023
' Purpose   : regular check to reflect changes to system volume
'---------------------------------------------------------------------------------------
'
Public Sub tmrSampleAudioVolume_Timer()

    On Error GoTo tmrSampleAudioVolume_Timer_Error
    
    If mSystemAudioLevel = SystemVolume.GetMasterVolume Then Exit Sub
    
    mSystemAudioLevel = SystemVolume.GetMasterVolume
    
    sldX = mSystemAudioLevel * (sldXMax - sldXMin) + sldXMin
    overlayWidget.sldCurWidth = sldXMax - sldX 'inform the Overlay about the new current "Cable-Width"
    SliderSet.Move sldX, sldFixedTop 'move the SliderSet-widget accordingly
    
    fVolume.volumeForm.Refresh

   On Error GoTo 0
   Exit Sub

tmrSampleAudioVolume_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrSampleAudioVolume_Timer of Class Module cwOverlayTemp"
End Sub

    
'---------------------------------------------------------------------------------------
' Procedure : tmrSampleAudioMute_Timer
' Author    :
' Date      : 30/07/2023
' Purpose   : regular check to reflect changes to system volume
'---------------------------------------------------------------------------------------
'
Public Sub tmrSampleAudioMute_Timer()

    On Error GoTo tmrSampleAudioMute_Timer_Error
    
    If mMute = SystemVolume.GetMute Then Exit Sub
    
    mMute = SystemVolume.GetMute
    
    Call setLamps
    
   On Error GoTo 0
   Exit Sub

tmrSampleAudioMute_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure tmrSampleAudioMute_Timer of Class Module cwOverlayTemp"
End Sub


' ----------------------------------------------------------------
' Procedure Name: setLamps
' Purpose:
' Procedure Kind: Sub
' Procedure Access: Private
' Author: beededea
' Date: 28/05/2024
' ----------------------------------------------------------------
Private Sub setLamps()
    
    On Error GoTo setLamps_Error
    If mMute = 1 Then
        fVolume.volumeForm.Widgets("indicatorgreen").Widget.Alpha = 0
        fVolume.volumeForm.Widgets("indicatorred").Widget.Alpha = Val(gblOpacity) / 100
    Else
        fVolume.volumeForm.Widgets("indicatorgreen").Widget.Alpha = Val(gblOpacity) / 100
        fVolume.volumeForm.Widgets("indicatorred").Widget.Alpha = 0
    End If

    fVolume.volumeForm.Refresh
    
    On Error GoTo 0
    Exit Sub

setLamps_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure setLamps, line " & Erl & "."

End Sub


