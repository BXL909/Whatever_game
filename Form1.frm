VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Option Explicit

Dim binit As Boolean 'A flag that states whether we've initialised or not. If the initialisation is successful this changes to true

Dim dx As New DirectX7 'DirectDraw is created from this
Dim dd As DirectDraw7
Dim primary As DirectDrawSurface7 'This surface represents the screen
Dim Mainsurf As DirectDrawSurface7
Dim Menusurf As DirectDrawSurface7
Dim Levelsurf As DirectDrawSurface7
Dim backbuffer As DirectDrawSurface7 'backbuffer
Dim ddsd1 As DDSURFACEDESC2 'describes the primary surface
Dim ddsd2 As DDSURFACEDESC2 'describes the bitmap
Dim ddsd3 As DDSURFACEDESC2 'describes size of the screen
Dim ddsd4 As DDSURFACEDESC2 'describes the menu bitmap
Dim ddsd5 As DDSURFACEDESC2 'describes the menu bitmap
Dim brunning As Boolean 'flag that states whether or not the main game loop is running.
Dim CurModeActiveStatus As Boolean 'checks we still have the correct display mode
Dim bRestore As Boolean 'If not the correct display mode this flag states that we need to restore the display mode
Dim Ds As DirectSound
Dim DsBuffer As DirectSoundBuffer
Dim DsBuffer2 As DirectSoundBuffer
Dim DsBuffer3 As DirectSoundBuffer
Dim DsBuffer4 As DirectSoundBuffer
Dim DsBuffer5 As DirectSoundBuffer
Dim DsBuffer6 As DirectSoundBuffer
Dim DsBuffer7 As DirectSoundBuffer
Dim DsBuffer8 As DirectSoundBuffer
Dim DsBuffer9 As DirectSoundBuffer
Dim DsBuffer10 As DirectSoundBuffer
Dim DsBuffer11 As DirectSoundBuffer
Dim DsBuffer12 As DirectSoundBuffer
Dim DsBuffer13 As DirectSoundBuffer
Dim DsBuffer14 As DirectSoundBuffer
Dim DsBuffer15 As DirectSoundBuffer
Dim DsBuffer16 As DirectSoundBuffer
Dim DsBuffer17 As DirectSoundBuffer
Dim DsBuffer18 As DirectSoundBuffer
Dim DsDesc As DSBUFFERDESC
Dim DsWave As WAVEFORMATEX
Dim intShipMiddleX As Single
Dim intShipMiddleY As Single
Dim intPlotShipTopX As Single
Dim intPlotShipTopY As Single
Dim intPlotShipLeftX As Single
Dim intPlotShipLeftY As Single
Dim intPlotShipRightX As Single
Dim intPlotShipRightY As Single
Dim intPlotShipCenterX As Single
Dim intPlotShipCenterY As Single
Dim intPlotCoreLeftX As Single
Dim intPlotCoreLeftY As Single
Dim intPlotCoreRightX As Single
Dim intPlotCoreRightY As Single
Dim intPlotCoreCenterX As Single
Dim intPlotCoreCenterY As Single
Dim intPlotFlameCenterX As Single
Dim intPlotFlameCenterY As Single
Dim intFlameSize As Single
Dim flameshown As Integer
Dim leftpressed As Boolean
Dim rightpressed As Boolean
Dim thrustpressed As Boolean
Dim beampressed As Boolean
Dim gravity As Single
Dim velocity As Single
Dim velocityx As Single
Dim singSpeed As Single
Dim xres As Integer
Dim yres As Integer
Dim colourDepth As Integer
Dim intRotationRate As Integer
Dim blnCollided As Boolean
Dim intLandscapeY(1600) As Integer
Dim intLandscapeYCollision(1700) As Single
Dim intLandscapeCount As Single
Dim intLandscapeCount2 As Single
Dim intlandscapemarker As Integer
Dim intLandscapeXIncrement As Single
Dim intLandscapePositionX As Single
Dim intLowerLandscape As Integer
Dim intTerrainVariance As Integer
Dim intTerrainComplexity As Integer
Dim intLandscapeHighestPoint As Single
Dim intLandscapeLowestPoint As Single
Dim intStatsXCoordinate As Integer
Dim i As Single
Dim i2 As Double
Dim intFuelX As Single
Dim intFuelY As Single
Dim blnFuel As Boolean
Dim intFuelLevel As Integer
Dim intRefuelLevel As Integer
Dim intFuelWarningLevel As Integer
Dim intDrugsX As Single
Dim intDrugsY As Single
Dim blnDrugs As Boolean
Dim intGeneralCounter As Integer
Dim intDrugType As Integer
Dim blnDrugsIncrement As Boolean
Dim intRandomNumber As Single
Dim blnAntiGravity As Boolean
Dim blnAntiGravityShown As Boolean
Dim intAntiGravityX As Single
Dim intAntiGravityY As Single
Dim blnAntiGravityIncrement As Boolean
Dim blnScientist1 As Boolean
Dim blnScientist1Increment As Boolean
Dim intScientist1X As Single
Dim intScientist1Y As Single
Dim blnScientist2 As Boolean
Dim blnScientist2Increment As Boolean
Dim intScientist2X As Single
Dim intScientist2Y As Single
Dim blnScientist3 As Boolean
Dim blnScientist3Increment As Boolean
Dim intScientist3X As Single
Dim intScientist3Y As Single
Dim blnScientist4 As Boolean
Dim blnScientist4Increment As Boolean
Dim intScientist4X As Single
Dim intScientist4Y As Single
Dim blnScientist5 As Boolean
Dim blnScientist5Increment As Boolean
Dim intScientist5X As Single
Dim intScientist5Y As Single
Dim intScientistsOutCount As Integer
Dim intBeamShown As Integer
Dim intBeamBottLeftX As Single
Dim intBeamBottLeftY As Single
Dim intBeamBottRightX As Single
Dim intBeamBottRightY As Single
Dim intBeamSize As Single
Dim blnAnyScientistOnBoard As Boolean
Dim blnScientist1OnBoard As Boolean
Dim blnScientist2OnBoard As Boolean
Dim blnScientist3OnBoard As Boolean
Dim blnScientist4OnBoard As Boolean
Dim blnScientist5OnBoard As Boolean
Dim blnShipLanded As Boolean
Dim intTakeOffDelay As Single
Dim blnTakeOffSoundPlayed As Boolean
Dim blnScientist1Free As Boolean
Dim blnScientist2Free As Boolean
Dim blnScientist3Free As Boolean
Dim blnScientist4Free As Boolean
Dim blnScientist5Free As Boolean
Dim blnScientist1Complete As Boolean
Dim blnScientist2Complete As Boolean
Dim blnScientist3Complete As Boolean
Dim blnScientist4Complete As Boolean
Dim blnScientist5Complete As Boolean
Dim intShootSourceX As Single
Dim intShootSourceY As Single
Dim shootpressed As Boolean
Dim blnShotFired As Boolean
Dim intShotRange As Single
Dim intShotX1 As Single
Dim intShotY1 As Single
Dim intShotX2 As Single
Dim intShotY2 As Single
Dim ShotAngle As Single
Dim intSwitchX As Single
Dim intSwitchY As Single
Dim blnSwitchPressed As Boolean
Dim intSwitchButtonTopX As Single
Dim intSwitchButtonTopY As Single
Dim intSwitchButtonBottomX As Single
Dim intSwitchButtonBottomY As Single
Dim intPadLeftTopLeftx As Single
Dim intPadLeftTopLefty As Single
Dim intPadLeftBottomRightx As Single
Dim intPadLeftBottomRighty As Single
Dim intPadRightTopLeftx As Single
Dim intPadRightTopLefty As Single
Dim intPadRightBottomRightx As Single
Dim intPadRightBottomRighty As Single
Dim intGateTopLeftx As Single
Dim intGateTopLefty As Single
Dim intGateBottomRightx As Single
Dim intGateBottomRighty As Single
Dim blnSlidingDoorPlayed As Boolean
Dim intScore As Single
Dim intAddScore As Single
Dim fontinfo As New StdFont   'used in game
Dim fontinfo2 As New StdFont  'used on menu screen
Dim blnMystery As Boolean
Dim blnMysteryIncrement As Boolean
Dim intMysteryX As Single
Dim intMysteryY As Single
Dim blnDanceMode As Boolean
Dim intLaserSourceX As Single
Dim intLaserSourceY As Single
Dim intLaserDest1X As Single
Dim intLaserDest1Y As Single
Dim intLaserStep As Integer
Dim intLaserStep2 As Integer
Dim intColourCycle As Single
Dim intColourCycle2 As Single
Dim intColourCycle3 As Single
Dim blncentermessage As Boolean
Dim intcentermessagetime As Integer
Dim charCenterMessage As Variant
Dim intDoubleGravityCount As Integer
Dim intShipDistMidBott As Single
Dim intShipDistMidSide As Single
Dim intShipDistMidCent As Single
Dim intShipDistMidPill As Single
Dim intPlotShipDistMidCent As Single
Dim intPlotShipDistMidOuterRing As Single
Dim intPlotShipDistMidCore As Single
Dim intBeamSizeMax As Single
Dim intBeamSizeMin As Single
Dim intFlameSizeMin As Single
Dim intFlameSizeMax As Single
Dim intDrugsWavelength As Single
Dim intMysteryWavelength As Single
Dim intAntiGravityWavelength1 As Single
Dim intAntiGravityWavelength2 As Single
Dim intScientistHeight As Single
Dim intScientistHeightFromGround As Single
Dim blnDoubleGravity As Boolean
Dim statspressed As Boolean
Dim blnLegendPressed As Boolean
Dim blnTinyShip As Boolean
Dim intTinyShipCount As Integer
Dim blnInGame As Boolean
Dim blnInMenu As Boolean
Dim intLives As Integer
Dim blnExit As Boolean
Dim counter As Single
Dim backdrop As Variant
Dim sngMenuLineY As Single
Dim sngMenuLineYDest As Single
Dim blnCreateNewLineDest As Boolean
Dim sngMenuLineY2 As Single
Dim sngMenuLineYDest2 As Single
Dim blnCreateNewLineDest2 As Boolean
Dim sngMenuLineY3 As Single
Dim sngMenuLineYDest3 As Single
Dim blnCreateNewLineDest3 As Boolean
Dim sngMenuLineY4 As Single
Dim sngMenuLineYDest4 As Single
Dim blnCreateNewLineDest4 As Boolean
Dim sngMenuLineY5 As Single
Dim sngMenuLineYDest5 As Single
Dim blnCreateNewLineDest5 As Boolean
Dim intLevel As Integer
Dim blnLevelComplete As Boolean
Dim blnShowLevelScreen As Boolean
Dim intLevelGapCounter As Integer
Dim lngCurrentTime As Long
Dim lngDesiredTime As Long
Dim intCircleWidth As Single
   
Dim intTurretX As Single
Dim intTurretY As Single
Dim intTurretRatio As Single
Dim intTurretShipDistance As Single
Dim intTurretXDistance As Single
Dim intTurretYDistance As Single
Dim intTurretCircleX As Single
Dim intTurretCircleY As Single
Dim blnTurret As Boolean
Dim intTurretLength As Single
Dim blnTurretDestroyed As Boolean
Dim intTurretShotLength As Single
Dim intTurretShotXDistance As Single
Dim intTurretShotYDistance As Single
Dim intTurretShotDistance As Single
Dim intTurretShotRatio As Single
Dim blnTurretShot As Boolean
Dim intTurretShotX As Single
Dim intTurretShotY As Single
   
Dim intTurret2X As Single
Dim intTurret2Y As Single
Dim intTurret2Ratio As Single
Dim intTurret2ShipDistance As Single
Dim intTurret2XDistance As Single
Dim intTurret2YDistance As Single
Dim intTurret2CircleX As Single
Dim intTurret2CircleY As Single
Dim blnTurret2 As Boolean
Dim intTurret2Length As Single
Dim blnTurret2Destroyed As Boolean
Dim intTurret2ShotLength As Single
Dim intTurret2ShotXDistance As Single
Dim intTurret2ShotYDistance As Single
Dim intTurret2ShotDistance As Single
Dim intTurret2ShotRatio As Single
Dim blnTurret2Shot As Boolean
Dim intTurret2ShotX As Single
Dim intTurret2ShotY As Single
   
Dim intTurret3X As Single
Dim intTurret3Y As Single
Dim intTurret3Ratio As Single
Dim intTurret3ShipDistance As Single
Dim intTurret3XDistance As Single
Dim intTurret3YDistance As Single
Dim intTurret3CircleX As Single
Dim intTurret3CircleY As Single
Dim blnTurret3 As Boolean
Dim intTurret3Length As Single
Dim blnTurret3Destroyed As Boolean
Dim intTurret3ShotLength As Single
Dim intTurret3ShotXDistance As Single
Dim intTurret3ShotYDistance As Single
Dim intTurret3ShotDistance As Single
Dim intTurret3ShotRatio As Single
Dim blnTurret3Shot As Boolean
Dim intTurret3ShotX As Single
Dim intTurret3ShotY As Single
   
Dim intTurret4X As Single
Dim intTurret4Y As Single
Dim intTurret4Ratio As Single
Dim intTurret4ShipDistance As Single
Dim intTurret4XDistance As Single
Dim intTurret4YDistance As Single
Dim intTurret4CircleX As Single
Dim intTurret4CircleY As Single
Dim blnTurret4 As Boolean
Dim intTurret4Length As Single
Dim blnTurret4Destroyed As Boolean
Dim intTurret4ShotLength As Single
Dim intTurret4ShotXDistance As Single
Dim intTurret4ShotYDistance As Single
Dim intTurret4ShotDistance As Single
Dim intTurret4ShotRatio As Single
Dim blnTurret4Shot As Boolean
Dim intTurret4ShotX As Single
Dim intTurret4ShotY As Single
   
Dim intTurret5X As Single
Dim intTurret5Y As Single
Dim intTurret5Ratio As Single
Dim intTurret5ShipDistance As Single
Dim intTurret5XDistance As Single
Dim intTurret5YDistance As Single
Dim intTurret5CircleX As Single
Dim intTurret5CircleY As Single
Dim blnTurret5 As Boolean
Dim intTurret5Length As Single
Dim blnTurret5Destroyed As Boolean
Dim intTurret5ShotLength As Single
Dim intTurret5ShotXDistance As Single
Dim intTurret5ShotYDistance As Single
Dim intTurret5ShotDistance As Single
Dim intTurret5ShotRatio As Single
Dim blnTurret5Shot As Boolean
Dim intTurret5ShotX As Single
Dim intTurret5ShotY As Single
   
Dim intTurret6X As Single
Dim intTurret6Y As Single
Dim intTurret6Ratio As Single
Dim intTurret6ShipDistance As Single
Dim intTurret6XDistance As Single
Dim intTurret6YDistance As Single
Dim intTurret6CircleX As Single
Dim intTurret6CircleY As Single
Dim blnTurret6 As Boolean
Dim intTurret6Length As Single
Dim blnTurret6Destroyed As Boolean
Dim intTurret6ShotLength As Single
Dim intTurret6ShotXDistance As Single
Dim intTurret6ShotYDistance As Single
Dim intTurret6ShotDistance As Single
Dim intTurret6ShotRatio As Single
Dim blnTurret6Shot As Boolean
Dim intTurret6ShotX As Single
Dim intTurret6ShotY As Single
   
Dim intTurret7X As Single
Dim intTurret7Y As Single
Dim intTurret7Ratio As Single
Dim intTurret7ShipDistance As Single
Dim intTurret7XDistance As Single
Dim intTurret7YDistance As Single
Dim intTurret7CircleX As Single
Dim intTurret7CircleY As Single
Dim blnTurret7 As Boolean
Dim intTurret7Length As Single
Dim blnTurret7Destroyed As Boolean
Dim intTurret7ShotLength As Single
Dim intTurret7ShotXDistance As Single
Dim intTurret7ShotYDistance As Single
Dim intTurret7ShotDistance As Single
Dim intTurret7ShotRatio As Single
Dim blnTurret7Shot As Boolean
Dim intTurret7ShotX As Single
Dim intTurret7ShotY As Single
   
Sub Init()
ShowCursor (0)                                                                                  'Hide the cursor
On Local Error GoTo errOut
Set dd = dx.DirectDrawCreate("")                                                                'the default driver
Me.Show
Call dd.SetCooperativeLevel(Me.hWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE)   'links the DirectDraw object to the form
xres = Form2.txtXRes.Text                                                                       'get screen width entered by player on settings screen
yres = Form2.txtYRes.Text                                                                       'get screen height entered by player on settings screen
colourDepth = Form2.txtColourDepth.Text                                                         'get colour depth entered by player on settings screen
Call dd.SetDisplayMode(xres, yres, colourDepth, 0, DDSDM_DEFAULT)                               'set the screen as the player selected
ddsd1.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT                                                'get the screen surface and create back buffer
ddsd1.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
ddsd1.lBackBufferCount = 1
Set primary = dd.CreateSurface(ddsd1)
Set Ds = dx.DirectSoundCreate("")                                                               'create directsound
Ds.SetCooperativeLevel Form1.hWnd, DSSCL_NORMAL                                                 'let other programs use directsound at the same time
DsDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
DsWave.nFormatTag = WAVE_FORMAT_PCM
DsWave.nChannels = 2                                                                            '1= Mono, 2 = Stereo
DsWave.lSamplesPerSec = 22050
DsWave.nBitsPerSample = 16
DsWave.nBlockAlign = DsWave.nBitsPerSample / 8 * DsWave.nChannels
DsWave.lAvgBytesPerSec = DsWave.lSamplesPerSec * DsWave.nBlockAlign
Set DsBuffer = Ds.CreateSoundBufferFromFile(App.Path & "\thrust.Wav", DsDesc, DsWave)
Set DsBuffer2 = Ds.CreateSoundBufferFromFile(App.Path & "\crashed.Wav", DsDesc, DsWave)
Set DsBuffer3 = Ds.CreateSoundBufferFromFile(App.Path & "\landed2.Wav", DsDesc, DsWave)
Set DsBuffer4 = Ds.CreateSoundBufferFromFile(App.Path & "\takeoff.Wav", DsDesc, DsWave)
Set DsBuffer5 = Ds.CreateSoundBufferFromFile(App.Path & "\warning2.Wav", DsDesc, DsWave)
Set DsBuffer6 = Ds.CreateSoundBufferFromFile(App.Path & "\collected.Wav", DsDesc, DsWave)
Set DsBuffer7 = Ds.CreateSoundBufferFromFile(App.Path & "\itemappears.Wav", DsDesc, DsWave)
Set DsBuffer8 = Ds.CreateSoundBufferFromFile(App.Path & "\newscientist.Wav", DsDesc, DsWave)
Set DsBuffer9 = Ds.CreateSoundBufferFromFile(App.Path & "\tractorbeam.Wav", DsDesc, DsWave)
Set DsBuffer10 = Ds.CreateSoundBufferFromFile(App.Path & "\gotscientist.Wav", DsDesc, DsWave)
Set DsBuffer11 = Ds.CreateSoundBufferFromFile(App.Path & "\shoot.Wav", DsDesc, DsWave)
Set DsBuffer12 = Ds.CreateSoundBufferFromFile(App.Path & "\gateopen.Wav", DsDesc, DsWave)
Set DsBuffer13 = Ds.CreateSoundBufferFromFile(App.Path & "\denied.Wav", DsDesc, DsWave)
Set DsBuffer14 = Ds.CreateSoundBufferFromFile(App.Path & "\vault.Wav", DsDesc, DsWave)
Set DsBuffer15 = Ds.CreateSoundBufferFromFile(App.Path & "\slidingdoor.Wav", DsDesc, DsWave)
Set DsBuffer16 = Ds.CreateSoundBufferFromFile(App.Path & "\music1.Wav", DsDesc, DsWave)
Set DsBuffer17 = Ds.CreateSoundBufferFromFile(App.Path & "\photon.Wav", DsDesc, DsWave)
Set DsBuffer18 = Ds.CreateSoundBufferFromFile(App.Path & "\explos.Wav", DsDesc, DsWave)
Dim caps As DDSCAPS2                                                                            'Get the backbuffer
caps.lCaps = DDSCAPS_BACKBUFFER
Set backbuffer = primary.GetAttachedSurface(caps)
backbuffer.GetSurfaceDesc ddsd3
blnInGame = False
blnInMenu = True
InitialiseMenuLines
InitSurfaces                                                                                    'init the surfaces
backbuffer.SetFontTransparency True                                                             'make sure text (score, etc) doesnt obscure the game
If Form2.chkStats = 1 Then                                                                      'did player choose to view statistics by default on settings screen?
    statspressed = True
Else
    statspressed = False
End If
If Form2.chkLegend = 1 Then                                                                      'did player choose to view legend by default on settings screen?
    blnLegendPressed = True
Else
    blnLegendPressed = False
End If

binit = True
blnExit = False
Do Until blnExit = True                                                                         'main program loop
    
    If blnInMenu = True Then                                                                    'if we should be in the menu, display the menu
        blt
    End If
    
    If blnShowLevelScreen = True Then                                                           'if player has just completed a level, show the level number screen
        blt
    End If
        
    If blnInGame = True Then                                                                    'if player is starting a new game, draw the game
        brunning = True                                                                         'brunning is what determines when its time to exit the game
        SetUpNewGame                                                                            'initial settings that will apply ONCE ONLY
        SetUpNewLevel                                                                           'initial settings that will apply on a NEW LEVEL
        SetUpNewLife                                                                            'initial settnigs that will apply on a NEW LIFE
        Do Until brunning = False                                                               'This is the main loop. It only runs whilst brunning=true
            If blnLevelComplete = True Then                                                     'player has just completed a level, so...
                blnShipLanded = False                                                           'make sure ship is ready to go for a new level
                SetUpNewLevel                                                                   'create the new level
                SetUpNewLife                                                                    'set up new life - this doesnt award a life...it just resets ship, drugs, effects, etc
                intLevelGapCounter = 0                                                          'counter used to time the delay between levels
                blnShowLevelScreen = True                                                       'get ready to show the level number screen
                blnInGame = False                                                               'make sure the game is not displayed during the level number screen
                blnLevelComplete = False                                                        'a level has just been completed, so reset the level completed flag
            End If
            CollisionDetection                                                                  'Check whether ship has hit landscape, roof, landing pad walls
            CheckForLandingOrCrash                                                              'Check whether ship is landing, or is crash-landing
            CheckForPowerUpCollect                                                              'Check whether ship is collecting any powerups
            If blnShipLanded = False Then                                                       'If the ship is airborne....
                CheckForRotation                                                                'Check if player wants ship to rotate
                CheckForScientistCollect                                                        'Check if player is collecting a scientist
                CheckForThrust                                                                  'Check if the player is applying thrust
            End If
            CheckForGateOpenClose                                                               'See if gates coordinates need changing
            SetShipCoordinates                                                                  'Set the new coordinates needed to draw ship and flame
            If blnDoubleGravity = True Then                                                     'If double gravity is on...
                DoubleGravity                                                                   'Double it (or normalise again if time is up
            End If
            If blnTinyShip = True Then                                                          'If tiny ship mode is on...
                TinyShip                                                                        'Shrink ship (or normalise again if time is up)
            End If
            blt                                                                                 'Draw the game
                
            DoEvents                                                                            'avoid an overflow of messages being sent to DirectX
        Loop                                                                                    'end of the game loop
    End If
Loop                                                                                            'end of the program loop
errOut:                                                                                         'If there is an error we want to close the program down straight away.
EndIt                                                                                           'terminate program
End Sub

Private Sub form_keypress(key As Integer)
    If blnInMenu = True Then                                                                    'check for key presses while menu is being shown
        Select Case key
            Case vbKeyEscape                                                                    'escape - Quit Game
                blnExit = True
                EndIt
            Case 49                                                                             '1 Start Game
                blnInMenu = False
                blnInGame = False
                blnShowLevelScreen = True                                                       'get ready to show the level intro screen
        End Select
    End If
            
    If blnInGame = True Then                                                                    'check for key presses while in game
        If blnShipLanded = False Then                                                           'only bother if ship is airborne
            Select Case key
                Case 90 Or 122                                                                  'z - rotate left
                    leftpressed = True
                Case 88 Or 120                                                                  'x - rotate right
                    rightpressed = True
                Case 77 Or 109                                                                  'm - tractor beam
                    beampressed = True
                Case 78 Or 110                                                                  'n - shoot
                    shootpressed = True
                Case 32                                                                         'space - thrust
                    thrustpressed = True
                Case 83 Or 115                                                                  's - toggle stats
                    If statspressed = True Then
                        statspressed = False
                    Else
                        statspressed = True
                    End If
                Case 75 Or 107                                                                  'k - toggle legend
                    If blnLegendPressed = True Then
                        blnLegendPressed = False
                    Else
                        blnLegendPressed = True
                    End If
                Case 81 Or 113                                                                  'q - quit
                    counter = 256
                    blnInGame = False
                    blnShowLevelScreen = False
                    blnInMenu = True                                                            'quit back to menu
                    InitialiseMenuLines
                    StopAccessDeniedSound
                    StopCollectedSound
                    StopCrashedSound
                    StopDance1Sound
                    StopGateOpenSound
                    StopGotScientistSound
                    StopItemAppearsSound
                    StopLandedSound
                    StopNewScientistSound
                    StopPhotonSound
                    StopShootSound
                    StopSlidingDoorSound
                    StopTakeoffSound
                    StopThrustSound
                    StopTractorBeamSound
                    StopTurretDestroyedSound
                    StopVaultSound
                    StopWarningSound
                    brunning = False
            End Select
        End If
    End If
End Sub

Private Sub form_keyup(key As Integer, shift As Integer)                                        'check for releasing of keys in game
    If blnShipLanded = False Then                                                               'only bother if ship is airborne
        Select Case key
            Case 90                                                                             'z - stop rotating left
                leftpressed = False
            Case 88                                                                             'x - stop rotating right
                rightpressed = False
            Case 77                                                                             'm - stop using tractor beam
                beampressed = False
                intBeamSize = intBeamSizeMin                                                    'reset the size of the beam, ready for the next time its applied
                StopTractorBeamSound
            Case 32                                                                             'space - stop applying thrust
                thrustpressed = False
                intFlameSize = intFlameSizeMin                                                  'reset the size of the thrusters flame, ready for the next time its applied
        End Select
    End If
End Sub

Sub InitSurfaces()

Set Mainsurf = Nothing
                                                                                                'load the bitmap into a surface - backdrop3.bmp
ddsd2.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH                                           'default flags
ddsd2.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ddsd2.lWidth = ddsd3.lWidth                                                                     'the ddsd3 structure already holds the size of the screen.
ddsd2.lHeight = ddsd3.lHeight
Set Mainsurf = dd.CreateSurfaceFromFile(App.Path & "\backdrop2.bmp", ddsd2)

Set Menusurf = Nothing
                                                                                                'load the bitmap into a surface - backdrop.bmp
ddsd4.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH                                           'default flags
ddsd4.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ddsd4.lWidth = ddsd3.lWidth                                                                     'the ddsd3 structure already holds the size of the screen.
ddsd4.lHeight = ddsd3.lHeight
Set Menusurf = dd.CreateSurfaceFromFile(App.Path & "\backdrop.bmp", ddsd4)

Set Levelsurf = Nothing
                                                                                                'load the bitmap into a surface - backdrop.bmp
ddsd5.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH                                           'default flags
ddsd5.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
ddsd5.lWidth = ddsd3.lWidth                                                                     'the ddsd3 structure already holds the size of the screen.
ddsd5.lHeight = ddsd3.lHeight
Set Levelsurf = dd.CreateSurfaceFromFile(App.Path & "\backdrop4.bmp", ddsd4)

End Sub
Sub blt()

On Local Error GoTo errOut                                                                      'If there is an error don't do anything - just skip the procedure
If binit = False Then Exit Sub                                                                  'If we haven't initiaised then don't try anything DirectDraw related.

Dim ddrval As Long
Dim rBack As RECT
Dim i As Integer
                                                                                                ' this will keep us from trying to blt in case we lose the surfaces (alt-tab)
bRestore = False
Do Until ExModeActive
    DoEvents
    bRestore = True
Loop
                                                                                                ' if we lost and got back the surfaces, then restore them
DoEvents
If bRestore Then
    bRestore = False
    dd.RestoreAllSurfaces                                                                       're-allocates memory back to program
    InitSurfaces                                                                                'must init the surfaces again if they we're lost
End If

rBack.Bottom = ddsd3.lHeight
rBack.Right = ddsd3.lWidth

If blnInGame = True Then
    ddrval = backbuffer.BltFast(0, 0, Mainsurf, rBack, DDBLTFAST_WAIT)
    DrawGame
Else
    If blnShowLevelScreen = True Then
        ddrval = backbuffer.BltFast(0, 0, Levelsurf, rBack, DDBLTFAST_WAIT)
        DrawLevelScreen
    Else
        ddrval = backbuffer.BltFast(0, 0, Menusurf, rBack, DDBLTFAST_WAIT)
        DrawMenu
    End If
End If

backbuffer.SetForeColor RGB(150, 150, 150)                                                      'set border colour that will go round edge of screen
backbuffer.DrawLine 0, 0, xres, 0                                                               'draw border - top line
backbuffer.DrawLine 0, 0, 0, yres - 1                                                           'draw border - left line
backbuffer.DrawLine 0, yres - 1, xres, yres - 1                                                 'draw border - bottom line
backbuffer.DrawLine xres - 1, yres - 1, xres - 1, 0                                             'draw border - right line

'flip the back buffer to the screen
primary.Flip Nothing, DDFLIP_WAIT

errOut:
End Sub

Sub EndIt()
ShowCursor (1)
Call dd.RestoreDisplayMode                                                                      'restores you back to your default (windows) resolution.
Call dd.SetCooperativeLevel(Me.hWnd, DDSCL_NORMAL)                                              'This tells windows/directX that we no longer want exclusive access to the graphics features/directdraw
End                                                                                             'Stop the program
End Sub

Private Sub Form_Load()
Init                                                                                            'Starts the whole program.
End Sub

Private Sub Form_Paint()
blt                                                                                             'If windows sends a "paint" message translate this into a call to DirectDraw.
End Sub

Function ExModeActive() As Boolean
                                                                                                'This is used to test if we're in the correct resolution.
Dim TestCoopRes As Long
TestCoopRes = dd.TestCooperativeLevel
If (TestCoopRes = DD_OK) Then
    ExModeActive = True
Else
    ExModeActive = False
End If
End Function

Sub LifeLost()                                                                                  'player has just lost a life

StopThrustSound                                                                                 'in case thrust was being applied at impact time, stop the thrust sound
PlayCrashedSound                                                                                'play crash (explosion) sound
counter = 0
velocity = 0
velocityx = 0
leftpressed = True                                                                              'this will make the ship spin as if out of control
If intShipMiddleX > 1615 Then                                                                   'make sure the ship stays on screen
    intShipMiddleX = -15
End If
If intShipMiddleX < -15 Then
    intShipMiddleX = 1615
End If
If intShipMiddleY < -15 Then
    intShipMiddleY = 1200
End If
If intShipMiddleY > 1215 Then
    intShipMiddleY = -15
End If
Do While counter < 255                                                                          'start the crash sequence
    counter = counter + 1
    blnTurretShot = False                                                                       'make sure turrets dont keep firing
    blnTurret2Shot = False
    blnTurret3Shot = False
    stepval = stepval + 15                                                                      'increase the angle of rotation by 15 radians to make ship spin
    backbuffer.SetForeColor RGB(255 - counter, 255 - counter, 255 - counter)                    'increasingly darken the ships colour, making it fade
    intPlotShipCenterX = intShipMiddleX + Sin(stepval * Rad) * (intPlotShipDistMidCent + counter) 'set the ships coordinates
    intPlotShipCenterY = intShipMiddleY + Cos(stepval * Rad) * (intPlotShipDistMidCent + counter)
    intPlotShipLeftX = intShipMiddleX + Sin((stepval - 40) * Rad) * (intPlotShipDistMidOuterRing + counter)
    intPlotShipLeftY = intShipMiddleY + Cos((stepval - 40) * Rad) * (intPlotShipDistMidOuterRing + counter)
    intPlotShipRightX = intShipMiddleX + Sin((stepval + 40) * Rad) * (intPlotShipDistMidOuterRing + counter)
    intPlotShipRightY = intShipMiddleY + Cos((stepval + 40) * Rad) * (intPlotShipDistMidOuterRing + counter)
    intPlotShipTopX = intShipMiddleX + Sin((stepval - 180) * Rad) * (intPlotShipDistMidOuterRing + counter)
    intPlotShipTopY = intShipMiddleY + Cos((stepval - 180) * Rad) * (intPlotShipDistMidOuterRing + counter)
    intPlotCoreLeftX = intShipMiddleX + Sin((stepval - 15) * Rad) * (intPlotShipDistMidOuterRing + counter)
    intPlotCoreLeftY = intShipMiddleY + Cos((stepval - 15) * Rad) * (intPlotShipDistMidOuterRing + counter)
    intPlotCoreRightX = intShipMiddleX + Sin((stepval + 15) * Rad) * (intPlotShipDistMidOuterRing + counter)
    intPlotCoreRightY = intShipMiddleY + Cos((stepval + 15) * Rad) * (intPlotShipDistMidOuterRing + counter)
    blt                                                                                         'draw the screen
    DoEvents
Loop
StopCrashedSound                                                                                'if the crash sound hasn't completed yet, stop it
intLives = intLives - 1                                                                         'subtract a life

If intLives <= 0 Then                                                                           'if no lives left...
    GameOver                                                                                    'game over!
Else                                                                                            'otherwise....
    backbuffer.SetForeColor RGB(255, 255, 255)                                                  'reset ships colour
    SetUpNewLife                                                                                'create a new life
End If
End Sub

Sub PlayThrustSound()
DsBuffer.Play DSBPLAY_DEFAULT                                                                   'Set to DSBPLAY_DEFAULT so it only plays once.
End Sub

Sub StopThrustSound()
DsBuffer.Stop                                                                                   'Stop the sound (acts like a pause)
DsBuffer.SetCurrentPosition 0                                                                   'reset the position of the sound for future use
End Sub

Sub PlayCrashedSound()
DsBuffer2.Play DSBPLAY_DEFAULT                                                                  'Set to DSBPLAY_DEFAULT so it only plays once.
End Sub

Sub StopCrashedSound()
DsBuffer2.Stop                                                                                  'Stop the sound (acts like a pause)
DsBuffer2.SetCurrentPosition 0                                                                  'reset the position of the sound for future use
End Sub

Sub PlayLandedSound()
DsBuffer3.Play DSBPLAY_DEFAULT                                                                  'Set to DSBPLAY_DEFAULT so it only plays once.
End Sub

Sub StopLandedSound()
DsBuffer3.Stop                                                                                  'Stop the sound (acts like a pause)
DsBuffer3.SetCurrentPosition 0                                                                  'reset the position of the sound for future use
End Sub

Sub PlayTakeoffSound()
DsBuffer4.Play DSBPLAY_DEFAULT                                                                  'Set to DSBPLAY_DEFAULT so it only plays once.
End Sub

Sub StopTakeoffSound()
DsBuffer4.Stop                                                                                  'Stop the sound (acts like a pause)
DsBuffer4.SetCurrentPosition 0                                                                  'reset the position of the sound for future use
End Sub

Sub PlayWarningSound()
DsBuffer5.Play DSBPLAY_DEFAULT                                                                  'Set to DSBPLAY_DEFAULT so it only plays once.
End Sub

Sub StopWarningSound()
DsBuffer5.Stop                                                                                  'Stop the sound (acts like a pause)
DsBuffer5.SetCurrentPosition 0                                                                  'reset the position of the sound for future use
End Sub
Sub PlayCollectedSound()
DsBuffer6.Play DSBPLAY_DEFAULT                                                                  'Set to DSBPLAY_DEFAULT so it only plays once.
End Sub

Sub StopCollectedSound()
DsBuffer6.Stop                                                                                  'Stop the sound (acts like a pause)
DsBuffer6.SetCurrentPosition 0                                                                  'reset the position of the sound for future use
End Sub
Sub PlayItemAppearsSound()
DsBuffer7.Play DSBPLAY_DEFAULT                                                                  'Set to DSBPLAY_DEFAULT so it only plays once.
End Sub

Sub StopItemAppearsSound()
DsBuffer7.Stop                                                                                  'Stop the sound (acts like a pause)
DsBuffer7.SetCurrentPosition 0                                                                  'reset the position of the sound for future use
End Sub
Sub PlayNewScientistSound()
DsBuffer8.Play DSBPLAY_DEFAULT                                                                  'Set to DSBPLAY_DEFAULT so it only plays once.
End Sub

Sub StopNewScientistSound()
DsBuffer8.Stop                                                                                  'Stop the sound (acts like a pause)
DsBuffer8.SetCurrentPosition 0                                                                  'reset the position of the sound for future use
End Sub

Sub PlayTractorBeamSound()
DsBuffer9.Play DSBPLAY_LOOPING                                                                  'Set to DSBPLAY_LOOPING so it loops.
End Sub

Sub StopTractorBeamSound()
DsBuffer9.Stop                                                                                  'Stop the sound (acts like a pause)
DsBuffer9.SetCurrentPosition 0                                                                  'reset the position of the sound for future use
End Sub
Sub PlayGotScientistSound()
DsBuffer10.Play DSBPLAY_DEFAULT                                                                 'Set to DSBPLAY_DEFAULT so it only plays once.
End Sub

Sub StopGotScientistSound()
DsBuffer10.Stop                                                                                 'Stop the sound (acts like a pause)
DsBuffer10.SetCurrentPosition 0                                                                 'reset the position of the sound for future use
End Sub
Sub PlayShootSound()
DsBuffer11.Play DSBPLAY_DEFAULT                                                                 'Set to DSBPLAY_DEFAULT so it only plays once.
End Sub

Sub StopShootSound()
DsBuffer11.Stop                                                                                 'Stop the sound (acts like a pause)
DsBuffer11.SetCurrentPosition 0                                                                 'reset the position of the sound for future use
End Sub
Sub PlayGateOpenSound()
DsBuffer12.Play DSBPLAY_DEFAULT                                                                 'Set to DSBPLAY_DEFAULT so it only plays once.
End Sub

Sub StopGateOpenSound()
DsBuffer12.Stop                                                                                 'Stop the sound (acts like a pause)
DsBuffer12.SetCurrentPosition 0                                                                 'reset the position of the sound for future use
End Sub
Sub PlayAccessDeniedSound()
DsBuffer13.Play DSBPLAY_DEFAULT                                                                 'Set to DSBPLAY_DEFAULT so it only plays once.
End Sub

Sub StopAccessDeniedSound()
DsBuffer13.Stop                                                                                 'Stop the sound (acts like a pause)
DsBuffer13.SetCurrentPosition 0                                                                 'reset the position of the sound for future use
End Sub
Sub PlayVaultSound()
DsBuffer14.Play DSBPLAY_DEFAULT                                                                 'Set to DSBPLAY_DEFAULT so it only plays once.
End Sub

Sub StopVaultSound()
DsBuffer14.Stop                                                                                 'Stop the sound (acts like a pause)
DsBuffer14.SetCurrentPosition 0                                                                 'reset the position of the sound for future use
End Sub
Sub PlaySlidingDoorSound()
DsBuffer15.Play DSBPLAY_LOOPING                                                                 'Set to DSBPLAY_LOOPING so it loops.
End Sub

Sub StopSlidingDoorSound()
DsBuffer15.Stop                                                                                 'Stop the sound (acts like a pause)
DsBuffer15.SetCurrentPosition 0                                                                 'reset the position of the sound for future use
End Sub
Sub PlayDance1Sound()
DsBuffer16.Play DSBPLAY_DEFAULT                                                                 'Set to DSBPLAY_DEFAULT so it only plays once.
End Sub

Sub StopDance1Sound()
DsBuffer16.Stop                                                                                 'Stop the sound (acts like a pause)
DsBuffer16.SetCurrentPosition 0                                                                 'reset the position of the sound for future use
End Sub
Sub PlayPhotonSound()
DsBuffer17.Play DSBPLAY_DEFAULT                                                                 'Set to DSBPLAY_DEFAULT so it only plays once.
End Sub

Sub StopPhotonSound()
DsBuffer17.Stop                                                                                 'Stop the sound (acts like a pause)
DsBuffer17.SetCurrentPosition 0                                                                 'reset the position of the sound for future use
End Sub

Sub PlayTurretDestroyedSound()
DsBuffer18.Play DSBPLAY_DEFAULT                                                                 'Set to DSBPLAY_DEFAULT so it only plays once.
End Sub

Sub StopTurretDestroyedSound()
DsBuffer18.Stop                                                                                 'Stop the sound (acts like a pause)
DsBuffer18.SetCurrentPosition 0                                                                 'reset the position of the sound for future use
End Sub

Sub GenerateTurret()                                                                            'work out where a gun turret is to be placed
Randomize Timer
intTurretCircleX = Rnd * xres                                                                   'create a random x position
intTurretCircleY = intLandscapeYCollision(intTurretCircleX) - 8                                 'set the y position to just above the land level
blnTurret = True                                                                                'set flag to say that gun turret is on screen
End Sub

Sub DrawTurret()
    backbuffer.SetFillStyle 0
    backbuffer.SetFillColor RGB(50, 50, 50)
    backbuffer.SetForeColor RGB(150, 150, 150)
                                                                                                'calculate the coordinates of the turret itself
    intTurretLength = yres / 50                                                                 'the length of the turret
    intTurretXDistance = intTurretCircleX - intShipMiddleX                                      'x distance between turretcircle and ship
    intTurretYDistance = intTurretCircleY - intShipMiddleY                                      'y distance between turretcircle and ship
    intTurretShipDistance = Sqr((intTurretYDistance * intTurretYDistance) + (intTurretXDistance * intTurretXDistance)) 'distance of the xy distance
    intTurretRatio = intTurretLength / intTurretShipDistance                                    'how many pixels long is the turret
    intTurretX = intTurretCircleX - (intTurretRatio * intTurretXDistance)                       'plot x end of turret
    intTurretY = intTurretCircleY - (intTurretRatio * intTurretYDistance)                       'plot y end of turret
    backbuffer.DrawLine intTurretCircleX, intTurretCircleY, intTurretX, intTurretY              'draw turret
    backbuffer.DrawCircle intTurretCircleX, intTurretCircleY, intCircleWidth                    'draw turret base
End Sub

Sub generateTurretShot()
    If blnShipLanded = False Then                                                               'dont shoot at our ship if its landed
        blnTurretShot = True                                                                    'make sure that the shot always gets drawn
        PlayPhotonSound                                                                         'sound effect
        intTurretShotLength = yres / 100
        intTurretShotXDistance = intTurretCircleX - intShipMiddleX
        intTurretShotYDistance = intTurretCircleY - intShipMiddleY
        intTurretShotDistance = Sqr((intTurretShotYDistance * intTurretShotYDistance) + (intTurretShotXDistance * intTurretShotXDistance))
    End If
End Sub

Sub drawTurretShot()
    backbuffer.SetFillStyle 0
    backbuffer.SetForeColor RGB(206, 159, 255)
    backbuffer.SetFillColor RGB(206, 159, 255)
    intTurretShotLength = intTurretShotLength + 4
    intTurretShotRatio = intTurretShotLength / intTurretShotDistance
    intTurretShotX = intTurretCircleX - (intTurretShotRatio * intTurretShotXDistance)
    intTurretShotY = intTurretCircleY - (intTurretShotRatio * intTurretShotYDistance)
    backbuffer.DrawCircle intTurretShotX, intTurretShotY, 2
    If intTurretShotX >= xres Then                                                               'has the shot hit the right of the screen?
        blnTurretShot = False
    End If
    If intTurretShotX <= 0 Then                                                                  'has the shot hit the left of the screen?
        blnTurretShot = False
    End If
    If intTurretShotY <= 0 Then                                                                  'has the shot hit the top of the screen?
        blnTurretShot = False
    End If
    If intTurretShotY >= intLandscapeYCollision(intTurretShotX) Then                             'has the shot hit the landscape?
        blnTurretShot = False
    End If
    If intTurretShotX <= (intShipMiddleX + intShipDistMidSide) And intTurretShotX >= (intShipMiddleX - intShipDistMidSide) Then 'has shot hit players ship?
        If intTurretShotY <= (intShipMiddleY + intShipDistMidBott) And intTurretShotY >= (intShipMiddleY - intShipDistMidBott) Then
            blnTurretShot = False
            blnTurret2Shot = False
            blnTurret3Shot = False
            blnTurret4Shot = False
            blnTurret5Shot = False
            blnTurret6Shot = False
            blnTurret7Shot = False
            blnCollided = True
            LifeLost
        End If
    End If
End Sub

Sub GenerateTurret2()                                                                            'work out where a gun turret is to be placed
Randomize Timer
intTurret2CircleX = Rnd * xres                                                                   'create a random x position
intTurret2CircleY = intLandscapeYCollision(intTurret2CircleX) - 8                                 'set the y position to just above the land level
blnTurret2 = True                                                                                'set flag to say that gun turret is on screen
End Sub

Sub DrawTurret2()
    backbuffer.SetFillStyle 0
    backbuffer.SetFillColor RGB(50, 50, 50)
    backbuffer.SetForeColor RGB(150, 150, 150)
                                                                                                'calculate the coordinates of the turret itself
    intTurret2Length = yres / 50                                                                 'the length of the turret
    intTurret2XDistance = intTurret2CircleX - intShipMiddleX                                      'x distance between turretcircle and ship
    intTurret2YDistance = intTurret2CircleY - intShipMiddleY                                      'y distance between turretcircle and ship
    intTurret2ShipDistance = Sqr((intTurret2YDistance * intTurret2YDistance) + (intTurret2XDistance * intTurret2XDistance)) 'distance of the xy distance
    intTurret2Ratio = intTurret2Length / intTurret2ShipDistance                                    'how many pixels long is the turret
    intTurret2X = intTurret2CircleX - (intTurret2Ratio * intTurret2XDistance)                       'plot x end of turret
    intTurret2Y = intTurret2CircleY - (intTurret2Ratio * intTurret2YDistance)                       'plot y end of turret
    backbuffer.DrawLine intTurret2CircleX, intTurret2CircleY, intTurret2X, intTurret2Y              'draw turret
    backbuffer.DrawCircle intTurret2CircleX, intTurret2CircleY, intCircleWidth                      'draw turret base
End Sub

Sub generateTurret2Shot()
    If blnShipLanded = False Then                                                               'dont shoot at our ship if its landed
        blnTurret2Shot = True                                                                    'make sure that the shot always gets drawn
        PlayPhotonSound                                                                         'sound effect
        intTurret2ShotLength = yres / 100
        intTurret2ShotXDistance = intTurret2CircleX - intShipMiddleX
        intTurret2ShotYDistance = intTurret2CircleY - intShipMiddleY
        intTurret2ShotDistance = Sqr((intTurret2ShotYDistance * intTurret2ShotYDistance) + (intTurret2ShotXDistance * intTurret2ShotXDistance))
    End If
End Sub

Sub drawTurret2Shot()
    backbuffer.SetFillStyle 0
    backbuffer.SetForeColor RGB(206, 159, 255)
    backbuffer.SetFillColor RGB(206, 159, 255)
    intTurret2ShotLength = intTurret2ShotLength + 4
    intTurret2ShotRatio = intTurret2ShotLength / intTurret2ShotDistance
    intTurret2ShotX = intTurret2CircleX - (intTurret2ShotRatio * intTurret2ShotXDistance)
    intTurret2ShotY = intTurret2CircleY - (intTurret2ShotRatio * intTurret2ShotYDistance)
    backbuffer.DrawCircle intTurret2ShotX, intTurret2ShotY, 2
    If intTurret2ShotX >= xres Then                                                               'has the shot hit the right of the screen?
        blnTurret2Shot = False
    End If
    If intTurret2ShotX <= 0 Then                                                                  'has the shot hit the left of the screen?
        blnTurret2Shot = False
    End If
    If intTurret2ShotY <= 0 Then                                                                  'has the shot hit the top of the screen?
        blnTurret2Shot = False
    End If
    If intTurret2ShotY >= intLandscapeYCollision(intTurret2ShotX) Then                             'has the shot hit the landscape?
        blnTurret2Shot = False
    End If
    If intTurret2ShotX <= (intShipMiddleX + intShipDistMidSide) And intTurret2ShotX >= (intShipMiddleX - intShipDistMidSide) Then 'has shot hit players ship?
        If intTurret2ShotY <= (intShipMiddleY + intShipDistMidBott) And intTurret2ShotY >= (intShipMiddleY - intShipDistMidBott) Then
            blnTurretShot = False
            blnTurret2Shot = False
            blnTurret3Shot = False
            blnTurret4Shot = False
            blnTurret5Shot = False
            blnTurret6Shot = False
            blnTurret7Shot = False
            blnCollided = True
            LifeLost
        End If
    End If
End Sub

Sub GenerateTurret3()                                                                            'work out where a gun turret3 is to be placed
Randomize Timer
intTurret3CircleX = Rnd * xres                                                                   'create a random x position
If intTurret3CircleX <= intPadRightTopLeftx + 10 Then                                            'make sure its further right than landing pad
    intTurret3CircleX = intTurret3CircleX + intPadRightTopLeftx + 10
End If
If intTurret3CircleX >= intStatsXCoordinate Then
    intTurret3CircleX = xres / 2
End If
intTurret3CircleY = yres / (0.5 * yres)  '(2)                                                    'set the y position so its hanging from the roof
blnTurret3 = True                                                                                'set flag to say that gun turret is on screen
End Sub

Sub DrawTurret3()
    backbuffer.SetFillStyle 0
    backbuffer.SetFillColor RGB(50, 50, 50)
    backbuffer.SetForeColor RGB(150, 150, 150)
                                                                                                 'calculate the coordinates of the turret3 itself
    intTurret3Length = yres / 50                                                                 'the length of the turret3
    intTurret3XDistance = intTurret3CircleX - intShipMiddleX                                     'x distance between turret3circle and ship
    intTurret3YDistance = intTurret3CircleY - intShipMiddleY                                     'y distance between turret3circle and ship
    intTurret3ShipDistance = Sqr((intTurret3YDistance * intTurret3YDistance) + (intTurret3XDistance * intTurret3XDistance)) 'distance of the xy distance
    intTurret3Ratio = intTurret3Length / intTurret3ShipDistance                                  'how many pixels long is the turret3
    intTurret3X = intTurret3CircleX - (intTurret3Ratio * intTurret3XDistance)                    'plot x end of turret3
    intTurret3Y = intTurret3CircleY - (intTurret3Ratio * intTurret3YDistance)                    'plot y end of turret3
    backbuffer.DrawLine intTurret3CircleX, intTurret3CircleY, intTurret3X, intTurret3Y           'draw turret3
    backbuffer.DrawCircle intTurret3CircleX, intTurret3CircleY, intCircleWidth                   'draw turret3 base
End Sub

Sub generateTurret3Shot()
    If blnShipLanded = False Then                                                                'dont shoot at our ship if its landed
        blnTurret3Shot = True                                                                    'make sure that the shot always gets drawn
        PlayPhotonSound                                                                          'sound effect
        intTurret3ShotLength = yres / 100
        intTurret3ShotXDistance = intTurret3CircleX - intShipMiddleX
        intTurret3ShotYDistance = intTurret3CircleY - intShipMiddleY
        intTurret3ShotDistance = Sqr((intTurret3ShotYDistance * intTurret3ShotYDistance) + (intTurret3ShotXDistance * intTurret3ShotXDistance))
    End If
End Sub

Sub drawTurret3Shot()
    backbuffer.SetFillStyle 0
    backbuffer.SetForeColor RGB(206, 159, 255)
    backbuffer.SetFillColor RGB(206, 159, 255)
    intTurret3ShotLength = intTurret3ShotLength + 4
    intTurret3ShotRatio = intTurret3ShotLength / intTurret3ShotDistance
    intTurret3ShotX = intTurret3CircleX - (intTurret3ShotRatio * intTurret3ShotXDistance)
    intTurret3ShotY = intTurret3CircleY - (intTurret3ShotRatio * intTurret3ShotYDistance)
    backbuffer.DrawCircle intTurret3ShotX, intTurret3ShotY, 2
    If intTurret3ShotX >= xres Then                                                               'has the shot hit the right of the screen?
        blnTurret3Shot = False
    End If
    If intTurret3ShotX <= 0 Then                                                                  'has the shot hit the left of the screen?
        blnTurret3Shot = False
    End If
    If intTurret3ShotY <= 0 Then                                                                  'has the shot hit the top of the screen?
        blnTurret3Shot = False
    End If
    If intTurret3ShotY >= intLandscapeYCollision(intTurret3ShotX) Then                             'has the shot hit the landscape?
        blnTurret3Shot = False
    End If
    If intTurret3ShotX <= (intShipMiddleX + intShipDistMidSide) And intTurret3ShotX >= (intShipMiddleX - intShipDistMidSide) Then 'has shot hit players ship?
        If intTurret3ShotY <= (intShipMiddleY + intShipDistMidBott) And intTurret3ShotY >= (intShipMiddleY - intShipDistMidBott) Then
            blnTurretShot = False
            blnTurret2Shot = False
            blnTurret3Shot = False
            blnTurret4Shot = False
            blnTurret5Shot = False
            blnTurret6Shot = False
            blnTurret7Shot = False
            blnCollided = True
            LifeLost
        End If
    End If
End Sub

Sub GenerateTurret4()                                                                            'work out where a gun turret4 is to be placed
Randomize Timer
intTurret4CircleX = Rnd * xres                                                                   'create a random x position
If intTurret4CircleX <= intPadRightTopLeftx + 10 Then                                            'make sure its further right than landing pad
    intTurret4CircleX = intTurret4CircleX + intPadRightTopLeftx + 10
End If
If intTurret4CircleX >= intStatsXCoordinate Then
    intTurret4CircleX = xres / 2
End If
intTurret4CircleY = yres / (0.5 * yres)  '(2)                                                    'set the y position so its hanging from the roof
blnTurret4 = True                                                                                'set flag to say that gun turret is on screen
End Sub

Sub DrawTurret4()
    backbuffer.SetFillStyle 0
    backbuffer.SetFillColor RGB(50, 50, 50)
    backbuffer.SetForeColor RGB(150, 150, 150)
                                                                                                 'calculate the coordinates of the turret4 itself
    intTurret4Length = yres / 50                                                                 'the length of the turret4
    intTurret4XDistance = intTurret4CircleX - intShipMiddleX                                     'x distance between turret4circle and ship
    intTurret4YDistance = intTurret4CircleY - intShipMiddleY                                     'y distance between turret4circle and ship
    intTurret4ShipDistance = Sqr((intTurret4YDistance * intTurret4YDistance) + (intTurret4XDistance * intTurret4XDistance)) 'distance of the xy distance
    intTurret4Ratio = intTurret4Length / intTurret4ShipDistance                                  'how many pixels long is the turret4
    intTurret4X = intTurret4CircleX - (intTurret4Ratio * intTurret4XDistance)                    'plot x end of turret4
    intTurret4Y = intTurret4CircleY - (intTurret4Ratio * intTurret4YDistance)                    'plot y end of turret4
    backbuffer.DrawLine intTurret4CircleX, intTurret4CircleY, intTurret4X, intTurret4Y           'draw turret4
    backbuffer.DrawCircle intTurret4CircleX, intTurret4CircleY, intCircleWidth                   'draw turret4 base
End Sub

Sub generateTurret4Shot()
    If blnShipLanded = False Then                                                                'dont shoot at our ship if its landed
        blnTurret4Shot = True                                                                    'make sure that the shot always gets drawn
        PlayPhotonSound                                                                          'sound effect
        intTurret4ShotLength = yres / 100
        intTurret4ShotXDistance = intTurret4CircleX - intShipMiddleX
        intTurret4ShotYDistance = intTurret4CircleY - intShipMiddleY
        intTurret4ShotDistance = Sqr((intTurret4ShotYDistance * intTurret4ShotYDistance) + (intTurret4ShotXDistance * intTurret4ShotXDistance))
    End If
End Sub

Sub drawTurret4Shot()
    backbuffer.SetFillStyle 0
    backbuffer.SetForeColor RGB(206, 159, 255)
    backbuffer.SetFillColor RGB(206, 159, 255)
    intTurret4ShotLength = intTurret4ShotLength + 4
    intTurret4ShotRatio = intTurret4ShotLength / intTurret4ShotDistance
    intTurret4ShotX = intTurret4CircleX - (intTurret4ShotRatio * intTurret4ShotXDistance)
    intTurret4ShotY = intTurret4CircleY - (intTurret4ShotRatio * intTurret4ShotYDistance)
    backbuffer.DrawCircle intTurret4ShotX, intTurret4ShotY, 2
    If intTurret4ShotX >= xres Then                                                               'has the shot hit the right of the screen?
        blnTurret4Shot = False
    End If
    If intTurret4ShotX <= 0 Then                                                                  'has the shot hit the left of the screen?
        blnTurret4Shot = False
    End If
    If intTurret4ShotY <= 0 Then                                                                  'has the shot hit the top of the screen?
        blnTurret4Shot = False
    End If
    If intTurret4ShotY >= intLandscapeYCollision(intTurret4ShotX) Then                             'has the shot hit the landscape?
        blnTurret4Shot = False
    End If
    If intTurret4ShotX <= (intShipMiddleX + intShipDistMidSide) And intTurret4ShotX >= (intShipMiddleX - intShipDistMidSide) Then 'has shot hit players ship?
        If intTurret4ShotY <= (intShipMiddleY + intShipDistMidBott) And intTurret4ShotY >= (intShipMiddleY - intShipDistMidBott) Then
            blnTurretShot = False
            blnTurret2Shot = False
            blnTurret3Shot = False
            blnTurret4Shot = False
            blnTurret5Shot = False
            blnTurret6Shot = False
            blnTurret7Shot = False
            blnCollided = True
            LifeLost
        End If
    End If
End Sub

Sub GenerateTurret5()                                                                            'work out where a gun turret5 is to be placed
Randomize Timer
intTurret5CircleX = Rnd * xres                                                                   'create a random x position
If intTurret5CircleX <= intPadRightTopLeftx + 10 Then                                            'make sure its further right than landing pad
    intTurret5CircleX = intTurret5CircleX + intPadRightTopLeftx + 10
End If
If intTurret5CircleX >= intStatsXCoordinate Then
    intTurret5CircleX = xres / 2
End If
intTurret5CircleY = yres / (0.5 * yres)  '(2)                                                    'set the y position so its hanging from the roof
blnTurret5 = True                                                                                'set flag to say that gun turret is on screen
End Sub

Sub DrawTurret5()
    backbuffer.SetFillStyle 0
    backbuffer.SetFillColor RGB(50, 50, 50)
    backbuffer.SetForeColor RGB(150, 150, 150)
                                                                                                 'calculate the coordinates of the turret5 itself
    intTurret5Length = yres / 50                                                                 'the length of the turret5
    intTurret5XDistance = intTurret5CircleX - intShipMiddleX                                     'x distance between turret5circle and ship
    intTurret5YDistance = intTurret5CircleY - intShipMiddleY                                     'y distance between turret5circle and ship
    intTurret5ShipDistance = Sqr((intTurret5YDistance * intTurret5YDistance) + (intTurret5XDistance * intTurret5XDistance)) 'distance of the xy distance
    intTurret5Ratio = intTurret5Length / intTurret5ShipDistance                                  'how many pixels long is the turret5
    intTurret5X = intTurret5CircleX - (intTurret5Ratio * intTurret5XDistance)                    'plot x end of turret5
    intTurret5Y = intTurret5CircleY - (intTurret5Ratio * intTurret5YDistance)                    'plot y end of turret5
    backbuffer.DrawLine intTurret5CircleX, intTurret5CircleY, intTurret5X, intTurret5Y           'draw turret5
    backbuffer.DrawCircle intTurret5CircleX, intTurret5CircleY, intCircleWidth                   'draw turret5 base
End Sub

Sub generateTurret5Shot()
    If blnShipLanded = False Then                                                                'dont shoot at our ship if its landed
        blnTurret5Shot = True                                                                    'make sure that the shot always gets drawn
        PlayPhotonSound                                                                          'sound effect
        intTurret5ShotLength = yres / 100
        intTurret5ShotXDistance = intTurret5CircleX - intShipMiddleX
        intTurret5ShotYDistance = intTurret5CircleY - intShipMiddleY
        intTurret5ShotDistance = Sqr((intTurret5ShotYDistance * intTurret5ShotYDistance) + (intTurret5ShotXDistance * intTurret5ShotXDistance))
    End If
End Sub

Sub drawTurret5Shot()
    backbuffer.SetFillStyle 0
    backbuffer.SetForeColor RGB(206, 159, 255)
    backbuffer.SetFillColor RGB(206, 159, 255)
    intTurret5ShotLength = intTurret5ShotLength + 4
    intTurret5ShotRatio = intTurret5ShotLength / intTurret5ShotDistance
    intTurret5ShotX = intTurret5CircleX - (intTurret5ShotRatio * intTurret5ShotXDistance)
    intTurret5ShotY = intTurret5CircleY - (intTurret5ShotRatio * intTurret5ShotYDistance)
    backbuffer.DrawCircle intTurret5ShotX, intTurret5ShotY, 2
    If intTurret5ShotX >= xres Then                                                               'has the shot hit the right of the screen?
        blnTurret5Shot = False
    End If
    If intTurret5ShotX <= 0 Then                                                                  'has the shot hit the left of the screen?
        blnTurret5Shot = False
    End If
    If intTurret5ShotY <= 0 Then                                                                  'has the shot hit the top of the screen?
        blnTurret5Shot = False
    End If
    If intTurret5ShotY >= intLandscapeYCollision(intTurret5ShotX) Then                             'has the shot hit the landscape?
        blnTurret5Shot = False
    End If
    If intTurret5ShotX <= (intShipMiddleX + intShipDistMidSide) And intTurret5ShotX >= (intShipMiddleX - intShipDistMidSide) Then 'has shot hit players ship?
        If intTurret5ShotY <= (intShipMiddleY + intShipDistMidBott) And intTurret5ShotY >= (intShipMiddleY - intShipDistMidBott) Then
            blnTurretShot = False
            blnTurret2Shot = False
            blnTurret3Shot = False
            blnTurret4Shot = False
            blnTurret5Shot = False
            blnTurret6Shot = False
            blnTurret7Shot = False
            blnCollided = True
            LifeLost
        End If
    End If
End Sub

Sub Generateturret6()                                                                            'work out where a gun turret is to be placed
Randomize Timer
intTurret6CircleX = Rnd * xres                                                                   'create a random x position
intTurret6CircleY = intLandscapeYCollision(intTurret6CircleX) - 8                                 'set the y position to just above the land level
blnTurret6 = True                                                                                'set flag to say that gun turret is on screen
End Sub

Sub Drawturret6()
    backbuffer.SetFillStyle 0
    backbuffer.SetFillColor RGB(50, 50, 50)
    backbuffer.SetForeColor RGB(150, 150, 150)
                                                                                                'calculate the coordinates of the turret itself
    intTurret6Length = yres / 50                                                                 'the length of the turret
    intTurret6XDistance = intTurret6CircleX - intShipMiddleX                                      'x distance between turretcircle and ship
    intTurret6YDistance = intTurret6CircleY - intShipMiddleY                                      'y distance between turretcircle and ship
    intTurret6ShipDistance = Sqr((intTurret6YDistance * intTurret6YDistance) + (intTurret6XDistance * intTurret6XDistance)) 'distance of the xy distance
    intTurret6Ratio = intTurret6Length / intTurret6ShipDistance                                    'how many pixels long is the turret
    intTurret6X = intTurret6CircleX - (intTurret6Ratio * intTurret6XDistance)                       'plot x end of turret
    intTurret6Y = intTurret6CircleY - (intTurret6Ratio * intTurret6YDistance)                       'plot y end of turret
    backbuffer.DrawLine intTurret6CircleX, intTurret6CircleY, intTurret6X, intTurret6Y              'draw turret
    backbuffer.DrawCircle intTurret6CircleX, intTurret6CircleY, intCircleWidth                      'draw turret base
End Sub

Sub generateturret6Shot()
    If blnShipLanded = False Then                                                               'dont shoot at our ship if its landed
        blnTurret6Shot = True                                                                    'make sure that the shot always gets drawn
        PlayPhotonSound                                                                         'sound effect
        intTurret6ShotLength = yres / 100
        intTurret6ShotXDistance = intTurret6CircleX - intShipMiddleX
        intTurret6ShotYDistance = intTurret6CircleY - intShipMiddleY
        intTurret6ShotDistance = Sqr((intTurret6ShotYDistance * intTurret6ShotYDistance) + (intTurret6ShotXDistance * intTurret6ShotXDistance))
    End If
End Sub

Sub drawturret6Shot()
    backbuffer.SetFillStyle 0
    backbuffer.SetForeColor RGB(206, 159, 255)
    backbuffer.SetFillColor RGB(206, 159, 255)
    intTurret6ShotLength = intTurret6ShotLength + 4
    intTurret6ShotRatio = intTurret6ShotLength / intTurret6ShotDistance
    intTurret6ShotX = intTurret6CircleX - (intTurret6ShotRatio * intTurret6ShotXDistance)
    intTurret6ShotY = intTurret6CircleY - (intTurret6ShotRatio * intTurret6ShotYDistance)
    backbuffer.DrawCircle intTurret6ShotX, intTurret6ShotY, 2
    If intTurret6ShotX >= xres Then                                                               'has the shot hit the right of the screen?
        blnTurret6Shot = False
    End If
    If intTurret6ShotX <= 0 Then                                                                  'has the shot hit the left of the screen?
        blnTurret6Shot = False
    End If
    If intTurret6ShotY <= 0 Then                                                                  'has the shot hit the top of the screen?
        blnTurret6Shot = False
    End If
    If intTurret6ShotY >= intLandscapeYCollision(intTurret6ShotX) Then                             'has the shot hit the landscape?
        blnTurret6Shot = False
    End If
    If intTurret6ShotX <= (intShipMiddleX + intShipDistMidSide) And intTurret6ShotX >= (intShipMiddleX - intShipDistMidSide) Then 'has shot hit players ship?
        If intTurret6ShotY <= (intShipMiddleY + intShipDistMidBott) And intTurret6ShotY >= (intShipMiddleY - intShipDistMidBott) Then
            blnTurretShot = False
            blnTurret2Shot = False
            blnTurret3Shot = False
            blnTurret4Shot = False
            blnTurret5Shot = False
            blnTurret6Shot = False
            blnTurret7Shot = False
            blnCollided = True
            LifeLost
        End If
    End If
End Sub


Sub Generateturret7()                                                                            'work out where a gun turret is to be placed
Randomize Timer
intTurret7CircleX = Rnd * xres                                                                   'create a random x position
intTurret7CircleY = intLandscapeYCollision(intTurret7CircleX) - 8                                 'set the y position to just above the land level
blnTurret7 = True                                                                                'set flag to say that gun turret is on screen
End Sub

Sub Drawturret7()
    backbuffer.SetFillStyle 0
    backbuffer.SetFillColor RGB(50, 50, 50)
    backbuffer.SetForeColor RGB(150, 150, 150)
                                                                                                'calculate the coordinates of the turret itself
    intTurret7Length = yres / 50                                                                 'the length of the turret
    intTurret7XDistance = intTurret7CircleX - intShipMiddleX                                      'x distance between turretcircle and ship
    intTurret7YDistance = intTurret7CircleY - intShipMiddleY                                      'y distance between turretcircle and ship
    intTurret7ShipDistance = Sqr((intTurret7YDistance * intTurret7YDistance) + (intTurret7XDistance * intTurret7XDistance)) 'distance of the xy distance
    intTurret7Ratio = intTurret7Length / intTurret7ShipDistance                                    'how many pixels long is the turret
    intTurret7X = intTurret7CircleX - (intTurret7Ratio * intTurret7XDistance)                       'plot x end of turret
    intTurret7Y = intTurret7CircleY - (intTurret7Ratio * intTurret7YDistance)                       'plot y end of turret
    backbuffer.DrawLine intTurret7CircleX, intTurret7CircleY, intTurret7X, intTurret7Y              'draw turret
    backbuffer.DrawCircle intTurret7CircleX, intTurret7CircleY, intCircleWidth                      'draw turret base
End Sub

Sub generateturret7Shot()
    If blnShipLanded = False Then                                                               'dont shoot at our ship if its landed
        blnTurret7Shot = True                                                                    'make sure that the shot always gets drawn
        PlayPhotonSound                                                                         'sound effect
        intTurret7ShotLength = yres / 100
        intTurret7ShotXDistance = intTurret7CircleX - intShipMiddleX
        intTurret7ShotYDistance = intTurret7CircleY - intShipMiddleY
        intTurret7ShotDistance = Sqr((intTurret7ShotYDistance * intTurret7ShotYDistance) + (intTurret7ShotXDistance * intTurret7ShotXDistance))
    End If
End Sub

Sub drawturret7Shot()
    backbuffer.SetFillStyle 0
    backbuffer.SetForeColor RGB(206, 159, 255)
    backbuffer.SetFillColor RGB(206, 159, 255)
    intTurret7ShotLength = intTurret7ShotLength + 4
    intTurret7ShotRatio = intTurret7ShotLength / intTurret7ShotDistance
    intTurret7ShotX = intTurret7CircleX - (intTurret7ShotRatio * intTurret7ShotXDistance)
    intTurret7ShotY = intTurret7CircleY - (intTurret7ShotRatio * intTurret7ShotYDistance)
    backbuffer.DrawCircle intTurret7ShotX, intTurret7ShotY, 2
    If intTurret7ShotX >= xres Then                                                               'has the shot hit the right of the screen?
        blnTurret7Shot = False
    End If
    If intTurret7ShotX <= 0 Then                                                                  'has the shot hit the left of the screen?
        blnTurret7Shot = False
    End If
    If intTurret7ShotY <= 0 Then                                                                  'has the shot hit the top of the screen?
        blnTurret7Shot = False
    End If
    If intTurret7ShotY >= intLandscapeYCollision(intTurret7ShotX) Then                             'has the shot hit the landscape?
        blnTurret7Shot = False
    End If
    If intTurret7ShotX <= (intShipMiddleX + intShipDistMidSide) And intTurret7ShotX >= (intShipMiddleX - intShipDistMidSide) Then 'has shot hit players ship?
        If intTurret7ShotY <= (intShipMiddleY + intShipDistMidBott) And intTurret7ShotY >= (intShipMiddleY - intShipDistMidBott) Then
            blnTurretShot = False
            blnTurret2Shot = False
            blnTurret3Shot = False
            blnTurret4Shot = False
            blnTurret5Shot = False
            blnTurret6Shot = False
            blnTurret7Shot = False
            blnCollided = True
            LifeLost
        End If
    End If
End Sub


Sub GenerateFuel()                                                                              'work out where a fuel capsule is to be placed
Randomize Timer
intFuelX = Rnd * xres                                                                           'create a random x position
intFuelY = intLandscapeYCollision(intFuelX) - 14                                                'set the y position to just above the land level
PlayItemAppearsSound                                                                            'ping!!
blnFuel = True                                                                                  'set flag to say that fuel is on screen
End Sub

Sub drawFuel()                                                                                  'draw the fuel
backbuffer.SetFillStyle 0                                                                       'make the fuel capsule filled
backbuffer.SetFillColor RGB(100, 100, 150)                                                      'with this colour
backbuffer.SetForeColor RGB(0, 0, 100)                                                          'text (F) colour
backbuffer.DrawCircle intFuelX, intFuelY, intCircleWidth                                        'draw the capsule
'backbuffer.DrawText intFuelX - 3, intFuelY - 8, "F", False                                      'and it letter (F)
End Sub

Sub GenerateSwitch()                                                                            'work out where the switch is to be placed
Randomize Timer
intSwitchX = Rnd * xres                                                                         'create a random x position
intSwitchY = intLandscapeYCollision(intSwitchX) - (yres / 120)                                          'set the y position to just above the land level
blnSwitchPressed = False                                                                        'make sure the switch is initialised ready to be pressed
End Sub

Sub drawSwitch()                                                                                'draw the switch
backbuffer.SetFillStyle 0                                                                       'make the switch filled
backbuffer.SetFillColor RGB(50, 50, 50)                                                         'with this colour
backbuffer.SetForeColor RGB(246, 253, 78)                                                       'colour of the switch
If blnSwitchPressed = False Then                                                                'if not switched
    intSwitchButtonTopY = intSwitchY - (yres / 90)                                              'set the height of the switch so it looks unpressed
Else                                                                                            'if switched
    intSwitchButtonTopY = intSwitchY - (yres / 110)                                             'set the height of the switch so it looks pressed
End If
intSwitchButtonTopX = intSwitchX - (xres / 533)
intSwitchButtonBottomX = intSwitchX + (xres / 533)
intSwitchButtonBottomY = intSwitchY
backbuffer.DrawBox intSwitchButtonTopX, intSwitchButtonTopY, intSwitchButtonBottomX, intSwitchButtonBottomY 'draw the switch
backbuffer.DrawBox intSwitchX - (xres / 230), intSwitchY - (xres / 230), intSwitchX + (xres / 230), intSwitchY + (xres / 230) 'draw the switchbox
End Sub


Sub GenerateDrugs()
                                                                                                'remove the effects of any drugs that may already have been taken
intRotationRate = 2                                                                             'in case player had drugs that sped up rotation
intDrugType = 3                                                                                 'in case player had drugs that reversed rotation... drugtype of 3 is the equivalent of no drugs

Randomize Timer
intRandomNumber = Rnd(Timer)                                                                    'create a random number between 0-1 to create a 50/50 chance of whether the drugs appear from the left of right of screen
If intRandomNumber > 0.5 Then
    intDrugsX = 0                                                                               'drugs will appear from left of screen
    blnDrugsIncrement = True                                                                    'means the x position will be incremented
Else
    intDrugsX = xres                                                                            'drugs will appear from right of screen
    blnDrugsIncrement = False                                                                   'means the x position will be decremented
End If
intDrugsY = intRandomNumber * (yres - (yres - intLandscapeHighestPoint))                        'creates random height for drugs, ensuring its higher than heighest landscape point
PlayItemAppearsSound                                                                            'ping!!
blnDrugs = True                                                                                 'set flag to say drugs are on screen
End Sub

Sub drawDrugs()
backbuffer.SetFillStyle 0                                                                       'to make sure the drug capsule is filled
backbuffer.SetFillColor RGB(100, 150, 100)                                                      'colour of drug capsule
backbuffer.SetForeColor RGB(0, 70, 0)                                                           'colour of outline
If blnDrugsIncrement = True Then                                                                'moving from right to left
    intDrugsX = intDrugsX + 3                                                                   'move the drugs right
    If intDrugsX >= xres Then                                                                   'check if we've reached edge of screen
        blnDrugs = False                                                                        'if we have, stop drawing the drugs
    End If
Else                                                                                            'moving from left to right
    intDrugsX = intDrugsX - 3                                                                   'move drugs left
    If intDrugsX <= 0 Then                                                                      'check if we've reached the edge of the screen
        blnDrugs = False                                                                        'if we have, stop drawing the drugs
    End If
End If
intDrugsY = intDrugsY - 3 * Sin((4 * pi * intDrugsX) / intDrugsWavelength)                      'follow sine wave path
If intDrugsY > intLandscapeHighestPoint Then                                                    'if drugs reach landscape level...
    intDrugsY = intLandscapeHighestPoint                                                        'let them go no lower
End If
backbuffer.DrawCircle intDrugsX, intDrugsY, intCircleWidth                                      'draw the drugs capsule
'backbuffer.DrawText intDrugsX - 3, intDrugsY - 8, "d", False                                    'and the 'd' within it
End Sub

Sub GenerateAntiGravity()
Randomize Timer                                                                                 'create a random number between 0-1 to create a 50/50 chance of whether the drugs appear from the left of right of screen
intRandomNumber = Rnd(Timer)
If intRandomNumber > 0.5 Then
    intAntiGravityX = 0                                                                         'antigrav will appear from the left of the screen
    blnAntiGravityIncrement = True                                                              'means antigrav x position will be incremented
Else
    intAntiGravityX = xres                                                                      'antigrav will appear from the right of the screen
    blnAntiGravityIncrement = False                                                             'means antigrav x position will be decremented
End If
intAntiGravityY = intRandomNumber * (yres - (yres - intLandscapeHighestPoint))                  'creates random height for antigrav power up, ensuring its higher than heighest landscape point
PlayItemAppearsSound                                                                            'ping!!
blnAntiGravityShown = True                                                                      'set flag to say antigrav is on screen
End Sub
 
Sub drawAntiGravity()
backbuffer.SetFillStyle 0                                                                       'make sure antigrav will be filled
backbuffer.SetFillColor RGB(150, 100, 100)                                                      'fill colour
backbuffer.SetForeColor RGB(70, 0, 0)                                                           'draw colour
If blnAntiGravityIncrement = True Then                                                          'if moving from left to right
    intAntiGravityX = intAntiGravityX + 3                                                       'move anti grav power up right
    If intAntiGravityX >= xres Then                                                             'check if we've reached the edge of the screen
        blnAntiGravityShown = False                                                             'if we have, stop drawing the anti grav power up
    End If
Else                                                                                            'moving from right to left
    intAntiGravityX = intAntiGravityX - 3                                                       'move anti grav power up left
    If intAntiGravityX <= 0 Then                                                                'check if we've reached the edge of the screen
        blnAntiGravityShown = False                                                             'if we have, stop drawing the anti grav power up
    End If
End If
                                                                                                'make anti-grav follow two compounded sine waves
intAntiGravityY = intAntiGravityY - 2 * Sin((2 * pi * intAntiGravityX) / intAntiGravityWavelength1) + ((0.2 * pi * intAntiGravityX) / intAntiGravityWavelength2)
If intAntiGravityY > intLandscapeHighestPoint Then                                              'if antigrav reaches landscape level
    intAntiGravityY = intLandscapeHighestPoint                                                  'let it go no lower
End If
backbuffer.DrawCircle intAntiGravityX, intAntiGravityY, intCircleWidth                          'draw the antigrav capsule
'backbuffer.DrawText intAntiGravityX - 4, intAntiGravityY - 10, "g", False                       'and it's 'g' caption
End Sub

Sub GenerateScientist1()
Randomize Timer                                                                                 'random number will be used to say whether scientist appears from left or right
intRandomNumber = Rnd(Timer)
If intRandomNumber > 0.5 Then
    intScientist1X = 0                                                                          'scientist will appear from left
    blnScientist1Increment = True                                                               'makes sure x position will be incremented
Else
    intScientist1X = xres                                                                       'scientist will appear from right
    blnScientist1Increment = False                                                              'makes sure x position will be decremented
End If
intScientist1Y = intLandscapeYCollision(intScientist1X) - intScientistHeightFromGround          'make scientist y position just above the landscape
PlayNewScientistSound                                                                           'ping!!
blnScientist1 = True                                                                            'flag to say scientist 1 is on screen
intScientistsOutCount = intScientistsOutCount + 1                                               'add to the number of scientists out
End Sub

Sub DrawScientist1()
    backbuffer.SetFillStyle 0
    backbuffer.SetFillColor RGB(100, 100, 100)
    backbuffer.SetForeColor RGB(255, 250, 70)
If blnDanceMode = False Then                                                                    'if dance mode is on, the dance routine will deal with drawing the scientist!
    If blnScientist1Increment = True Then                                                       'if moving from left to right...
        intScientist1X = intScientist1X + 1                                                     'move scientist right
        If intScientist1X = xres Then                                                           'if we've reached edge of screen
            blnScientist1Increment = False                                                      'make sure x position moves from right to left from now on
        End If
    Else                                                                                        'if moving right to left...
        intScientist1X = intScientist1X - 1                                                     'move scientist left
        If intScientist1X <= 1 Then                                                             'if we've reached edge of screen
            blnScientist1Increment = True                                                       'make sure x position moves from left to right from now on
        End If
    End If
    intScientist1Y = intLandscapeYCollision(intScientist1X) - intScientistHeightFromGround      'set y position of scientist1 to just above the landscape
End If
backbuffer.DrawLine intScientist1X, intScientist1Y, intScientist1X, intScientist1Y + intScientistHeight 'draw the scientist
End Sub

Sub DrawFreeScientist1()
blnScientist1Free = True                                                                        'flag the scientist as free
backbuffer.SetFillStyle 0
backbuffer.SetFillColor RGB(100, 100, 100)
backbuffer.SetForeColor RGB(255, 250, 70)
intScientist1X = intScientist1X - (xres / 10000)                                                'move scientist from ship to left of landing pad
backbuffer.DrawLine intScientist1X, intScientist1Y, intScientist1X, intScientist1Y - intScientistHeight 'draw scientist
If intScientist1X <= 0 And blnTakeOffSoundPlayed = False And blnScientist1OnBoard = True Then   'if scientist1 has made it to the exit
    AddScore (500)
    CheckForLevelComplete (1)                                                                   'was this the last scientist to be rescued? if so, level is complete
    PlayTakeoffSound                                                                            'clear for take off
    blnTakeOffSoundPlayed = True                                                                'dont play this sound again
    intScientist1X = 2000                                                                       'stick scientist1 well off screen
End If
If blnTakeOffSoundPlayed = True Then                                                            'if ship has finished landing
    intTakeOffDelay = intTakeOffDelay + 1                                                       'force a delay between land and take off
    If intTakeOffDelay = 600 Then                                                               'check for time to take off
        blnScientist1Complete = True
    End If
End If
End Sub

Sub GenerateScientist2()                                                                        'same as generatescientist1
  Randomize Timer
intRandomNumber = Rnd(Timer)
If intRandomNumber > 0.5 Then
    intScientist2X = 0
    blnScientist2Increment = True
Else
    intScientist2X = xres
    blnScientist2Increment = False
End If
intScientist2Y = intLandscapeYCollision(intScientist2X) - intScientistHeightFromGround
PlayNewScientistSound
blnScientist2 = True
intScientistsOutCount = intScientistsOutCount + 1
End Sub

Sub DrawScientist2()                                                                            'same as drawscientist1
backbuffer.SetFillStyle 0
backbuffer.SetFillColor RGB(100, 100, 100)
backbuffer.SetForeColor RGB(255, 250, 70)
If blnDanceMode = False Then
    If blnScientist2Increment = True Then
        intScientist2X = intScientist2X + 1
        If intScientist2X = xres Then
            blnScientist2Increment = False
        End If
    Else
        intScientist2X = intScientist2X - 1
        If intScientist2X <= 1 Then
            blnScientist2Increment = True
        End If
    End If
    intScientist2Y = intLandscapeYCollision(intScientist2X) - intScientistHeightFromGround
End If
backbuffer.DrawLine intScientist2X, intScientist2Y, intScientist2X, intScientist2Y + intScientistHeight
End Sub

Sub DrawFreeScientist2()                                                                        'same as drawfreescientist1
blnScientist2Free = True
backbuffer.SetFillStyle 0
backbuffer.SetFillColor RGB(100, 100, 100)
backbuffer.SetForeColor RGB(255, 250, 70)
intScientist2X = intScientist2X - (xres / 10000)
backbuffer.DrawLine intScientist2X, intScientist2Y, intScientist2X, intScientist2Y - intScientistHeight
If intScientist2X <= 0 And blnTakeOffSoundPlayed = False And blnScientist2OnBoard = True Then
    AddScore (500)
    CheckForLevelComplete (2)
    PlayTakeoffSound
    blnTakeOffSoundPlayed = True
    intScientist2X = 2000
End If
If blnTakeOffSoundPlayed = True Then
    intTakeOffDelay = intTakeOffDelay + 1
    If intTakeOffDelay = 600 Then
        blnScientist2Complete = True
    End If
End If
End Sub


Sub GenerateScientist3()                                                                        'same as generatescientist1
Randomize Timer
intRandomNumber = Rnd(Timer)
If intRandomNumber > 0.5 Then
    intScientist3X = 0
    blnScientist3Increment = True
Else
    intScientist3X = xres
    blnScientist3Increment = False
End If
intScientist3Y = intLandscapeYCollision(intScientist3X) - intScientistHeightFromGround
PlayNewScientistSound
blnScientist3 = True
intScientistsOutCount = intScientistsOutCount + 1
End Sub

Sub DrawScientist3()                                                                            'same as drawscientist1
backbuffer.SetFillStyle 0
backbuffer.SetFillColor RGB(100, 100, 100)
backbuffer.SetForeColor RGB(255, 250, 70)
If blnDanceMode = False Then
    If blnScientist3Increment = True Then
        intScientist3X = intScientist3X + 1
        If intScientist3X = xres Then
            blnScientist3Increment = False
        End If
    Else
        intScientist3X = intScientist3X - 1
        If intScientist3X <= 1 Then
            blnScientist3Increment = True
        End If
    End If
    intScientist3Y = intLandscapeYCollision(intScientist3X) - intScientistHeightFromGround
End If
backbuffer.DrawLine intScientist3X, intScientist3Y, intScientist3X, intScientist3Y + intScientistHeight
End Sub

Sub DrawFreeScientist3()                                                                        'same as drawfreescientist1
blnScientist3Free = True
backbuffer.SetFillStyle 0
backbuffer.SetFillColor RGB(100, 100, 100)
backbuffer.SetForeColor RGB(255, 250, 70)
intScientist3X = intScientist3X - (xres / 10000)
backbuffer.DrawLine intScientist3X, intScientist3Y, intScientist3X, intScientist3Y - intScientistHeight
If intScientist3X <= 0 And blnTakeOffSoundPlayed = False And blnScientist3OnBoard = True Then
    AddScore (500)
    CheckForLevelComplete (3)
    PlayTakeoffSound
    blnTakeOffSoundPlayed = True
    intScientist3X = 2000
End If
If blnTakeOffSoundPlayed = True Then
    intTakeOffDelay = intTakeOffDelay + 1
    If intTakeOffDelay = 600 Then
        blnScientist3Complete = True
    End If
End If
End Sub


Sub GenerateScientist4()                                                                        'same as generatescientist1
Randomize Timer
intRandomNumber = Rnd(Timer)
If intRandomNumber > 0.5 Then
    intScientist4X = 0
    blnScientist4Increment = True
Else
    intScientist4X = xres
    blnScientist4Increment = False
End If
intScientist4Y = intLandscapeYCollision(intScientist4X) - intScientistHeightFromGround
PlayNewScientistSound
blnScientist4 = True
intScientistsOutCount = intScientistsOutCount + 1
End Sub

Sub DrawScientist4()                                                                            'same as drawscientist1
backbuffer.SetFillStyle 0
backbuffer.SetFillColor RGB(100, 100, 100)
backbuffer.SetForeColor RGB(255, 250, 70)
If blnDanceMode = False Then
    If blnScientist4Increment = True Then
        intScientist4X = intScientist4X + 1
        If intScientist4X = xres Then
            blnScientist4Increment = False
        End If
    Else
        intScientist4X = intScientist4X - 1
        If intScientist4X <= 1 Then
            blnScientist4Increment = True
        End If
    End If
    intScientist4Y = intLandscapeYCollision(intScientist4X) - intScientistHeightFromGround
End If
backbuffer.DrawLine intScientist4X, intScientist4Y, intScientist4X, intScientist4Y + intScientistHeight
End Sub

Sub DrawFreeScientist4()                                                                        'same as drawfreescientist1
blnScientist4Free = True
backbuffer.SetFillStyle 0
backbuffer.SetFillColor RGB(100, 100, 100)
backbuffer.SetForeColor RGB(255, 250, 70)
intScientist4X = intScientist4X - (xres / 10000)
backbuffer.DrawLine intScientist4X, intScientist4Y, intScientist4X, intScientist4Y - intScientistHeight
If intScientist4X <= 0 And blnTakeOffSoundPlayed = False And blnScientist4OnBoard = True Then
    AddScore (500)
    CheckForLevelComplete (4)
    PlayTakeoffSound
    blnTakeOffSoundPlayed = True
    intScientist4X = 2000
End If
If blnTakeOffSoundPlayed = True Then
    intTakeOffDelay = intTakeOffDelay + 1
    If intTakeOffDelay = 600 Then
        blnScientist4Complete = True
    End If
End If
End Sub


Sub GenerateScientist5()                                                                        'same as generatescientist1
Randomize Timer
intRandomNumber = Rnd(Timer)
If intRandomNumber > 0.5 Then
    intScientist5X = 0
    blnScientist5Increment = True
Else
    intScientist5X = xres
    blnScientist5Increment = False
End If
intScientist5Y = intLandscapeYCollision(intScientist5X) - intScientistHeightFromGround
PlayNewScientistSound
blnScientist5 = True
intScientistsOutCount = intScientistsOutCount + 1
End Sub

Sub DrawScientist5()                                                                            'same as drawscientist1
backbuffer.SetFillStyle 0
backbuffer.SetFillColor RGB(100, 100, 100)
backbuffer.SetForeColor RGB(255, 250, 70)
If blnDanceMode = False Then
    If blnScientist5Increment = True Then
        intScientist5X = intScientist5X + 1
        If intScientist5X = xres Then
            blnScientist5Increment = False
        End If
    Else
        intScientist5X = intScientist5X - 1
        If intScientist5X <= 1 Then
            blnScientist5Increment = True
        End If
    End If
    intScientist5Y = intLandscapeYCollision(intScientist5X) - intScientistHeightFromGround
End If
backbuffer.DrawLine intScientist5X, intScientist5Y, intScientist5X, intScientist5Y + intScientistHeight
End Sub

Sub DrawFreeScientist5()                                                                        'same as drawfreescientist1
blnScientist5Free = True
backbuffer.SetFillStyle 0
backbuffer.SetFillColor RGB(100, 100, 100)
backbuffer.SetForeColor RGB(255, 250, 70)
intScientist5X = intScientist5X - (xres / 10000)
backbuffer.DrawLine intScientist5X, intScientist5Y, intScientist5X, intScientist5Y - intScientistHeight
If intScientist5X <= 0 And blnTakeOffSoundPlayed = False And blnScientist5OnBoard = True Then
    AddScore (500)
    CheckForLevelComplete (5)
    PlayTakeoffSound
    blnTakeOffSoundPlayed = True
    intScientist5X = 2000
End If
If blnTakeOffSoundPlayed = True Then
    intTakeOffDelay = intTakeOffDelay + 1
    If intTakeOffDelay = 600 Then
        blnScientist5Complete = True
    End If
End If
End Sub

Sub AddScore(intAddScore As Single)                                                             'this function is called whenever score is to be added to
    intScore = intScore + intAddScore                                                           'add the passed score to the total score
End Sub

Sub GenerateMystery()

Randomize Timer                                                                                 'random number used to decide whether mystery appears from left or right
intRandomNumber = Rnd(Timer)
If intRandomNumber > 0.5 Then
    intMysteryX = 0                                                                             'mystery will appear from left of screen
    blnMysteryIncrement = True                                                                  'make sure x position gets incremented
Else
    intMysteryX = xres                                                                          'mystery will appear from right of screen
    blnMysteryIncrement = False                                                                 'make sure x position gets decremented
End If
intMysteryY = intRandomNumber * (yres - (yres - intLandscapeHighestPoint))                      'creates random height for mystery, ensuring its higher than heighest landscape point
PlayItemAppearsSound                                                                            'ping!!
blnMystery = True                                                                               'flag to say mystery is on screen
End Sub

Sub drawMystery()
backbuffer.SetFillStyle 0                                                                       'make sure mystery is filled
backbuffer.SetFillColor RGB(229, 199, 245)
backbuffer.SetForeColor RGB(120, 67, 253)
If blnMysteryIncrement = True Then                                                              'if moving left to right
    intMysteryX = intMysteryX + 3                                                               'move the mystery right
    If intMysteryX >= xres Then                                                                 'check if we've reached edge of screen
        blnMystery = False                                                                      'if we have, stop drawing the mystery
    End If
Else
    intMysteryX = intMysteryX - 3                                                               'move mystery left
    If intMysteryX <= 0 Then                                                                    'check if we've reached the edge of the screen
        blnMystery = False                                                                      'if we have, stop drawing the mystery
    End If
End If
intMysteryY = intMysteryY - 2 * Sin((2 * pi * intMysteryX) / intMysteryWavelength)              'move mystery along a sine wave
backbuffer.DrawCircle intMysteryX, intMysteryY, intCircleWidth                                  'draw mystery capsule
'backbuffer.DrawText intMysteryX - 3, intMysteryY - 8, "?", False                                'draw the '?' in the capsule
End Sub

Sub DanceMode()                                                                                 'party!!!
    PlayDance1Sound                                                                             'pump up the volume!!
    Dim counter As Single
    Dim loopcounter As Integer
    Dim velocityTemp As Single
    Dim velocityXTemp As Single
    Dim scientist1YTemp As Single
    Dim scientist2YTemp As Single
    Dim scientist3YTemp As Single
    Dim scientist4YTemp As Single
    Dim scientist5YTemp As Single
    Dim i As Integer
    
    velocityTemp = velocity                                                                     'remember ships vertical velocity when party started
    velocityXTemp = velocityx                                                                   'remember ships horizontal velocity when party started
    scientist1YTemp = intScientist1Y                                                            'remember each scientists y position when party started
    scientist2YTemp = intScientist2Y
    scientist3YTemp = intScientist3Y
    scientist4YTemp = intScientist4Y
    scientist5YTemp = intScientist5Y
    
    blnDanceMode = True                                                                         'flag to say party is on
    loopcounter = 0
    counter = 0
    velocity = 0                                                                                'stop ships vertical velocity
    velocityx = 0                                                                               'stop ships horizontal velocity
    For i = 1 To 10000                                                                          'delay to try to sync lasers with music
    Next i
    intLaserSourceX = xres / 14.5                                                               'position of laser source x
    intLaserSourceY = yres / 5                                                                  'position of laser source y
    intLaserDest1X = xres / 3                                                                   'position of laser destination x
    intLaserDest1Y = intLandscapeYCollision(intLaserDest1X)                                     'position of laser destination y (this one hits the ground so make sure it stops at ground level)
    Do Until counter = 17
        intLaserStep = xres / 50                                                                'laser will 'jump' by this amount of pixels
        intLaserStep2 = 0 - (xres / 50)
        Do Until loopcounter = 17                                                               'repeat 17 times
            loopcounter = loopcounter + 1
            intScientist1Y = intScientist1Y - 2                                                 'make scientists jump up
            intScientist2Y = intScientist2Y - 2
            intScientist3Y = intScientist3Y - 2
            intScientist4Y = intScientist4Y - 2
            intScientist5Y = intScientist5Y - 2
            blt                                                                                 'draw screen
            DoEvents
        Loop
        intLaserStep = 0 - (xres / 50)
        intLaserStep2 = xres / 50
        Do Until loopcounter = 0                                                                'repeat 17 times
            loopcounter = loopcounter - 1
            intScientist1Y = intScientist1Y + 2                                                 'make scientists just down
            intScientist2Y = intScientist2Y + 2
            intScientist3Y = intScientist3Y + 2
            intScientist4Y = intScientist4Y + 2
            intScientist5Y = intScientist5Y + 2
            blt                                                                                 'draw screen
            DoEvents
        Loop
        counter = counter + 1
    Loop
    intScientist1Y = scientist1YTemp                                                            'restore scientists to their original height
    intScientist2Y = scientist2YTemp
    intScientist3Y = scientist3YTemp
    intScientist4Y = scientist4YTemp
    intScientist5Y = scientist5YTemp
    
    velocity = velocityTemp                                                                     'restore ships y velocity
    velocityx = velocityXTemp                                                                   'restore ships x velocity
    blnDanceMode = False                                                                        'party is over
End Sub

Sub DoubleGravity()                                                                             'double gravity drug has been collected
    
    If intDoubleGravityCount = 0 Then                                                           'double gravity hasn't yet started so...
        blnDoubleGravity = True                                                                 'start it
        gravity = gravity * 2                                                                   'double the gravity
    End If
    If intDoubleGravityCount <= 3000 Then                                                       'if it has started...
        intDoubleGravityCount = intDoubleGravityCount + 1                                       'keep track of how long its been going
    Else                                                                                        'if its time to set it back to normal...
        gravity = gravity / 2                                                                   'half it
        blncentermessage = True                                                                 'and alert the player
        charCenterMessage = "Normal gravity restored"
        intDoubleGravityCount = 0                                                               'reset the counter for the next time
        blnDoubleGravity = False                                                                'turn off double gravity
    End If
    
End Sub

Sub TinyShip()                                                                                  'tiny ship drugs taken
    If intTinyShipCount = 0 Then                                                                'we have just started tiny ship mode
        blnTinyShip = True                                                                      'set the tiny flag
        intShipDistMidBott = intShipDistMidBott / 2                                             'half all the ships dimensions...
        intShipDistMidSide = intShipDistMidSide / 2
        intShipDistMidCent = intShipDistMidCent / 2
        intPlotShipDistMidCent = intPlotShipDistMidCent / 2
        intPlotShipDistMidOuterRing = intPlotShipDistMidOuterRing / 2
        intPlotShipDistMidCore = intPlotShipDistMidCore / 2
        intBeamSizeMax = intBeamSizeMax / 2
        intBeamSizeMin = intBeamSizeMin / 2
        intBeamSize = intBeamSize / 2
        intFlameSizeMax = intFlameSizeMax / 2
        intFlameSizeMin = intFlameSizeMin / 2
        intFlameSize = intFlameSize / 2
    End If
    If intTinyShipCount <= 3000 Then                                                            'control how long the ship stays tiny
        intTinyShipCount = intTinyShipCount + 1
    Else                                                                                        'ship has been tiny for long enough
        If blnShipLanded = False Then                                                           'if it isn't landed...
            intShipDistMidBott = intShipDistMidBott * 2                                         'restore normal dimensions
            intShipDistMidSide = intShipDistMidSide * 2
            intShipDistMidCent = intShipDistMidCent * 2
            intPlotShipDistMidCent = intPlotShipDistMidCent * 2
            intPlotShipDistMidOuterRing = intPlotShipDistMidOuterRing * 2
            intPlotShipDistMidCore = intPlotShipDistMidCore * 2
            intBeamSizeMax = intBeamSizeMax * 2
            intBeamSizeMin = intBeamSizeMin * 2
            intBeamSize = intBeamSize * 2
            intFlameSizeMax = intFlameSizeMax * 2
            intFlameSizeMin = intFlameSizeMin * 2
            intFlameSize = intFlameSize * 2
            blncentermessage = True
            charCenterMessage = "Ship size restored"                                            'alert the player of size restore
            intTinyShipCount = 0                                                                'reset the tiny timer for next time
            blnTinyShip = False                                                                 'tiny flag off
        Else
            intTinyShipCount = 2000                                                             'step the timer back. Ship changing size while landed causes explosion!
        End If
    End If
End Sub

Sub InitialiseShip()
    intFuelLevel = 2000                                                                         'set initial fuel level
    intShipMiddleX = xres / 2                                                                   'set initial position of ship
    intShipMiddleY = yres / 4
    stepval = 360                                                                               'initial bearing of the ship
    velocity = -1                                                                               'initial velocity - make ship travel slightly upwards at start of game to allow player time to react
    velocityx = 0                                                                               'before gravity kicks in and ship begins to fall
    leftpressed = False                                                                         'initialise flags as if no keys are pressed
    rightpressed = False
    thrustpressed = False
    beampressed = False
    shootpressed = False
End Sub

Sub DrawGame()
If Not blnCollided Then                                                                         'if the ship hasnt crashed, set its colour to white
    backbuffer.SetForeColor RGB(255, 255, 255)
End If
backbuffer.DrawLine intPlotShipTopX, intPlotShipTopY, intPlotShipLeftX, intPlotShipLeftY        'draw the ship
backbuffer.DrawLine intPlotShipLeftX, intPlotShipLeftY, intPlotShipCenterX, intPlotShipCenterY
backbuffer.DrawLine intPlotShipCenterX, intPlotShipCenterY, intPlotShipRightX, intPlotShipRightY
backbuffer.DrawLine intPlotShipRightX, intPlotShipRightY, intPlotShipTopX, intPlotShipTopY
If thrustpressed = True Then                                                                    'if thrust is being applied, draw the flames
    If flameshown >= 2 And blnCollided = False Then                                             'flameshown is a counter from 0-2. Flames are only drawn when flameshown = 2 to imitate flicker
        backbuffer.SetForeColor RGB(297, 194, 131)                                              'inner flame
        backbuffer.DrawLine intPlotCoreLeftX, intPlotCoreLeftY, intPlotCoreRightX, intPlotCoreRightY
        backbuffer.DrawLine intPlotCoreLeftX, intPlotCoreLeftY, intPlotCoreCenterX, intPlotCoreCenterY
        backbuffer.DrawLine intPlotCoreRightX, intPlotCoreRightY, intPlotCoreCenterX, intPlotCoreCenterY
        backbuffer.SetForeColor RGB(187, 59, 3)                                                 'outer flame
        backbuffer.DrawLine intPlotCoreLeftX, intPlotCoreLeftY, intPlotFlameCenterX, intPlotFlameCenterY
        backbuffer.DrawLine intPlotCoreRightX, intPlotCoreRightY, intPlotFlameCenterX, intPlotFlameCenterY
        intFlameSize = intFlameSize + 0.4                                                       'increase the size of the flame the longer its lit
        flameshown = 0                                                                          'reset the flicker counter
    Else
        flameshown = flameshown + 1                                                             'increment the flicker counter
    End If
End If

If blnShipLanded = False Then                                                                   'if ship is airborne
    If beampressed = True Then                                                                  'if tractor beam is being applied, draw the beam
        If intBeamShown >= 2 And blnCollided = False Then                                       'beamshown is a counter from 0-2. Beam is only drawn when beamshown = 2 to imitate flicker
            backbuffer.SetForeColor RGB(150, 150, 200)                                          'beamcolour
            intBeamBottLeftX = intShipMiddleX + Sin((stepval - 30) * Rad) * intBeamSize         'determine beam coordinates
            intBeamBottLeftY = intShipMiddleY + Cos((stepval - 30) * Rad) * intBeamSize
            intBeamBottRightX = intShipMiddleX + Sin((stepval + 30) * Rad) * intBeamSize
            intBeamBottRightY = intShipMiddleY + Cos((stepval + 30) * Rad) * intBeamSize
            intBeamSize = intBeamSize + 0.8                                                     'make beam grow the longer its on
            If intBeamSize > intBeamSizeMax Then                                                'make sure beam doesnt get too big
                intBeamSize = intBeamSizeMax
            End If
                                                                                                'draw the beam
            backbuffer.DrawLine intPlotShipLeftX, intPlotShipLeftY, intBeamBottLeftX, intBeamBottLeftY
            backbuffer.DrawLine intPlotShipRightX, intPlotShipRightY, intBeamBottRightX, intBeamBottRightY
            intBeamShown = 0                                                                    'reset the flicker counter
        Else
            intBeamShown = intBeamShown + 1                                                     'increment the flicker counter
        End If
    End If
End If
                                                                                                'set landscape colour
If blnDanceMode = False Then                                                                    'if theres no party on...
    backbuffer.SetForeColor RGB(175, 100, 100)                                                  'set normal landscape colour
Else                                                                                            'otherwise, party is on
    backbuffer.SetForeColor RGB(intColourCycle / 1.3, intColourCycle2 / 2, intColourCycle2 / 3) 'cycle the landscape colour
End If

intLandscapePositionX = 0                                                                       'draw landscape
For intLandscapeCount = 0 To (intTerrainComplexity - 1) Step 1
    backbuffer.DrawLine intLandscapePositionX, intLandscapeY(intLandscapeCount), intLandscapePositionX + intLandscapeXIncrement, intLandscapeY(intLandscapeCount + 1)
    intLandscapePositionX = intLandscapePositionX + intLandscapeXIncrement
Next intLandscapeCount

backbuffer.SetForeColor RGB(40, 40, 40)                                                         'gate colours
backbuffer.SetFillColor RGB(50, 50, 50)
backbuffer.DrawBox intGateTopLeftx, intGateTopLefty, intGateBottomRightx, intGateBottomRighty   'draw gate

drawSwitch                                                                                      'draw switch that opens the lab gates
                                                                                                'draw fuel powerups
If blnFuel = False Then                                                                         'is fuel on screen?
    If intFuelLevel < intRefuelLevel Then                                                       'is player getting low on fuel?
        GenerateFuel                                                                            'set up coordinates to draw fuel
    End If
Else                                                                                            'fuel either is, or should be, on screen
    drawFuel                                                                                    'so draw it
End If

                                                                                                
If intLevel > 1 Then                                                                                                'draw turret
    If blnTurret = False Then                                                                       'is turret on screen?
        GenerateTurret                                                                              'set up coordinates to draw turret
    Else                                                                                            'turret either is, or should be, on screen
        If blnTurretDestroyed = False Then                                                          'if turret hasn't already been destroyed...
            DrawTurret                                                                              'draw it
            If intGeneralCounter = 1000 Or _
            intGeneralCounter = 2000 Or _
            intGeneralCounter = 3000 Or _
            intGeneralCounter = 4000 Or _
            intGeneralCounter = 5000 Then                                                           'set periods when turret will shoot
                If intLives > 0 Then                                                                'probs occur when turrets keep shooting after game over
                    If blnCollided = False Then
                        generateTurretShot
                    End If
                End If
            End If
        End If
    End If
End If

If intLevel > 2 Then
    If blnTurret2 = False Then                                                                       'is turret2 on screen?
        GenerateTurret2                                                                              'set up coordinates to draw turret
    Else                                                                                            'turret either is, or should be, on screen
        If blnTurret2Destroyed = False Then                                                          'if turret2 hasn't already been destroyed...
            DrawTurret2                                                                              'draw it
            If intGeneralCounter = 500 Or _
            intGeneralCounter = 1500 Or _
            intGeneralCounter = 2500 Or _
            intGeneralCounter = 3500 Or _
            intGeneralCounter = 4500 Then                                                           'set periods when turret2 will shoot
                If intLives > 0 Then                                                                'probs occur when turrets keep shooting after game over
                    If blnCollided = False Then
                        generateTurret2Shot
                    End If
                End If
            End If
        End If
    End If
End If

If intLevel > 3 Then
    If blnTurret3 = False Then                                                                       'is turret3 on screen?
        GenerateTurret3                                                                              'set up coordinates to draw turret
    Else                                                                                            'turret either is, or should be, on screen
        If blnTurret3Destroyed = False Then                                                          'if turret3 hasn't already been destroyed...
            DrawTurret3                                                                              'draw it
            If intGeneralCounter = 750 Or _
            intGeneralCounter = 1350 Or _
            intGeneralCounter = 2700 Or _
            intGeneralCounter = 3200 Or _
            intGeneralCounter = 4700 Then                                                           'set periods when turret3 will shoot
                If intLives > 0 Then                                                                'probs occur when turrets keep shooting after game over
                    If blnCollided = False Then
                        generateTurret3Shot
                    End If
                End If
            End If
        End If
    End If
End If

If intLevel > 4 Then
    If blnTurret4 = False Then                                                                       'is turret4 on screen?
        GenerateTurret4                                                                              'set up coordinates to draw turret
    Else                                                                                            'turret either is, or should be, on screen
        If blnTurret4Destroyed = False Then                                                          'if turret4 hasn't already been destroyed...
            DrawTurret4                                                                              'draw it
            If intGeneralCounter = 750 Or _
            intGeneralCounter = 1350 Or _
            intGeneralCounter = 2700 Or _
            intGeneralCounter = 3200 Or _
            intGeneralCounter = 4700 Then                                                           'set periods when turret4 will shoot
                If intLives > 0 Then                                                                'probs occur when turrets keep shooting after game over
                    If blnCollided = False Then
                        generateTurret4Shot
                    End If
                End If
            End If
        End If
    End If
End If

If intLevel > 5 Then
    If blnTurret5 = False Then                                                                       'is turret5 on screen?
        GenerateTurret5                                                                              'set up coordinates to draw turret
    Else                                                                                            'turret either is, or should be, on screen
        If blnTurret5Destroyed = False Then                                                          'if turret5 hasn't already been destroyed...
            DrawTurret5                                                                              'draw it
            If intGeneralCounter = 750 Or _
            intGeneralCounter = 1350 Or _
            intGeneralCounter = 2700 Or _
            intGeneralCounter = 3200 Or _
            intGeneralCounter = 4700 Then                                                           'set periods when turret5 will shoot
                If intLives > 0 Then                                                                'probs occur when turrets keep shooting after game over
                    If blnCollided = False Then
                        generateTurret5Shot
                    End If
                End If
            End If
        End If
    End If
End If

If intLevel > 6 Then
    If blnTurret6 = False Then                                                                       'is turret6 on screen?
        Generateturret6                                                                              'set up coordinates to draw turret
    Else                                                                                            'turret either is, or should be, on screen
        If blnTurret6Destroyed = False Then                                                          'if turret6 hasn't already been destroyed...
            Drawturret6                                                                              'draw it
            If intGeneralCounter = 500 Or _
            intGeneralCounter = 1500 Or _
            intGeneralCounter = 2500 Or _
            intGeneralCounter = 3500 Or _
            intGeneralCounter = 4500 Then                                                           'set periods when turret6 will shoot
                If intLives > 0 Then                                                                'probs occur when turrets keep shooting after game over
                    If blnCollided = False Then
                        generateturret6Shot
                    End If
                End If
            End If
        End If
    End If
End If

If intLevel > 7 Then
    If blnTurret7 = False Then                                                                       'is turret7 on screen?
        Generateturret7                                                                              'set up coordinates to draw turret
    Else                                                                                            'turret either is, or should be, on screen
        If blnTurret7Destroyed = False Then                                                          'if turret7 hasn't already been destroyed...
            Drawturret7                                                                              'draw it
            If intGeneralCounter = 500 Or _
            intGeneralCounter = 1500 Or _
            intGeneralCounter = 2500 Or _
            intGeneralCounter = 3500 Or _
            intGeneralCounter = 4500 Then                                                           'set periods when turret7 will shoot
                If intLives > 0 Then                                                                'probs occur when turrets keep shooting after game over
                    If blnCollided = False Then
                        generateturret7Shot
                    End If
                End If
            End If
        End If
    End If
End If

If blnCollided = False Then
    If blnTurretShot = True Then                                                                     'the turret has fired a shot which is still on screen, so draw it
        drawTurretShot
    End If
    
    If blnTurret2Shot = True Then                                                                     'the turret has fired a shot which is still on screen, so draw it
        drawTurret2Shot
    End If
    
    If blnTurret3Shot = True Then                                                                     'the turret has fired a shot which is still on screen, so draw it
        drawTurret3Shot
    End If
    
    If blnTurret4Shot = True Then                                                                     'the turret has fired a shot which is still on screen, so draw it
        drawTurret4Shot
    End If
    
    If blnTurret5Shot = True Then                                                                     'the turret has fired a shot which is still on screen, so draw it
        drawTurret5Shot
    End If
    
    If blnTurret6Shot = True Then                                                                     'the turret has fired a shot which is still on screen, so draw it
        drawturret6Shot
    End If
    
    If blnTurret7Shot = True Then                                                                     'the turret has fired a shot which is still on screen, so draw it
        drawturret7Shot
    End If
End If

intGeneralCounter = intGeneralCounter + 1                                                       'used as a trigger to kick off various things (eg appearance of capsules)
If intGeneralCounter >= 5000 Then                                                               'the counter keeps looping from 1 to 5000
    intGeneralCounter = 0
End If
                                                                                                'draw drugs powerups
If blnDrugs = False Then                                                                        'are drugs on screen?
    If intGeneralCounter = 1600 Then                                                            'creates periods where there are no drugs on screen
        GenerateDrugs                                                                           'set up coordinates to draw drugs
    End If
Else                                                                                            'drugs are either on screen, or should be
    drawDrugs                                                                                   'so draw them
End If
                                                                                                'draw mystery powerups
If blnMystery = False Then                                                                      'are drugs on screen?
    If intGeneralCounter = 3200 Then                                                            'creates periods where there is no mystery on screen
        GenerateMystery                                                                         'set up coordinates to draw mystery
    End If
Else                                                                                            'mystery is either on screen, or should be
    drawMystery                                                                                 'so draw it
End If
                                                                                                'draw anti-gravity powerups
If blnAntiGravityShown = False Then                                                             'is antigravity powerup on screen?
    If intGeneralCounter = 4800 Then                                                            'creates period where antigrav wont be on screen
        GenerateAntiGravity                                                                     'set up coordinates to draw antigrav powerup
    End If
Else                                                                                            'antigrav is either on screen, or should be
    drawAntiGravity                                                                             'so draw it
End If
                                                                                                'draw scientists
If blnScientist1 = False Then
    If intTakeOffDelay >= 600 Then
        blnShipLanded = False                                                                   'provides a trigger for the ship to take off once the scientist reaches the edge of screen
        intTakeOffDelay = 0
        blnTakeOffSoundPlayed = False
        blnScientist1OnBoard = False
    End If
    If blnShipLanded = True Then
        If blnScientist1OnBoard = True And blnScientist1Complete = False Then
            DrawFreeScientist1                                                                  'the scientist must have been rescued
        End If
    Else
        If intGeneralCounter = 200 And blnScientist1Free = False And blnScientist1OnBoard = False Then
            GenerateScientist1                                                                  'create a new scientist
        End If
    End If
Else
    DrawScientist1                                                                              'move the scientist
End If

If blnScientist2 = False Then
    If intTakeOffDelay >= 600 Then
        blnShipLanded = False                                                                   'provides a trigger for the ship to take off once the scientist reaches the edge of screen
        intTakeOffDelay = 0
        blnTakeOffSoundPlayed = False
        blnScientist2OnBoard = False
    End If
    If blnShipLanded = True Then
        If blnScientist2OnBoard = True And blnScientist2Complete = False Then
            DrawFreeScientist2                                                                  'the scientist must have been rescued
        End If
    Else
        If intGeneralCounter = 500 And blnScientist2Free = False And blnScientist2OnBoard = False Then
            GenerateScientist2                                                                  'create a new scientist
        End If
    End If
Else
    DrawScientist2                                                                              'move the scientist
End If

If blnScientist3 = False Then
    If intTakeOffDelay >= 600 Then
        blnShipLanded = False                                                                   'provides a trigger for the ship to take off once the scientist reaches the edge of screen
        intTakeOffDelay = 0
        blnTakeOffSoundPlayed = False
        blnScientist3OnBoard = False
    End If
    If blnShipLanded = True Then
        If blnScientist3OnBoard = True And blnScientist3Complete = False Then
            DrawFreeScientist3                                                                  'the scientist must have been rescued
        End If
    Else
        If intGeneralCounter = 650 And blnScientist3Free = False And blnScientist3OnBoard = False Then
            GenerateScientist3                                                                  'create a new scientist
        End If
    End If
Else
    DrawScientist3                                                                              'move the scientist
End If

If blnScientist4 = False Then
    If intTakeOffDelay >= 600 Then
        blnShipLanded = False                                                                   'provides a trigger for the ship to take off once the scientist reaches the edge of screen
        intTakeOffDelay = 0
        blnTakeOffSoundPlayed = False
        blnScientist4OnBoard = False
    End If
    If blnShipLanded = True Then
        If blnScientist4OnBoard = True And blnScientist4Complete = False Then
            DrawFreeScientist4                                                                  'the scientist must have been rescued
        End If
    Else
        If intGeneralCounter = 900 And blnScientist4Free = False And blnScientist4OnBoard = False Then
            GenerateScientist4                                                                  'create a new scientist
        End If
    End If
Else
    DrawScientist4                                                                              'move the scientist
End If

If blnScientist5 = False Then
    If intTakeOffDelay >= 600 Then
        blnShipLanded = False                                                                   'provides a trigger for the ship to take off once the scientist reaches the edge of screen
        intTakeOffDelay = 0
        blnTakeOffSoundPlayed = False
        blnScientist5OnBoard = False
    End If
    If blnShipLanded = True Then
        If blnScientist5OnBoard = True And blnScientist5Complete = False Then
            DrawFreeScientist5                                                                  'the scientist must have been rescued
        End If
    Else
        If intGeneralCounter = 1200 And blnScientist5Free = False And blnScientist5OnBoard = False Then
            GenerateScientist5                                                                  'create a new scientist
        End If
    End If
Else
    DrawScientist5                                                                              'move the scientist
End If

If blnDanceMode = True Then                                                                     'do lasers need to be drawn?
    intLaserDest1X = intLaserDest1X + intLaserStep                                              'move the laser
    intLaserDest1Y = intLandscapeYCollision(intLaserDest1X)
    backbuffer.SetForeColor RGB(150, 0, 0)                                                      '3 red lasers getting progressively darker
    backbuffer.DrawLine intLaserSourceX, intLaserSourceY, intLaserDest1X, intLaserDest1Y
    backbuffer.SetForeColor RGB(75, 0, 0)
    backbuffer.DrawLine intLaserSourceX, intLaserSourceY, intLaserDest1X + (intLaserStep * 2), intLandscapeYCollision(intLaserDest1X + (intLaserStep * 2))
    backbuffer.SetForeColor RGB(20, 0, 0)
    backbuffer.DrawLine intLaserSourceX, intLaserSourceY, intLaserDest1X + (intLaserStep * 3), intLandscapeYCollision(intLaserDest1X + (intLaserStep * 3))
    
    backbuffer.SetForeColor RGB(0, 100, 0)                                                      '3 green lasers getting progressively darker
    backbuffer.DrawLine xres, yres / 2, intLaserDest1X + intLaserDest1X / 3, intLandscapeYCollision(intLaserDest1X + (intLaserDest1X / 3))
    backbuffer.SetForeColor RGB(0, 50, 0)
    backbuffer.DrawLine xres, yres / 2, intLaserDest1X + intLaserDest1X / 3 + (intLaserStep2 * 1.5), intLandscapeYCollision(intLaserDest1X + (intLaserDest1X / 3))
    backbuffer.SetForeColor RGB(0, 20, 0)
    backbuffer.DrawLine xres, yres / 2, intLaserDest1X + intLaserDest1X / 3 + (intLaserStep2 * 2), intLandscapeYCollision(intLaserDest1X + (intLaserDest1X / 3))
End If
    
If shootpressed = True Then                                                                     'check to see if ships shots need to be drawn
    If blnShotFired = False Then                                                                'if this is the first time this shot needs drawing
        StopShootSound                                                                          'in case its still playing from a previous shot
        PlayShootSound                                                                          'blam!!
        intShootSourceX = intPlotShipTopX                                                       'source of shot coordinates (front of ship)
        intShootSourceY = intPlotShipTopY
        ShotAngle = stepval                                                                     'angle at which shot will travel (same as ship was when fired)
        intShotRange = 15
        blnShotFired = True                                                                     'flag to say the shot has been fired and will now need drawing
    Else                                                                                        'the shot has already been fired, so draw it
        intShotX1 = intShootSourceX + Sin((ShotAngle - 180) * Rad) * (intShotRange + 8)         'plot the coordinates of the shot
        intShotY1 = intShootSourceY + Cos((ShotAngle - 180) * Rad) * (intShotRange + 8)
        intShotX2 = intShootSourceX + Sin((ShotAngle - 180) * Rad) * (intShotRange)
        intShotY2 = intShootSourceY + Cos((ShotAngle - 180) * Rad) * (intShotRange)
        intShotRange = intShotRange + 10                                                        'increment the shot range (to make shot travel!)
        backbuffer.SetForeColor RGB(237, 255, 33)                                               'shot colour
        backbuffer.DrawLine intShotX1, intShotY1, intShotX2, intShotY2                          'draw the shot
        If intShotX1 >= xres Then                                                               'has the shot hit the right of the screen?
            shootpressed = False                                                                'stop the life of this shot
            blnShotFired = False
        End If
        If intShotX1 <= 0 Then                                                                  'has the shot hit the left of the screen?
            shootpressed = False                                                                'stop the life of this shot
            blnShotFired = False
        End If
        If intShotY1 <= 0 Then                                                                  'has the shot hit the top of the screen?
            shootpressed = False                                                                'stop the life of this shot
            blnShotFired = False
        End If
        If intShotY1 >= intLandscapeYCollision(intShotX1) Then                                  'has the shot hit the landscape?
            shootpressed = False                                                                'stop the life of this shot
            blnShotFired = False
        End If
        If intShotX1 > (intSwitchX - 10) And intShotX1 < (intSwitchX + 10) Then                 'has the shot hit the switch?
            If intShotY1 > (intSwitchButtonTopY - 10) And intShotY1 < (intSwitchY + 10) Then
                If blnSwitchPressed = False Then                                                'if the switch hasnt already been pressed
                    If blnAnyScientistOnBoard = True Then                                       'and if a scientist is on board
                        blnSwitchPressed = True                                                 'set switch pressed flag to true
                        PlayGateOpenSound                                                       'play sound
                        AddScore (20)                                                           'award some points
                    Else                                                                        'no scientist on board
                        PlayAccessDeniedSound                                                   'so dont allow gates to open
                    End If
                End If
                shootpressed = False                                                            'stop the life of this shot
                blnShotFired = False
            End If
        End If
        If intShotX1 > (intTurretCircleX - 10) And intShotX1 < (intTurretCircleX + 10) Then     'has shot hit the turret?
            If intShotY1 > (intTurretCircleY - 10) And intShotY1 < (intTurretCircleY + 10) Then
                If blnTurretDestroyed = False Then                                              'if turret hasnt already been destroyed
                    blnTurretDestroyed = True                                                   'destroy it
                    PlayTurretDestroyedSound                                                    'boooom!
                    AddScore (50)                                                               'award some points
                    shootpressed = False                                                        'stop the life of this shot
                    blnShotFired = False
                End If
            End If
        End If
        If intShotX1 > (intTurret2CircleX - 10) And intShotX1 < (intTurret2CircleX + 10) Then     'has shot hit the turret2?
            If intShotY1 > (intTurret2CircleY - 10) And intShotY1 < (intTurret2CircleY + 10) Then
                If blnTurret2Destroyed = False Then                                              'if turret2 hasnt already been destroyed
                    blnTurret2Destroyed = True                                                   'destroy it
                    PlayTurretDestroyedSound                                                    'boooom!
                    AddScore (50)                                                               'award some points
                    shootpressed = False                                                        'stop the life of this shot
                    blnShotFired = False
                End If
            End If
        End If
        If intShotX1 > (intTurret3CircleX - 10) And intShotX1 < (intTurret3CircleX + 10) Then     'has shot hit the turret3?
            If intShotY1 > (intTurret3CircleY - 10) And intShotY1 < (intTurret3CircleY + 10) Then
                If blnTurret3Destroyed = False Then                                              'if turret3 hasnt already been destroyed
                    blnTurret3Destroyed = True                                                   'destroy it
                    PlayTurretDestroyedSound                                                    'boooom!
                    AddScore (50)                                                               'award some points
                    shootpressed = False                                                        'stop the life of this shot
                    blnShotFired = False
                End If
            End If
        End If
        If intShotX1 > (intTurret4CircleX - 10) And intShotX1 < (intTurret4CircleX + 10) Then     'has shot hit the turret4?
            If intShotY1 > (intTurret4CircleY - 10) And intShotY1 < (intTurret4CircleY + 10) Then
                If blnTurret4Destroyed = False Then                                              'if turret4 hasnt already been destroyed
                    blnTurret4Destroyed = True                                                   'destroy it
                    PlayTurretDestroyedSound                                                    'boooom!
                    AddScore (50)                                                               'award some points
                    shootpressed = False                                                        'stop the life of this shot
                    blnShotFired = False
                End If
            End If
        End If
        If intShotX1 > (intTurret5CircleX - 10) And intShotX1 < (intTurret5CircleX + 10) Then     'has shot hit the turret5?
            If intShotY1 > (intTurret5CircleY - 10) And intShotY1 < (intTurret5CircleY + 10) Then
                If blnTurret5Destroyed = False Then                                              'if turret5 hasnt already been destroyed
                    blnTurret5Destroyed = True                                                   'destroy it
                    PlayTurretDestroyedSound                                                    'boooom!
                    AddScore (50)                                                               'award some points
                    shootpressed = False                                                        'stop the life of this shot
                    blnShotFired = False
                End If
            End If
        End If
        If intShotX1 > (intTurret6CircleX - 10) And intShotX1 < (intTurret6CircleX + 10) Then     'has shot hit the turret6?
            If intShotY1 > (intTurret6CircleY - 10) And intShotY1 < (intTurret6CircleY + 10) Then
                If blnTurret6Destroyed = False Then                                              'if turret6 hasnt already been destroyed
                    blnTurret6Destroyed = True                                                   'destroy it
                    PlayTurretDestroyedSound                                                    'boooom!
                    AddScore (50)                                                               'award some points
                    shootpressed = False                                                        'stop the life of this shot
                    blnShotFired = False
                End If
            End If
        End If
        If intShotX1 > (intTurret7CircleX - 10) And intShotX1 < (intTurret7CircleX + 10) Then     'has shot hit the turret7?
            If intShotY1 > (intTurret7CircleY - 10) And intShotY1 < (intTurret7CircleY + 10) Then
                If blnTurret7Destroyed = False Then                                              'if turret7 hasnt already been destroyed
                    blnTurret7Destroyed = True                                                   'destroy it
                    PlayTurretDestroyedSound                                                    'boooom!
                    AddScore (50)                                                               'award some points
                    shootpressed = False                                                        'stop the life of this shot
                    blnShotFired = False
                End If
            End If
        End If
    End If
End If

fontinfo.Size = 8                                                                               'font size for the statistics
backbuffer.SetFont fontinfo

If statspressed = True Then                                                                     'if player has chosen to view statistics
    backbuffer.SetForeColor RGB(0, 200, 0)                                                      'set foreground colour
    backbuffer.DrawText intStatsXCoordinate, 10, "X Velocity: " & velocityx, False              'output X velocity
    backbuffer.DrawText intStatsXCoordinate, 25, "Y Velocity: " & velocity, False               'output Y velocity
    backbuffer.DrawText intStatsXCoordinate, 40, "Altitude: " & (yres - intShipMiddleY), False  'output altitude (currently y coordinate - will need changing to something more suitable
    If velocity < 0 Then                                                                        'output speed
        singSpeed = velocity - (2 * velocity)                                                   'velocity is negative so make the value positive for speed
    Else
        singSpeed = velocity
    End If
    backbuffer.DrawText intStatsXCoordinate, 55, "Speed: " & singSpeed, False
    If stepval <> 360 Then                                                                      'output bearing - make sure that up (or north) always reads 0 and not 360
        backbuffer.DrawText intStatsXCoordinate, 70, "Bearing: " & stepval, False
    Else
        backbuffer.DrawText intStatsXCoordinate, 70, "Bearing: 0", False
    End If
                                                                                                'and some other statistics...
    backbuffer.DrawText intStatsXCoordinate, 85, "Rotation Rate: " & intRotationRate, False
    backbuffer.DrawText intStatsXCoordinate, 100, "Terrain variance: " & intTerrainVariance, False
    backbuffer.DrawText intStatsXCoordinate, 115, "Terrain complexity: " & intTerrainComplexity, False
    backbuffer.DrawText intStatsXCoordinate, 130, "Terrain highest peak: " & (yres - intLandscapeHighestPoint), False
    backbuffer.DrawText intStatsXCoordinate, 145, "Terrain lowest point: " & (yres - intLandscapeLowestPoint), False
    backbuffer.DrawText intStatsXCoordinate, 160, "X coordinate: " & intShipMiddleX, False
    backbuffer.DrawText intStatsXCoordinate, 175, "Y coordinate: " & intShipMiddleY, False
End If

If blnLegendPressed = True Then
    backbuffer.SetFillStyle 0                                                                   'to make sure the drug capsule is filled
    backbuffer.SetFillColor RGB(100, 150, 100)                                                  'colour of drug capsule
    backbuffer.SetForeColor RGB(0, 70, 0)                                                       'colour of outline
    backbuffer.DrawCircle intStatsXCoordinate + (intCircleWidth / 2), 200, intCircleWidth
    backbuffer.SetForeColor RGB(0, 200, 0)
    backbuffer.DrawText intStatsXCoordinate + (intCircleWidth * 1.5), 200 - (intCircleWidth), "    Drugs", False
    backbuffer.SetFillStyle 0                                                                   'make sure mystery is filled
    backbuffer.SetFillColor RGB(229, 199, 245)
    backbuffer.SetForeColor RGB(120, 67, 253)
    backbuffer.DrawCircle intStatsXCoordinate + (intCircleWidth / 2), 220, intCircleWidth
    backbuffer.SetForeColor RGB(0, 200, 0)
    backbuffer.DrawText intStatsXCoordinate + (intCircleWidth * 1.5), 220 - (intCircleWidth), "    Mystery", False
    backbuffer.SetFillStyle 0                                                                   'make sure antigrav will be filled
    backbuffer.SetFillColor RGB(150, 100, 100)                                                  'fill colour
    backbuffer.SetForeColor RGB(70, 0, 0)                                                       'draw colour
    backbuffer.DrawCircle intStatsXCoordinate + (intCircleWidth / 2), 240, intCircleWidth
    backbuffer.SetForeColor RGB(0, 200, 0)
    backbuffer.DrawText intStatsXCoordinate + (intCircleWidth * 1.5), 240 - (intCircleWidth), "    Gravity", False
    backbuffer.SetFillStyle 0                                                                   'make the fuel capsule filled
    backbuffer.SetFillColor RGB(100, 100, 150)                                                  'with this colour
    backbuffer.SetForeColor RGB(0, 0, 100)                                                      'text (F) colour
    backbuffer.DrawCircle intStatsXCoordinate + (intCircleWidth / 2), 260, intCircleWidth
    backbuffer.SetForeColor RGB(0, 200, 0)
    backbuffer.DrawText intStatsXCoordinate + (intCircleWidth * 1.5), 260 - (intCircleWidth), "    Extra Fuel", False
End If

fontinfo.Size = 10                                                                              'increase font size for more important stats, etc
backbuffer.SetFont fontinfo
                                                                                                'bottom left of screen
backbuffer.SetForeColor RGB(50, 150, 50)
If blnAnyScientistOnBoard = True Then                                                           'if a scientist is on board
    backbuffer.DrawText 10, yres - 20, "Objective: Return the scientist to the laboratory", False   'state the objective of freeing him
Else                                                                                            'no scientist on board
    If intScientistsOutCount = 0 Then                                                           'if no scientists on screen
        backbuffer.DrawText 10, yres - 20, "Objective: Wait for scientists to appear", False    'objective is to wait
    Else                                                                                        'otherwise objective is to rescue one
        backbuffer.DrawText 10, yres - 20, "Objective: Rescue 1 of the " & intScientistsOutCount & " remaining scientists", False
    End If
End If

Select Case intDrugType                                                                         'write drug effects
Case 1
    backbuffer.DrawText 10, yres - 50, "Drug effects: Loss of orientation", False
Case 2
    backbuffer.DrawText 10, yres - 50, "Drug effects: Accelerated response to ships controls", False
Case 3
    backbuffer.DrawText 10, yres - 50, "Drug effects: None", False
End Select

Select Case blnAntiGravity                                                                      'write gravity state
Case True
    backbuffer.DrawText 10, yres - 65, "Gravity: Reversed", False
Case False
    If intDoubleGravityCount = 0 Then
        backbuffer.DrawText 10, yres - 65, "Gravity: Normal", False
    Else
        backbuffer.DrawText 10, yres - 65, "Gravity: Double", False
    End If
End Select

backbuffer.DrawText 10, yres - 35, "Scientists at risk: " & intScientistsOutCount, False        'count of number of scientists on the ground

fontinfo.Size = 14                                                                              'even bigger font for score, fuel, etc
backbuffer.SetFont fontinfo

backbuffer.DrawText xres / 2, yres - 25, "Score: " & intScore, False                            'bottom middle of screen

backbuffer.DrawText intStatsXCoordinate - 15, yres - 65, "Level: " & intLevel, False            'bottom right of screen
backbuffer.DrawText intStatsXCoordinate - 15, yres - 45, "Lives: " & intLives, False

If intFuelLevel > intRefuelLevel Then                                                           'if player has loads of fuel
    backbuffer.SetForeColor RGB(50, 150, 50)                                                    'write fuel level in green
Else
    If intFuelLevel < 500 Then                                                                  'if very little fuel
        backbuffer.SetForeColor RGB(175, 0, 0)                                                  'write fuel level in red
    Else                                                                                        'if time to refuel
        backbuffer.SetForeColor RGB(247, 247, 105)                                              'write fuel in yellow
    End If
End If
backbuffer.DrawText intStatsXCoordinate - 15, yres - 25, "Fuel level: " & intFuelLevel, False

If blncentermessage = True Then                                                                 'output messages to center of screen
    If intcentermessagetime < 100 Then                                                          'used to keep message visible for a fixed period
        intColourCycle = intColourCycle + 5                                                     'cycle the message colour
        If intColourCycle >= 255 Then                                                           'dont exceed RGB max of 255
            intColourCycle = 20
        End If
        backbuffer.SetForeColor RGB(intColourCycle, intColourCycle2, intColourCycle2)           'text colour
        backbuffer.DrawText xres / 2.2, yres / 2.3, charCenterMessage, False                    'output the message
        intcentermessagetime = intcentermessagetime + 1                                         'keep track of how long message has been displayed
    Else                                                                                        'message has been up for long enough so
        blncentermessage = False                                                                'stop showing it
        intcentermessagetime = 0                                                                'reset the timer
    End If
End If

If blnDanceMode = True Then                                                                     'if dance mode is on, display text in middle of screen
    intColourCycle = intColourCycle + 1                                                         'cycle the colours of the center message and landscape
    intColourCycle2 = intColourCycle2 + 3
    intColourCycle3 = intColourCycle3 + 2
    If intColourCycle >= 255 Then
        intColourCycle = 0
    End If
    If intColourCycle2 >= 255 Then
        intColourCycle2 = 0
    End If
    If intColourCycle3 >= 255 Then
        intColourCycle3 = 0
    End If
    backbuffer.SetForeColor RGB(intColourCycle, intColourCycle2, intColourCycle2)
    backbuffer.DrawText xres / 2.3, yres / 2.3, "Intermission... Let's party :)", False
    backbuffer.DrawText xres / 2.1, yres / 2.2, "+ 500 points", False
End If

fontinfo.Size = 10                                                                              'restore the font size ready for use on drug capsules, switch, etc
backbuffer.SetFont fontinfo

End Sub

Sub GenerateLandscape()
intTerrainVariance = yres / 6                                                                   'used to determine how rough the landscape is
intLowerLandscape = ((yres / 4) * 3) - intTerrainVariance / 2                                   'calculate 3/4's of the height of the users screen. To make sure most of the terrain is not off the bottom of screen, raise this level by half of the terrain variance.The landscape will be placed at this level
intTerrainComplexity = 16                                                                       'used to determine complexity of landscape (how many lines its made of)
                                                                                                'generate all the y coordinates to make a landscape from. The number created depends on the complexity of the landscape
intLandscapeHighestPoint = 10000                                                                'set silly levels to start with for max and min heights
intLandscapeLowestPoint = 0                                                                     'this COULD be raised as the levels get harder
For intLandscapeCount = 0 To (intTerrainComplexity - 1) Step 1                                  'go through and generate each peak and trough
    intLandscapeY(intLandscapeCount) = Rnd * intTerrainVariance + intLowerLandscape             'set a random height for the peak/trough
    If intLandscapeY(intLandscapeCount) < intLandscapeHighestPoint Then                         'if this is the heighest peak so far...
        intLandscapeHighestPoint = intLandscapeY(intLandscapeCount)                             'record it
    End If
    If intLandscapeY(intLandscapeCount) > intLandscapeLowestPoint Then                          'if this is the lowest trough so far...
        intLandscapeLowestPoint = intLandscapeY(intLandscapeCount)                              'record it
    End If
Next intLandscapeCount
intLandscapeY(intTerrainComplexity) = intLandscapeY(0)                                          'make the height of the end of the landscape the same as the start

intLandscapeXIncrement = xres / intTerrainComplexity                                            'calculate a intTerrainComplexity'th of the width of the users screen in pixels as the landscape will be drawn with intTerrainComplexity number of lines accross the screen
                                                                                                'Construct an array with as many occurences as the width of the screen, holding a list of all y coordinates to use for collision detection
intlandscapemarker = 1
intLandscapeYCollision(0) = intLandscapeY(0)
For intLandscapeCount = 0 To intTerrainComplexity - 1
    intLandscapeYCollision(intLandscapeCount2) = intLandscapeY(intLandscapeCount) + ((intLandscapeY(intLandscapeCount + 1) - intLandscapeY(intLandscapeCount)) / (xres / intTerrainComplexity))
        
    For intLandscapeCount2 = intlandscapemarker To (intlandscapemarker + intLandscapeXIncrement) Step 1
        intLandscapeYCollision(intLandscapeCount2) = intLandscapeYCollision(intLandscapeCount2 - 1) + ((intLandscapeY(intLandscapeCount + 1) - intLandscapeY(intLandscapeCount)) / (xres / intTerrainComplexity))
    Next intLandscapeCount2
    intlandscapemarker = intlandscapemarker + intLandscapeXIncrement
Next intLandscapeCount

GenerateSwitch                                                                                  'create coordinates for the switch to open the lab gates
                                                                                                'set coordinates of left wall of landing pad
intPadLeftTopLeftx = 1                                                                          'permanently fixed
intPadLeftTopLefty = 1                                                                          'permanently fixed
intPadLeftBottomRightx = xres / 200                                                             'permanently fixed
intPadLeftBottomRighty = yres / 5.3                                                             'permanently fixed
                                                                                                'set coordinates of right wall of landing pad
intPadRightTopLeftx = xres / 16.5                                                               'permanently fixed
intPadRightTopLefty = 1                                                                         'permanently fixed
intPadRightBottomRightx = xres / 14.5                                                           'permanently fixed
intPadRightBottomRighty = yres / 17.5                                                           'permanently fixed
                                                                                                'set coordinates of gate (closed)
intGateTopLeftx = xres / 15.8                                                                   'permanently fixed
intGateTopLefty = yres / 17.5                                                                   'permanently fixed
intGateBottomRightx = xres / 14.5                                                               'permanently fixed
intGateBottomRighty = yres / 5.3                                                                'varies according to whether gate open or not

End Sub

Sub SetUpNewGame()
                                                                                                '******** initial settings that will apply to every new GAME *****************
intCircleWidth = xres / 200                                                                     'used to determine radius of capsules, turrets, etc
counter = 256
intColourCycle = 50
intColourCycle2 = 150
intColourCycle3 = 250
intStatsXCoordinate = xres - 140                                                                'calculate where to start printing text for the statistics
                                                                                                'calculate some distances from middle of ship to edges (to allow collision detection regardless of screen resolution
intShipDistMidPill = 20                                                                         '20 in 1600x1200
intShipDistMidBott = yres / 80                                                                  '15 in 1600x1200
intShipDistMidSide = xres / 123                                                                 '13 in 1600x1200
intShipDistMidCent = yres / 171                                                                 '7 in 1600x1200
intDrugsWavelength = xres / 2.66                                                                '600 in 1600x1200
intMysteryWavelength = xres / 2.66                                                              '600 in 1600x1200
intAntiGravityWavelength1 = xres / 2.66                                                         '600 in 1600x1200
intAntiGravityWavelength2 = xres / 2.13                                                         '750 in 1600x1200
intScientistHeight = yres / 172                                                                 '7 in 1600x1200
intScientistHeightFromGround = yres / 109                                                       '11 in 1600x1200
intTakeOffDelay = 0
blnCollided = False                                                                             'set crashed flag to false
intLevel = 0                                                                                    'will get added to when the level is created
intScore = 0                                                                                    'zero score
intLives = 2                                                                                    '1 extra gets added as soon as the first level is entered
charCenterMessage = ""                                                                          'clear any messages that may be hanging around from last game
intLevelGapCounter = 0
End Sub

Sub SetUpNewLevel()
                                                                                                '********* initial settings that will apply to every new LEVEL *****************
StopLandedSound                                                                                 'just in case
Randomize Timer
intGeneralCounter = 0                                                                           'reset 'trigger' counter
intScientistsOutCount = 0                                                                       'number of scientists needing to be rescued
blnDrugs = False                                                                                'used to determine whether drugs are on screen or not
intDrugType = 3                                                                                 'used to say which drugs have been taken - 1 is reverse controls, 2 is faster rotation, 3 is no drugs
blnFuel = False                                                                                 'used to determine whether fuel is on screen or not
GenerateLandscape                                                                               'create a new landscape
gravity = 0.002                                                                                 'gravity - velocity gets incremented by this amount
intFuelWarningLevel = 500                                                                       'COULD be changed for each new level
intRefuelLevel = 1000                                                                           'COULD be changed for each new level
intRotationRate = 2                                                                             'speed that ship rotates at
intLives = intLives + 1                                                                         'bonus life for completing level
blnScientist1 = False                                                                           'set all scientists to 'not shown yet' status
blnScientist2 = False
blnScientist3 = False
blnScientist4 = False
blnScientist5 = False
blnScientist1Complete = False                                                                   'set all scientists to 'not collected' status
blnScientist2Complete = False
blnScientist3Complete = False
blnScientist4Complete = False
blnScientist5Complete = False
blnScientist1Free = False                                                                       'set all scientists to 'not free' status
blnScientist2Free = False
blnScientist3Free = False
blnScientist4Free = False
blnScientist5Free = False
blnMystery = False
blnMysteryIncrement = False
blnDanceMode = False
blnTurret = False                                                                               'ready for new turret to be generated
blnTurretDestroyed = False                                                                      'ready for new turret to be generated
blnTurretShot = False
blnTurret2 = False                                                                               'ready for new turret to be generated
blnTurret2Destroyed = False                                                                      'ready for new turret to be generated
blnTurret2Shot = False
blnTurret3 = False                                                                               'ready for new turret to be generated
blnTurret3Destroyed = False                                                                      'ready for new turret to be generated
blnTurret3Shot = False
blnTurret4 = False                                                                               'ready for new turret to be generated
blnTurret4Destroyed = False                                                                      'ready for new turret to be generated
blnTurret4Shot = False
blnTurret5 = False                                                                               'ready for new turret to be generated
blnTurret5Destroyed = False                                                                      'ready for new turret to be generated
blnTurret5Shot = False
blnTurret6 = False                                                                               'ready for new turret to be generated
blnTurret6Destroyed = False                                                                      'ready for new turret to be generated
blnTurret6Shot = False
blnTurret7 = False                                                                               'ready for new turret to be generated
blnTurret7Destroyed = False                                                                      'ready for new turret to be generated
blnTurret7Shot = False
intLevel = intLevel + 1                                                                         'add to the level number
blnLevelComplete = False                                                                        'ready for completion of the new level
blnShowLevelScreen = True                                                                       'before player starts this level, an intro screen will be shown
blnInGame = False                                                                               'once the intro screen is shown this will get set to true
End Sub

Sub SetUpNewLife()
                                                                                                '************ initial settings that will apply to every new LIFE *****************
                                                                                                'this does NOT award a new life...it is used to start using a life the player already has (eg after crashing)
blnCollided = False                                                                             'make sure crashed status is false
intGeneralCounter = 0
intPlotShipDistMidCent = xres / 228.57                                                          '7 in 1600x1200
intPlotShipDistMidOuterRing = xres / 114.28                                                     '14 in 1600x1200
intPlotShipDistMidCore = xres / 133.33                                                          '12 in 1600x1200
intBeamSizeMax = yres / 24                                                                      '50 in 1600x1200
intBeamSizeMin = yres / 60                                                                      '20 in 1600x1200
intBeamSize = intBeamSizeMin
intFlameSizeMax = yres / 40                                                                     '30 in 1600x1200
intFlameSizeMin = yres / 70.6                                                                   '17 in 1600x1200
intFlameSize = intFlameSizeMin
blnDoubleGravity = False
intDoubleGravityCount = 0
blnTinyShip = False
intTinyShipCount = 0
blnAntiGravity = False
If gravity < 0 Then
    gravity = 0 - gravity
End If
intDrugType = 3                                                                                 'cancel the effect of any previously collected drugs
intRotationRate = 2                                                                             'cancel effect of rotation boost that may have been in effect
blnAnyScientistOnBoard = False
blnScientist1OnBoard = False
blnScientist2OnBoard = False
blnScientist3OnBoard = False
blnScientist4OnBoard = False
blnScientist5OnBoard = False
InitialiseShip                                                                                  'Put ship in place on screen
leftpressed = False
rightpressed = False
thrustpressed = False
beampressed = False
shootpressed = False
blnSwitchPressed = False
intAddScore = 0                                                                                 'ready for points being awarded
End Sub

Sub CollisionDetection()
                                                                                                'Landscape collision
    If intShipMiddleY + intShipDistMidBott > intLandscapeHighestPoint Then                      'check ships height..if its higher than highest landscape point, theres no need to check for collision
        If intShipMiddleX - intShipDistMidSide >= 0 And intShipMiddleX + intShipDistMidSide <= xres Then 'make sure the ship is in a position within the scope of the array used to check the y coordinate (it drops below 0 briefly when moving off side of screen)
            If intShipMiddleY + intShipDistMidCent > intLandscapeYCollision(intShipMiddleX) Or _
            intShipMiddleY + intShipDistMidBott > yres Or _
            intPlotShipLeftY >= intLandscapeYCollision(intPlotShipLeftX) Or _
            intPlotShipRightY >= intLandscapeYCollision(intPlotShipRightX) Or _
            intPlotShipTopY >= intLandscapeYCollision(intPlotShipTopX) Then
                blnCollided = True                                                              'ready for the crash sequence
                LifeLost                                                                        'lose a life
            End If
        End If
    End If
    
    If intShipMiddleY + intShipDistMidBott < intPadLeftBottomRighty Then                        'collision with left wall of landing pad
        If (intPlotShipLeftX >= intPadLeftTopLeftx And intPlotShipLeftX < intPadLeftBottomRightx) Or _
        (intPlotShipRightX >= intPadLeftTopLeftx And intPlotShipRightX < intPadLeftBottomRightx) Or _
        (intPlotShipTopX >= intPadLeftTopLeftx And intPlotShipTopX < intPadLeftBottomRightx) Or _
        (intPlotShipCenterX >= intPadLeftTopLeftx And intPlotShipCenterX < intPadLeftBottomRightx) Then
            blnCollided = True
            LifeLost
        End If
    End If
    
    If intShipMiddleY + intShipDistMidBott < intPadRightBottomRighty Then                       'collision with right wall of landing pad
        If (intPlotShipLeftX >= intPadRightTopLeftx And intPlotShipLeftX < intPadRightBottomRightx) Or _
        (intPlotShipRightX >= intPadRightTopLeftx And intPlotShipRightX < intPadRightBottomRightx) Or _
        (intPlotShipTopX >= intPadRightTopLeftx And intPlotShipTopX < intPadRightBottomRightx) Or _
        (intPlotShipCenterX >= intPadRightTopLeftx And intPlotShipCenterX < intPadRightBottomRightx) Then
            blnCollided = True
            LifeLost
        End If
    End If
    
    If intShipMiddleY + intShipDistMidBott < intGateBottomRighty Then                           'collision with landing pad gate
        If (intPlotShipLeftX >= intGateTopLeftx And intPlotShipLeftX < intGateBottomRightx) Or _
        (intPlotShipRightX >= intGateTopLeftx And intPlotShipRightX < intGateBottomRightx) Or _
        (intPlotShipTopX >= intGateTopLeftx And intPlotShipTopX < intGateBottomRightx) Or _
        (intPlotShipCenterX >= intGateTopLeftx And intPlotShipCenterX < intGateBottomRightx) Then
            blnCollided = True
            LifeLost
        End If
    End If
End Sub

Sub CheckForLandingOrCrash()
                                                                                                'check for good landing
    If intPlotShipLeftX < xres / 14.5 And intPlotShipLeftY < yres / 5 Then
        If (intPlotShipLeftX >= 0 And intPlotShipLeftX < xres / intShipDistMidSide) Or _
        (intPlotShipRightX >= 0 And intPlotShipRightX < xres / intShipDistMidSide) Or _
        (intPlotShipTopX >= 0 And intPlotShipTopX < xres / intShipDistMidSide) Or _
        (intPlotShipCenterX >= 0 And intPlotShipCenterX < xres / intShipDistMidSide) Then
            If (intPlotShipLeftY > yres / 5.5 And intPlotShipLeftY < yres / 5.4) Or _
            (intPlotShipRightY > yres / 5.5 And intPlotShipRightY < yres / 5.4) Or _
            (intPlotShipTopY > yres / 5.5 And intPlotShipTopY < yres / 5.4) Or _
            (intPlotShipCenterY > yres / 5.5 And intPlotShipCenterY < yres / 5.4) Then
                                                                                                'check conditions for landing
                If velocity >= 0 And velocity <= 0.6 Then                                       'check speed isn't too great
                    If stepval > 354 Or stepval < 6 Then                                        'check bearing is nearly vertical
                        StopThrustSound                                                         'silence thrusters
                        thrustpressed = False                                                   'kill thrusters
                        PlayLandedSound                                                         'the eagle has landed
                        stepval = 0                                                             'ensure ship is exactly vertical now its landed
                        intShipMiddleY = (yres / 5.4) - yres / 100                              'set the ship on level ground, ready for launch
                        blnShipLanded = True                                                    'landing is complete
                        velocity = -1                                                           'give a small booster to get ship airborne again (once launched)
                        velocityx = 1                                                           'give some horizontal velocity to force exit of lab area
                        intPlotShipDistMidCent = xres / 228.57                                  '7 in 1600x1200
                        intPlotShipDistMidOuterRing = xres / 114.28                             '14 in 1600x1200
                        intPlotShipDistMidCore = xres / 133.33                                  '12 in 1600x1200
                        If blnTinyShip = True Then
                            intPlotShipDistMidCent = intPlotShipDistMidCent / 2
                            intPlotShipDistMidOuterRing = intPlotShipDistMidOuterRing / 2
                            intPlotShipDistMidCore = intPlotShipDistMidCore / 2
                            intShipMiddleY = intShipMiddleY + (yres / 240)
                        End If
                                                                                                'redraw the ship in its adjusted position
                        intPlotShipCenterX = intShipMiddleX + Sin(stepval * Rad) * intPlotShipDistMidCent
                        intPlotShipCenterY = intShipMiddleY + Cos(stepval * Rad) * intPlotShipDistMidCent
                        intPlotShipLeftX = intShipMiddleX + Sin((stepval - 40) * Rad) * intPlotShipDistMidOuterRing
                        intPlotShipLeftY = intShipMiddleY + Cos((stepval - 40) * Rad) * intPlotShipDistMidOuterRing
                        intPlotShipRightX = intShipMiddleX + Sin((stepval + 40) * Rad) * intPlotShipDistMidOuterRing
                        intPlotShipRightY = intShipMiddleY + Cos((stepval + 40) * Rad) * intPlotShipDistMidOuterRing
                        intPlotShipTopX = intShipMiddleX + Sin((stepval - 180) * Rad) * intPlotShipDistMidOuterRing
                        intPlotShipTopY = intShipMiddleY + Cos((stepval - 180) * Rad) * intPlotShipDistMidOuterRing
                        intPlotCoreLeftX = intShipMiddleX + Sin((stepval - 15) * Rad) * intPlotShipDistMidCore
                        intPlotCoreLeftY = intShipMiddleY + Cos((stepval - 15) * Rad) * intPlotShipDistMidCore
                        intPlotCoreRightX = intShipMiddleX + Sin((stepval + 15) * Rad) * intPlotShipDistMidCore
                        intPlotCoreRightY = intShipMiddleY + Cos((stepval + 15) * Rad) * intPlotShipDistMidCore
                        blt                                                                     'draw screen
                        If blnScientist1OnBoard = True Then                                     'if scientist1 is about to be set free
                            intScientist1X = intPlotShipLeftX                                   'place him next to the ship
                            intScientist1Y = intPlotShipLeftY
                        End If
                        If blnScientist2OnBoard = True Then                                     'as above
                            intScientist2X = intPlotShipLeftX
                            intScientist2Y = intPlotShipLeftY
                        End If
                        If blnScientist3OnBoard = True Then                                     'as above
                            intScientist3X = intPlotShipLeftX
                            intScientist3Y = intPlotShipLeftY
                        End If
                        If blnScientist4OnBoard = True Then                                     'as above
                            intScientist4X = intPlotShipLeftX
                            intScientist4Y = intPlotShipLeftY
                        End If
                        If blnScientist5OnBoard = True Then                                     'as above
                            intScientist5X = intPlotShipLeftX
                            intScientist5Y = intPlotShipLeftY
                        End If
                                                                                                'at this point the ship will take off again
                        blnAnyScientistOnBoard = False                                          'no scientist on board as now set free
                        PlayThrustSound                                                         'make thrust sound play as launch occurs
                    End If
                End If
            End If
        End If
        
                                                                                                'Landing pad collision
        If (intPlotShipLeftX >= 0 And intPlotShipLeftX < xres / 14.5) Or _
        (intPlotShipRightX >= 0 And intPlotShipRightX < xres / 14.5) Or _
        (intPlotShipTopX >= 0 And intPlotShipTopX < xres / 14.5) Or _
        (intPlotShipCenterX >= 0 And intPlotShipCenterX < xres / 14.5) Then
            If (intPlotShipTopY > yres / 5.4 And intPlotShipTopY < yres / 5) Or _
            (intPlotShipLeftY > yres / 5.4 And intPlotShipLeftY < yres / 5) Or _
            (intPlotShipRightY > yres / 5.4 And intPlotShipRightY < yres / 5) Or _
            (intPlotShipCenterY > yres / 5.4 And intPlotShipCenterY < yres / 5) Then
                blnCollided = True                                                              'crashed!
                LifeLost                                                                        'lose a life
            End If
        End If
    End If
End Sub

Sub CheckForPowerUpCollect()
                                                                                                'check whether ship is collecting fuel
    If blnFuel = True Then                                                                      'only check if the fuel is actually on screen
        If (intShipMiddleX >= intFuelX - intShipDistMidBott And intShipMiddleX <= intFuelX + intShipDistMidBott) Then
            If (intShipMiddleY >= intFuelY - intShipDistMidPill And intShipMiddleY <= intFuelY + intShipDistMidPill) Then
                blnFuel = False                                                                 'enabling further fuel to be spawned
                AddScore (50)                                                                   'award some points
                PlayCollectedSound                                                              'ping!!
                blncentermessage = True                                                         'prepare to show a message
                charCenterMessage = "      1000 extra fuel"                                     'the message
                intFuelLevel = intFuelLevel + 1000                                              'add to the fuel level
            End If
        End If
    End If
    
                                                                                                'check whether ship is collecting drugs
    If blnDrugs = True Then                                                                     'only check if the drugs are actually on screen
        If (intShipMiddleX >= intDrugsX - intShipDistMidBott And intShipMiddleX <= intDrugsX + intShipDistMidBott) Then
            If (intShipMiddleY >= intDrugsY - intShipDistMidPill And intShipMiddleY <= intDrugsY + intShipDistMidPill) Then
                blnDrugs = False                                                                'enabling further drugs to be spawned
                Randomize Timer                                                                 'random number used to decide what drug effects are
                intRandomNumber = Rnd(Timer)
                If intRandomNumber > 0.5 Then
                    intDrugType = 1                                                             ' has reverse effect on rotation
                    blncentermessage = True                                                     'show message of what the drugs do
                    charCenterMessage = "Reversed controls.. + 100 points"
                Else
                    intDrugType = 2                                                             'increases rotation speed
                    intRotationRate = 5                                                         'speed up rotation
                    blncentermessage = True                                                     'show message of what the drugs do
                    charCenterMessage = "Rotation boost.. + 100 points"
                End If
                AddScore (100)                                                                  'award some points
                PlayCollectedSound                                                              'ping!!
            End If
        End If
    End If
    
                                                                                                'check whether ship is collecting mystery
    If blnMystery = True Then                                                                   'only check if the mystery is actually on screen
        If (intShipMiddleX >= intMysteryX - intShipDistMidBott And intShipMiddleX <= intMysteryX + intShipDistMidBott) Then
            If (intShipMiddleY >= intMysteryY - intShipDistMidPill And intShipMiddleY <= intMysteryY + intShipDistMidPill) Then
                blnMystery = False                                                              'enables more mysteries to be spawned
                Randomize Timer                                                                 'random number used to determine what the mystery does
                intRandomNumber = Rnd(Timer)
                If intRandomNumber >= 0.66 Then
                    DanceMode                                                                   'make the scientists dance
                    AddScore (500)                                                              'award some points
                End If
                If intRandomNumber >= 0.33 And intRandomNumber < 0.66 Then
                    PlayCollectedSound                                                          'ping!!
                    blncentermessage = True                                                     'center message to show what mystery does
                    charCenterMessage = "Double gravity.. +200 points"
                    DoubleGravity                                                               'doubles the gravity for a limited time
                    AddScore (150)                                                              'award some points
                End If
                If intRandomNumber >= 0 And intRandomNumber < 0.33 Then
                    PlayCollectedSound                                                          'ping!!
                    AddScore (200)                                                              'award some points
                    blncentermessage = True                                                     'center message to show what mystery does
                    charCenterMessage = "    Ship shrunk!.. +200"
                    TinyShip                                                                    'makes the ship shrink
                End If
            End If
        End If
    End If
    
                                                                                                'check whether ship is collecting antigravity powerup
    If blnAntiGravityShown = True Then                                                          'only check if the antigrav is actually on screen
        If (intShipMiddleX >= intAntiGravityX - intShipDistMidBott And intShipMiddleX <= intAntiGravityX + intShipDistMidBott) Then
            If (intShipMiddleY >= intAntiGravityY - intShipDistMidPill And intShipMiddleY <= intAntiGravityY + intShipDistMidPill) Then
                blnAntiGravityShown = False                                                     'enables more to be spawned
                AddScore (10)                                                                   'award some points
                PlayCollectedSound                                                              'ping!!
                blncentermessage = True                                                         'center message to explain gravity change
                charCenterMessage = "Reversed gravity.. + 10 points"
                If blnAntiGravity = True Then                                                   'if already in anti grav mode
                    gravity = 0 - gravity                                                       'revert to normal gravity
                    blnAntiGravity = False
                Else                                                                            'otherwise, turn on antigravity
                    gravity = gravity - (2 * gravity)
                    blnAntiGravity = True
                End If
            End If
        End If
    End If
End Sub

Sub CheckForRotation()
    If leftpressed = True Then                                                                  'check whether rotate left button has been pressed
        If stepval >= 360 Then
            stepval = stepval - 360
        End If
        If Not intDrugType = 1 Then                                                             'is player on adverse drugs?
            stepval = stepval + intRotationRate                                                 'if not, then rotate normally
        Else
            stepval = stepval - intRotationRate                                                 'if they are, reverse the rotation
        End If
    End If
    If rightpressed = True Then                                                                 'check whether rotate right button has been pressed
        If stepval <= 0 Then
           stepval = stepval + 360
        End If
        If Not intDrugType = 1 Then                                                             'is player on adverse drugs?
            stepval = stepval - intRotationRate                                                 'if not then rotate normally
        Else
            stepval = stepval + intRotationRate                                                 'if they are, reverse the rotation
        End If
    End If
End Sub

Sub CheckForScientistCollect()
    If beampressed = True And blnCollided = False Then                                          'check whether tractor beam button has been pressed
        PlayTractorBeamSound                                                                    'play the tractor beam sound
                                                                                                'check whether a scientist can be picked up
        If blnAnyScientistOnBoard = False Then                                                  'no point in checking if there's already 1 on board
                                                                                                
            If blnScientist1OnBoard = False Then                                                'check if scientist1 can be collected
                If intScientist1X > intPlotShipLeftX And intScientist1X < intPlotShipRightX Then 'check scientist coordinates are within the beams range
                    If intScientist1Y < intBeamBottLeftY And intScientist1Y < intBeamBottRightY Then
                        If intScientist1Y > intPlotShipLeftY And intScientist1Y > intPlotShipRightY Then
                            blnScientist1 = False                                               'stop drawing scientist1 now he's in ship
                            blnAnyScientistOnBoard = True                                       'flag to a scientist is in ship
                            PlayGotScientistSound                                               'ping!!
                            blnScientist1OnBoard = True                                         'flag to say scientist1 is in ship
                            AddScore (100)                                                      'award some points
                            intScientistsOutCount = intScientistsOutCount - 1                   'decrement count of number of scientists on the ground
                        End If
                    End If
                End If
            End If
            If blnScientist2OnBoard = False Then                                                'check if scientist2 can be collected
                If intScientist2X > intPlotShipLeftX And intScientist2X < intPlotShipRightX Then
                    If intScientist2Y < intBeamBottLeftY And intScientist2Y < intBeamBottRightY Then
                        If intScientist2Y > intPlotShipLeftY And intScientist2Y > intPlotShipRightY Then
                            blnScientist2 = False
                            blnAnyScientistOnBoard = True
                            PlayGotScientistSound
                            blnScientist2OnBoard = True
                            AddScore (100)
                            intScientistsOutCount = intScientistsOutCount - 1
                        End If
                    End If
                End If
            End If
            If blnScientist3OnBoard = False Then                                                'check if scientist3 can be collected
                If intScientist3X > intPlotShipLeftX And intScientist3X < intPlotShipRightX Then
                    If intScientist3Y < intBeamBottLeftY And intScientist3Y < intBeamBottRightY Then
                        If intScientist3Y > intPlotShipLeftY And intScientist3Y > intPlotShipRightY Then
                            blnScientist3 = False
                            blnAnyScientistOnBoard = True
                            PlayGotScientistSound
                            blnScientist3OnBoard = True
                            AddScore (100)
                            intScientistsOutCount = intScientistsOutCount - 1
                        End If
                    End If
                End If
            End If
            If blnScientist4OnBoard = False Then                                                'check if scientist4 can be collected
                If intScientist4X > intPlotShipLeftX And intScientist4X < intPlotShipRightX Then
                    If intScientist4Y < intBeamBottLeftY And intScientist4Y < intBeamBottRightY Then
                        If intScientist4Y > intPlotShipLeftY And intScientist4Y > intPlotShipRightY Then
                            blnScientist4 = False
                            blnAnyScientistOnBoard = True
                            PlayGotScientistSound
                            blnScientist4OnBoard = True
                            AddScore (100)
                            intScientistsOutCount = intScientistsOutCount - 1
                        End If
                    End If
                End If
            End If
            If blnScientist5OnBoard = False Then                                                'check if scientist5 can be collected
                If intScientist5X > intPlotShipLeftX And intScientist5X < intPlotShipRightX Then
                    If intScientist5Y < intBeamBottLeftY And intScientist5Y < intBeamBottRightY Then
                        If intScientist5Y > intPlotShipLeftY And intScientist5Y > intPlotShipRightY Then
                            blnScientist5 = False
                            blnAnyScientistOnBoard = True
                            PlayGotScientistSound
                            blnScientist5OnBoard = True
                            AddScore (100)
                            intScientistsOutCount = intScientistsOutCount - 1
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Sub CheckForThrust()
                                                                                                'check whether thrust button has been pressed
    If thrustpressed = True Then
        If intFuelLevel > 0 Then
            PlayThrustSound
                                                                                                'reduce fuel level
            intFuelLevel = intFuelLevel - 1
            If intFuelLevel = intFuelWarningLevel Then
                PlayWarningSound
            End If
            If blnAntiGravity = False Then
                velocity = velocity - (gravity * (-6 + 16 * Cos(stepval * Rad)))
            Else
                velocity = velocity + (gravity * (-6 + 16 * Cos(stepval * Rad)))
            End If
            If Not stepval = 360 And Not stepval = 0 And Not stepval = 180 Then
                If blnAntiGravity = False Then
                    velocityx = velocityx - gravity * (16 * Sin(stepval * Rad))
                Else
                    velocityx = velocityx + gravity * (16 * Sin(stepval * Rad))
                End If
            End If
                                                                                                'limit vertical downward top speed
            If velocity >= 10 Then
                velocity = 10
            End If
                                                                                                'limit vertical upward top speed
            If velocity <= -5 Then
                velocity = -5
            End If
                                                                                                'limit horizontal speed
            If velocityx >= 5 Then
                velocityx = 5
            End If
            If velocityx <= -5 Then
                velocityx = -5
            End If
        Else
                                                                                                'fuel has run out..stop thrusting
            thrustpressed = False
        End If
    Else
        StopThrustSound
                                                                                                'thrust is not pressed
        velocity = velocity + (gravity * 6)
        If velocity >= 5 Then
            If velocity >= 5.5 Then
                velocity = velocity - (gravity * 6)
                velocity = velocity - 0.01
            Else
                velocity = 5
            End If
        End If
    End If
End Sub

Sub CheckForGateOpenClose()
                                                                                                'check whether the gate needs its coordinates changed (opening or closing)
                                                                                                ' does gate need closing? (scientist dropped off)
    If blnSwitchPressed = True Then
        If blnAnyScientistOnBoard = False Then
            If intGateBottomRighty < yres / 5.3 Then
                 If intShipMiddleX > xres / 14.5 Then
                    intGateBottomRighty = intGateBottomRighty + 1
                    If blnSlidingDoorPlayed = False Then
                        PlaySlidingDoorSound
                        blnSlidingDoorPlayed = True
                    End If
                    If intGateBottomRighty >= yres / 5.3 Then
                        StopSlidingDoorSound
                        intGateBottomRighty = yres / 5.3
                        blnSlidingDoorPlayed = False
                        PlayVaultSound
                        blnSwitchPressed = False
                    End If
                End If
            End If
        End If
    End If
    
                                                                                                'does gate need closing? (player has lost life while the gate was open (and has therefore lost the scientist he was carrying))
    If blnSwitchPressed = False Then
        If blnAnyScientistOnBoard = False Then
            If intGateBottomRighty < yres / 5.3 Then
                intGateBottomRighty = intGateBottomRighty + 1
                If blnSlidingDoorPlayed = False Then
                    PlaySlidingDoorSound
                    blnSlidingDoorPlayed = True
                End If
                If intGateBottomRighty >= yres / 5.3 Then
                    StopSlidingDoorSound
                    intGateBottomRighty = yres / 5.3
                    blnSlidingDoorPlayed = False
                    PlayVaultSound
                    blnSwitchPressed = False
                End If
            End If
        End If
    End If
    
                                                                                                ' does gate need opening?
    If blnSwitchPressed = True Then
        If blnAnyScientistOnBoard = True Then
            If intGateBottomRighty > intGateTopLefty Then
                intGateBottomRighty = intGateBottomRighty - 1
                If blnSlidingDoorPlayed = False Then
                    PlaySlidingDoorSound
                    blnSlidingDoorPlayed = True
                End If
            Else
                StopSlidingDoorSound
                blnSlidingDoorPlayed = False
            End If
        End If
    End If
End Sub

Sub SetShipCoordinates()
                                                                                                'set the point of the middle of the ship on the screen
    If blnShipLanded = False Then                                                               'only bother if ship is airborne
        intShipMiddleY = intShipMiddleY + velocity
        intShipMiddleX = intShipMiddleX + velocityx
        If intShipMiddleX > xres + 15 Then
            intShipMiddleX = -15
        End If
        If intShipMiddleX < -15 Then
            intShipMiddleX = xres + 15
        End If
        If intShipMiddleY < 15 Then
            blnCollided = True
            LifeLost
        End If
        If intShipMiddleY > 1215 Then
            intShipMiddleY = -15
        End If
        
                                                                                                'set the ships coordinates
        intPlotShipCenterX = intShipMiddleX + Sin(stepval * Rad) * intPlotShipDistMidCent
        intPlotShipCenterY = intShipMiddleY + Cos(stepval * Rad) * intPlotShipDistMidCent
        intPlotShipLeftX = intShipMiddleX + Sin((stepval - 40) * Rad) * intPlotShipDistMidOuterRing
        intPlotShipLeftY = intShipMiddleY + Cos((stepval - 40) * Rad) * intPlotShipDistMidOuterRing
        intPlotShipRightX = intShipMiddleX + Sin((stepval + 40) * Rad) * intPlotShipDistMidOuterRing
        intPlotShipRightY = intShipMiddleY + Cos((stepval + 40) * Rad) * intPlotShipDistMidOuterRing
        intPlotShipTopX = intShipMiddleX + Sin((stepval - 180) * Rad) * intPlotShipDistMidOuterRing
        intPlotShipTopY = intShipMiddleY + Cos((stepval - 180) * Rad) * intPlotShipDistMidOuterRing
        intPlotCoreLeftX = intShipMiddleX + Sin((stepval - 15) * Rad) * intPlotShipDistMidCore
        intPlotCoreLeftY = intShipMiddleY + Cos((stepval - 15) * Rad) * intPlotShipDistMidCore
        intPlotCoreRightX = intShipMiddleX + Sin((stepval + 15) * Rad) * intPlotShipDistMidCore
        intPlotCoreRightY = intShipMiddleY + Cos((stepval + 15) * Rad) * intPlotShipDistMidCore
                                                                                                'limit the size that the flame can reach
        If intFlameSize >= intFlameSizeMax Then
            intFlameSize = intFlameSizeMax
        End If
                                                                                                'set the size of the core flame
        intPlotCoreCenterX = intShipMiddleX + Sin((stepval) * Rad) * (intFlameSize - (intFlameSize / 3))
        intPlotCoreCenterY = intShipMiddleY + Cos((stepval) * Rad) * (intFlameSize - (intFlameSize / 3))
                                                                                                'set the size of the external flame
        intPlotFlameCenterX = intShipMiddleX + Sin((stepval) * Rad) * intFlameSize
        intPlotFlameCenterY = intShipMiddleY + Cos((stepval) * Rad) * intFlameSize
    End If
End Sub


Sub DrawMenu()

                                                                                                'text
fontinfo2.Size = 48
backbuffer.SetFont fontinfo2

If blnCreateNewLineDest = True Then
    Randomize Timer
    sngMenuLineYDest = yres - yres / 10 - ((Rnd(Timer)) * yres / 5)
    blnCreateNewLineDest = False
Else
    If sngMenuLineY < sngMenuLineYDest Then
        sngMenuLineY = sngMenuLineY + 2
        If sngMenuLineY >= sngMenuLineYDest Then
            blnCreateNewLineDest = True
        End If
    Else
        sngMenuLineY = sngMenuLineY - 2
        If sngMenuLineY <= sngMenuLineYDest Then
            blnCreateNewLineDest = True
        End If
    End If
End If

If blnCreateNewLineDest2 = True Then
    Randomize Timer
    sngMenuLineYDest2 = yres - yres / 10 - ((Rnd(Timer)) * yres / 5)
    blnCreateNewLineDest2 = False
Else
    If sngMenuLineY2 < sngMenuLineYDest2 Then
        sngMenuLineY2 = sngMenuLineY2 + 2
        If sngMenuLineY2 >= sngMenuLineYDest2 Then
            blnCreateNewLineDest2 = True
        End If
    Else
        sngMenuLineY2 = sngMenuLineY2 - 2
        If sngMenuLineY2 <= sngMenuLineYDest2 Then
            blnCreateNewLineDest2 = True
        End If
    End If
End If

If blnCreateNewLineDest3 = True Then
    Randomize Timer
    sngMenuLineYDest3 = yres - yres / 10 - ((Rnd(Timer)) * yres / 5)
    blnCreateNewLineDest3 = False
Else
    If sngMenuLineY3 < sngMenuLineYDest3 Then
        sngMenuLineY3 = sngMenuLineY3 + 2
        If sngMenuLineY3 >= sngMenuLineYDest3 Then
            blnCreateNewLineDest3 = True
        End If
    Else
        sngMenuLineY3 = sngMenuLineY3 - 2
        If sngMenuLineY3 <= sngMenuLineYDest3 Then
            blnCreateNewLineDest3 = True
        End If
    End If
End If

If blnCreateNewLineDest4 = True Then
    Randomize Timer
    sngMenuLineYDest4 = yres - yres / 10 - ((Rnd(Timer)) * yres / 5)
    blnCreateNewLineDest4 = False
Else
    If sngMenuLineY4 < sngMenuLineYDest4 Then
        sngMenuLineY4 = sngMenuLineY4 + 2
        If sngMenuLineY4 >= sngMenuLineYDest4 Then
            blnCreateNewLineDest4 = True
        End If
    Else
        sngMenuLineY4 = sngMenuLineY4 - 2
        If sngMenuLineY4 <= sngMenuLineYDest4 Then
            blnCreateNewLineDest4 = True
        End If
    End If
End If

If blnCreateNewLineDest5 = True Then
    Randomize Timer
    sngMenuLineYDest5 = yres - yres / 10 - ((Rnd(Timer)) * yres / 5)
    blnCreateNewLineDest5 = False
Else
    If sngMenuLineY5 < sngMenuLineYDest5 Then
        sngMenuLineY5 = sngMenuLineY5 + 2
        If sngMenuLineY5 >= sngMenuLineYDest5 Then
            blnCreateNewLineDest5 = True
        End If
    Else
        sngMenuLineY5 = sngMenuLineY5 - 2
        If sngMenuLineY5 <= sngMenuLineYDest5 Then
            blnCreateNewLineDest5 = True
        End If
    End If
End If

                                                                                                    '5 horizontal lines

backbuffer.SetForeColor RGB(11, 16, 63)
backbuffer.DrawLine 0, sngMenuLineY, xres, sngMenuLineY
backbuffer.SetForeColor RGB(21, 36, 83)
backbuffer.DrawLine 0, sngMenuLineY2, xres, sngMenuLineY2
backbuffer.SetForeColor RGB(41, 56, 103)
backbuffer.DrawLine 0, sngMenuLineY3, xres, sngMenuLineY3
backbuffer.SetForeColor RGB(61, 76, 123)
backbuffer.DrawLine 0, sngMenuLineY4, xres, sngMenuLineY4
backbuffer.SetForeColor RGB(81, 96, 143)
backbuffer.DrawLine 0, sngMenuLineY5, xres, sngMenuLineY5

backbuffer.SetForeColor RGB(255, 255, 255)
backbuffer.DrawText xres / 3, yres / 2.4, "1 : Start New Game", False
backbuffer.DrawText xres / 3.1, yres / 2.4 + 60, "2 : View High Scores", False
backbuffer.DrawText xres / 2.3, yres / 2.4 + 125, "Esc : Quit", False

End Sub


Sub InitialiseMenuLines()
sngMenuLineY = yres - yres / 9
sngMenuLineY2 = yres - yres / 9
sngMenuLineY3 = yres - yres / 9
sngMenuLineY4 = yres - yres / 9
sngMenuLineY5 = yres - yres / 9
blnCreateNewLineDest = True
blnCreateNewLineDest2 = True
blnCreateNewLineDest3 = True
blnCreateNewLineDest4 = True
blnCreateNewLineDest5 = True
End Sub

Sub GameOver()
    Dim GameOverDelay As Integer
    
                                                                                                    'set coordinates of ship so it cant be seen
    intPlotShipCenterX = 0
    intPlotShipCenterY = 0
    intPlotShipLeftX = 0
    intPlotShipLeftY = 0
    intPlotShipRightX = 0
    intPlotShipRightY = 0
    intPlotShipTopX = 0
    intPlotShipTopY = 0
    intPlotCoreLeftX = 0
    intPlotCoreLeftY = 0
    intPlotCoreRightX = 0
    intPlotCoreRightY = 0
    Do While GameOverDelay < (yres / 2 + 25)
        blncentermessage = True
        charCenterMessage = "GAME OVER - SCORE: " & intScore
        blt
        DoEvents
        GameOverDelay = GameOverDelay + 1
    Loop
    GameOverDelay = 0
    blnCollided = False
    blnInGame = False
    blnShowLevelScreen = False
    blnInMenu = True
    InitialiseMenuLines
    brunning = False
End Sub

Sub CheckForLevelComplete(intWhichScientist As Integer)

Select Case intWhichScientist
    Case 1
        If blnScientist2Complete = True And blnScientist3Complete = True And blnScientist4Complete = True And blnScientist5Complete = True Then
            blnLevelComplete = True
        End If
    Case 2
        If blnScientist1Complete = True And blnScientist3Complete = True And blnScientist4Complete = True And blnScientist5Complete = True Then
            blnLevelComplete = True
        End If
    Case 3
        If blnScientist1Complete = True And blnScientist2Complete = True And blnScientist4Complete = True And blnScientist5Complete = True Then
            blnLevelComplete = True
        End If
    Case 4
        If blnScientist1Complete = True And blnScientist2Complete = True And blnScientist3Complete = True And blnScientist5Complete = True Then
            blnLevelComplete = True
        End If
    Case 5
        If blnScientist1Complete = True And blnScientist2Complete = True And blnScientist3Complete = True And blnScientist4Complete = True Then
            blnLevelComplete = True
        End If
End Select
End Sub

Sub DrawLevelScreen()

StopTakeoffSound

fontinfo2.Size = 48
backbuffer.SetFont fontinfo2

If blnCreateNewLineDest = True Then
    Randomize Timer
    sngMenuLineYDest = yres - yres / 10 - ((Rnd(Timer)) * yres / 5)
    blnCreateNewLineDest = False
Else
    If sngMenuLineY < sngMenuLineYDest Then
        sngMenuLineY = sngMenuLineY + 3
        If sngMenuLineY >= sngMenuLineYDest Then
            blnCreateNewLineDest = True
        End If
    Else
        sngMenuLineY = sngMenuLineY - 3
        If sngMenuLineY <= sngMenuLineYDest Then
            blnCreateNewLineDest = True
        End If
    End If
End If

If blnCreateNewLineDest2 = True Then
    Randomize Timer
    sngMenuLineYDest2 = yres - yres / 10 - ((Rnd(Timer)) * yres / 5)
    blnCreateNewLineDest2 = False
Else
    If sngMenuLineY2 < sngMenuLineYDest2 Then
        sngMenuLineY2 = sngMenuLineY2 + 3
        If sngMenuLineY2 >= sngMenuLineYDest2 Then
            blnCreateNewLineDest2 = True
        End If
    Else
        sngMenuLineY2 = sngMenuLineY2 - 3
        If sngMenuLineY2 <= sngMenuLineYDest2 Then
            blnCreateNewLineDest2 = True
        End If
    End If
End If

If blnCreateNewLineDest3 = True Then
    Randomize Timer
    sngMenuLineYDest3 = yres - yres / 10 - ((Rnd(Timer)) * yres / 5)
    blnCreateNewLineDest3 = False
Else
    If sngMenuLineY3 < sngMenuLineYDest3 Then
        sngMenuLineY3 = sngMenuLineY3 + 3
        If sngMenuLineY3 >= sngMenuLineYDest3 Then
            blnCreateNewLineDest3 = True
        End If
    Else
        sngMenuLineY3 = sngMenuLineY3 - 3
        If sngMenuLineY3 <= sngMenuLineYDest3 Then
            blnCreateNewLineDest3 = True
        End If
    End If
End If

If blnCreateNewLineDest4 = True Then
    Randomize Timer
    sngMenuLineYDest4 = yres - yres / 10 - ((Rnd(Timer)) * yres / 5)
    blnCreateNewLineDest4 = False
Else
    If sngMenuLineY4 < sngMenuLineYDest4 Then
        sngMenuLineY4 = sngMenuLineY4 + 3
        If sngMenuLineY4 >= sngMenuLineYDest4 Then
            blnCreateNewLineDest4 = True
        End If
    Else
        sngMenuLineY4 = sngMenuLineY4 - 3
        If sngMenuLineY4 <= sngMenuLineYDest4 Then
            blnCreateNewLineDest4 = True
        End If
    End If
End If

If blnCreateNewLineDest5 = True Then
    Randomize Timer
    sngMenuLineYDest5 = yres - yres / 10 - ((Rnd(Timer)) * yres / 5)
    blnCreateNewLineDest5 = False
Else
    If sngMenuLineY5 < sngMenuLineYDest5 Then
        sngMenuLineY5 = sngMenuLineY5 + 3
        If sngMenuLineY5 >= sngMenuLineYDest5 Then
            blnCreateNewLineDest5 = True
        End If
    Else
        sngMenuLineY5 = sngMenuLineY5 - 3
        If sngMenuLineY5 <= sngMenuLineYDest5 Then
            blnCreateNewLineDest5 = True
        End If
    End If
End If

                                                                                                        '5 horizontal lines

backbuffer.SetForeColor RGB(11, 16, 63)
backbuffer.DrawLine 0, sngMenuLineY, xres, sngMenuLineY
backbuffer.SetForeColor RGB(21, 36, 83)
backbuffer.DrawLine 0, sngMenuLineY2, xres, sngMenuLineY2
backbuffer.SetForeColor RGB(41, 56, 103)
backbuffer.DrawLine 0, sngMenuLineY3, xres, sngMenuLineY3
backbuffer.SetForeColor RGB(61, 76, 123)
backbuffer.DrawLine 0, sngMenuLineY4, xres, sngMenuLineY4
backbuffer.SetForeColor RGB(81, 96, 143)
backbuffer.DrawLine 0, sngMenuLineY5, xres, sngMenuLineY5

backbuffer.SetForeColor RGB(255, 255, 255)
If intLevel = 0 Then
    backbuffer.DrawText xres / 2.3, yres / 2.5 + 60, "Level: 1", False
Else
    backbuffer.DrawText xres / 2.3, yres / 2.5 + 60, "Level: " & intLevel, False
End If

If intLevelGapCounter = 0 Then                                                                          'first time round
    lngCurrentTime = GetTickCount()
    lngDesiredTime = 2000
End If
If lngDesiredTime <= GetTickCount() - lngCurrentTime Then                                               'after a 5 second delay
    blnShowLevelScreen = False
    intLevelGapCounter = 0
    blnInGame = True
End If
intLevelGapCounter = intLevelGapCounter + 1
End Sub

