Attribute VB_Name = "cards"
Public Declare Function cdtInit Lib "cards.dll" (ByRef width As Long, ByRef height As Long) As Long

Public Declare Function cdtAnimate Lib "cards.dll" (ByVal hdc As Long, ByVal cardback As Long, ByVal x As Long, ByVal y As Long, ByVal frame As Long) As Long

Public Declare Function cdtDraw Lib "cards.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal card As Long, ByVal cardtype As Long, ByVal color As Long) As Long

Public Declare Function cdtDrawExt Lib "cards.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal card As Long, ByVal cardtype As Long, ByVal color As Long) As Long

Public Declare Sub cdtTerm Lib "cards.dll" ()

'Card back styles
Const ecbCARDHAND = 65
Const ecbCASTLE = 63
Const ecbCROSSHATCH = 53
Const ecbFISH1 = 60
Const ecbFISH2 = 61
Const ecbFLOWERS = 57
Const ecbISLAND = 64
Const ecbROBOT = 56
Const ecbSHELLS = 62
Const ecbTHE_O = 68
Const ecbTHE_X = 67
Const ecbUNUSED = 66
Const ecbVINE1 = 58
Const ecbVINE2 = 59
Const ecbWEAVE1 = 54
Const ecbWEAVE2 = 55

'Suit types
Public Const ecsCLUBS = 0
Public Const ecsDIAMONDS = 1
Public Const ecsHEARTS = 2
Public Const ecsSPADES = 3

'Card types
Const ectBack = 1
Const ectInverted = 2
Const ectNormal = 0

'Card dimmensions
Public CardWidth As Long
Public CardHeight As Long
