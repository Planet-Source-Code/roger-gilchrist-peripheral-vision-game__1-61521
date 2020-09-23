Attribute VB_Name = "modSpectrum"
Option Explicit
'Thanks to aditya8000's 'Visible Light Spectrum' at PSC txtCodeId=61446"
'this is an abstaction of his code for improved portability
'
Private Const CFactor     As Single = 255 / 60
Private Const VFactor     As Single = 60 / 255

Public Sub DrawSpectrum(ctrl As Object, _
                        Optional ByVal bHorz As Boolean = True, _
                        Optional ByVal bDir As Boolean = True, _
                        Optional lngMin As Long = 400, _
                        Optional lngMax As Long = 700, _
                        Optional Wid As Long = -1)

'Draw a spectrun on a control or form Object
'bHorz draws the spectrum horizontally or vertically
'bDir  purple(700) to red(400) or red to purple
'lngMin,lngMax allow you to draw a section of the Wavelength values(400 - 700)
'Wid -1 makes the spectrum cover the whole object, otherwise the non-bHorz dimension

  Dim I     As Single
  Dim sA    As Double
  Dim Range As Single
  Dim si    As Long
  Dim ncol  As Long

  ctrl.ScaleMode = vbTwips
  If lngMax > 700 Then
    lngMax = 700
  End If
  If lngMin < 400 Then
    lngMin = 400
  End If
  If Wid = -1 Then
    ctrl.Cls
    Wid = ctrl.ScaleWidth
   Else
    si = Wid - 1
  End If
  lngMin = lngMin - 400
  lngMax = lngMax - 400
  Range = lngMax - lngMin
  If bHorz Then
    sA = Wid / Range - 1
   Else
    sA = ctrl.ScaleHeight / (Range - 1)
  End If
  If bHorz Then
    ncol = -1
    For I = 0 To ctrl.ScaleWidth
      If I Mod sA = 0 Then
        ncol = ncol + 1
      End If
      If I = ctrl.ScaleWidth - 2 Then
        ncol = Range
      End If
      ctrl.Line (I, 0)-(I, ctrl.Height), SpectralColour(IIf(bDir, 400 + ncol, 700 - ncol))
    Next I
   Else
    For I = si To Wid
      If I Mod sA = 0 Then
        ncol = ncol + 1
      End If
      If I = ctrl.ScaleHeight - 2 Then
        ncol = Range
      End If
      ctrl.Line (si, I)-(Wid, I), SpectralColour(IIf(bDir, 400 + ncol, 700 - ncol))
    Next I
  End If

End Sub

Public Function SpectralColour(ByVal WavLen As Single) As Long

'WavLen must be between 400 & 700 to be Visible
'generate a Long colour based on Visible WaveLengths

  Dim R As Long
  Dim G As Long
  Dim B As Long

'if not set explicitly then default value is 0
  If WavLen > 399 Then
    If WavLen < 701 Then
      Select Case WavLen
       Case Is < 460
        B = 255
        R = CFactor * (460 - WavLen)
       Case Is < 520
        B = 255
        G = 255 - CFactor * (520 - WavLen)
       Case Is < 580
        G = 255
        B = CFactor * (580 - WavLen)
       Case Is < 640
        G = 255
        R = 255 - CFactor * (640 - WavLen)
       Case Is <= 700
        R = 255
        G = CFactor * (700 - WavLen)
      End Select
      SpectralColour = RGB(R, G, B)
    End If
  End If

End Function

Public Function SpectralValue(ByVal Col As Long) As Single

'converts colours to their spectral value (400- 700)
'Note doesn't detect non-spectrum colours properly;
'non-spectrum colour usually returned as purplish

  Dim R As Long
  Dim G As Long
  Dim B As Long

  R = CLng(Col) Mod 256
  B = (Col And &HFF0000) / 65536
  G = ((Col And &HFF00) / 256&) Mod 256&
  If B = 255 Then
    If G = 0 Then
      SpectralValue = 460 - VFactor * R
     ElseIf R = 0 Then
      SpectralValue = 460 + VFactor * G
    End If
   ElseIf G = 255 Then
    If R = 0 Then
      SpectralValue = 580 - VFactor * B
     ElseIf B = 0 Then
      SpectralValue = 580 + VFactor * R
    End If
   ElseIf R = 255 Then
    If B = 0 Then
      SpectralValue = 700 - VFactor * G
    End If
  End If

End Function

':)Code Fixer V4.0.0 (Tuesday, 05 July 2005 11:41:10) 6 + 132 = 138 Lines Thanks Ulli for inspiration and lots of code.
':)SETTINGS DUMP: 13330232222333323|333320222222222222222222222222|1112222|2221222|222222222233|1111111111111|1122222222222|333333|

