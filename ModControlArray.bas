Attribute VB_Name = "ModControlArray"
Option Explicit
'this module contains several generic routintes for dealing with arrays of controls
'some are specialized to the PerifVis project (marked) but most could be taken for
'use in your own code. Feel free.
Private Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal Clr As Long, _
                                                          ByVal hpal As Long, _
                                                          ByRef lpcolorref As Long)

Public Sub ControlArrayCaptionWipe(CtrlA As Variant)

'wipe all captions

  Dim ctrl As Variant

  For Each ctrl In CtrlA
    ctrl.Caption = vbNullString
  Next ctrl

End Sub

Public Sub ControlArrayGridLoad(CtrlA As Variant, _
                                ByVal Horz As Long, _
                                ByVal Vert As Long, _
                                Optional ByVal bVisible As Boolean = False)

'Set the size of the root control before calling this
'as all the controls will take its size
'
'Note By default Loaded controls are Visible=False this routine presreves this
' but if you assign a value to the Optional parameter you can show controls immediately on creation
'left to right then top to bottom
'if you want top to bottom then left to right just reverse the For lines

  Dim I      As Long
  Dim J      As Long
  Dim CCount As Long

  On Error Resume Next
  For I = 0 To Horz
    For J = 0 To Vert
      If CCount > 0 Then
        Load CtrlA(CCount)
      End If
      CtrlA(CCount).Move I * CtrlA(0).Width, J * CtrlA(0).Height
      CCount = CCount + 1
      CtrlA(CCount).bVisible = bVisible
    Next J
  Next I
  On Error GoTo 0

End Sub

Public Sub ControlArrayHinting(CtrlA As Variant)

'specialized to the PerifVis project (could be adapted to others)
'call twice 1st to change colour of hint targets and 2nd to restore them

  Dim ctrl   As Variant
  Dim SysCol As Long

'System Colours present a special problem for the inversion code below(can't undo the hint) so get the correct colour just in case
  For Each ctrl In CtrlA
    If ctrl.BackColor < -1 Then 'system colours are negative values
      SysCol = ctrl.BackColor
'if a system color is found, save it. (always vbButtonFace in PerifVis)
      Exit For
    End If                  ' assumes that all controls are system colour
  Next ctrl
'Main loop
  For Each ctrl In CtrlA
    With ctrl
      If .Tag = "X" Then 'Marked as active
        If .BackColor < 0 Then
          If SysCol < -1 Then ' special to deal with the problem of hinting simpler games
            If .BackColor <> SysCol Then
              .BackColor = SysCol ' restore the discovered SysCol
             Else
              .BackColor = Abs(vbWhite - TranslateColor(.BackColor))
'usually black but for wierded Themes wh knows?
            End If
          End If
         Else
'standard colours dealt with here
          .BackColor = Abs(vbWhite - TranslateColor(.BackColor))
'cheap and nasty colour inversion
        End If
        .Refresh
      End If
    End With 'ctrl
  Next ctrl

End Sub

Public Sub controlArrayUnload(CtrlA As Variant)

'delete all controls in an array except for the root control (assumes its Index = 0)

  Static bWorking As Boolean

  Dim ctrl        As Variant
  If Not bWorking Then
    bWorking = True
    For Each ctrl In CtrlA
      If ctrl.Index <> 0 Then
        On Error Resume Next
        Unload ctrl
      End If
    Next ctrl
    bWorking = False
  End If
  On Error GoTo 0

End Sub

Public Sub ControlArrayVisible(CtrlA As Variant, _
                               ByVal bVisible As Boolean)

'hide/show all controls in the control array

  Dim ctrl As Variant

  For Each ctrl In CtrlA
    ctrl.Visible = bVisible
  Next ctrl

End Sub

Public Function TranslateColor(ByVal Clr As OLE_COLOR, _
                               Optional hpal As Long = 0) As Long

'translates System Colours to standard; no effect on standard colours

  OleTranslateColor Clr, hpal, TranslateColor 'Then

End Function

':)Code Fixer V4.0.0 (Tuesday, 05 July 2005 11:41:10) 5 + 127 = 132 Lines Thanks Ulli for inspiration and lots of code.
':)SETTINGS DUMP: 13330232222333323|333320222222222222222222222222|1112222|2221222|222222222233|1111111111111|1122222222222|333333|

