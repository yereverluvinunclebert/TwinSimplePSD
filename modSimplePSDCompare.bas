Attribute VB_Name = "modSimplePSDCompare"
Option Explicit

'==============================================================================
' Differential test harness for RC6.cSimplePSD vs the rebuilt class.
'
' V3 diagnostics add:
'   - composite comparison BEFORE any LayerSurface BindToArray calls
'   - first differing raw surface bytes, with local pixel/channel coordinates
'   - a second composite comparison after layer inspection, to detect any
'     accidental mutation caused by the diagnostic process itself
'==============================================================================

Private Const MAX_DIFF_DETAILS As Long = 8

Private mo_Log As RC6.cStringBuilder

Public Sub CompareSimplePSD(ByVal PSDFile As String, Optional ByVal OutputFolder As String = vbNullString)
   Dim RefPSD As RC6.cSimplePSD
   Dim TestPSD As cSimplePSD
   Dim I As Long
   Dim N As Long
   Dim RefS As cCairoSurface
   Dim TestS As cCairoSurface
   Dim Test2S As cCairoSurface
   Dim FreshCompositeDiff As Variant
   Dim PostInspectionDiff As Variant
   
   Set mo_Log = New_c.StringBuilder
   Set RefPSD = New_c.SimplePSD(PSDFile)

   Set TestPSD = New cSimplePSD
   TestPSD.Load PSDFile

   mo_Log.AppendNL String$(78, "=")
   mo_Log.AppendNL "PSD: " & vbTab & PSDFile
   mo_Log.AppendNL "Document: RC6=" & vbTab & RefPSD.Width & vbTab & "x" & vbTab & RefPSD.Height & vbTab & _
               "  rebuilt=" & vbTab & TestPSD.Width & vbTab & "x" & vbTab & TestPSD.Height
   mo_Log.AppendNL "Layers:   RC6=" & vbTab & RefPSD.LayersCount & vbTab & _
               "  rebuilt=" & vbTab & TestPSD.LayersCount

   'Do this BEFORE touching any individual LayerSurface buffers.  If this value
   'later differs from the post-inspection value, the diagnostics themselves
   'are affecting a surface and we need to know that.
   Set RefS = RefPSD.RenderAllLayers
   Set TestS = TestPSD.RenderAllLayers
   Set Test2S = RenderRC6LayersManually(RefPSD)

   If Len(OutputFolder) <> 0 Then
      If Not New_c.FSO.EnsurePath(OutputFolder) Then
         Err.Raise vbObjectError + 28761, "CompareSimplePSD", _
                   "Could not create comparison output folder: " & OutputFolder
      End If
      New_c.FSO.EnsurePathEndSep OutputFolder

      'Write before BindToArray diagnostics as another guard against the test
      'harness itself affecting pixels.
      RefS.WriteContentToPngFile OutputFolder & "RC6-cSimplePSD.png"
      TestS.WriteContentToPngFile OutputFolder & "Rebuilt-cSimplePSD.png"
      Test2S.WriteContentToPngFile OutputFolder & "RC6Manual-cSimplePSD.png"
   End If

   FreshCompositeDiff = SurfacePixelByteDifferences(RefS, TestS)
   mo_Log.AppendNL "Fresh composite pixel-byte differences: " & vbTab & FreshCompositeDiff

   N = RefPSD.LayersCount
   If TestPSD.LayersCount < N Then N = TestPSD.LayersCount

   For I = 0 To N - 1
      mo_Log.AppendNL String$(78, "-")
      mo_Log.AppendNL "Layer " & vbTab & I
      mo_Log.AppendNL "  Name:       " & vbTab & Q(RefPSD.LayerName(I)) & vbTab & " | " & vbTab & Q(TestPSD.LayerName(I))
      mo_Log.AppendNL "  Path:       " & vbTab & Q(RefPSD.LayerPath(I)) & vbTab & " | " & vbTab & Q(TestPSD.LayerPath(I))
      mo_Log.AppendNL "  XY:         " & vbTab & RefPSD.LayerX(I) & vbTab & "," & vbTab & RefPSD.LayerY(I) & vbTab & _
                  " | " & vbTab & TestPSD.LayerX(I) & vbTab & "," & vbTab & TestPSD.LayerY(I)
      mo_Log.AppendNL "  Size:       " & vbTab & RefPSD.LayerWidth(I) & vbTab & "x" & vbTab & RefPSD.LayerHeight(I) & vbTab & _
                  " | " & vbTab & TestPSD.LayerWidth(I) & vbTab & "x" & vbTab & TestPSD.LayerHeight(I)
      mo_Log.AppendNL "  Alpha %:    " & vbTab & RefPSD.LayerAlphaPercent(I) & vbTab & _
                  " | " & vbTab & TestPSD.LayerAlphaPercent(I)
      mo_Log.AppendNL "  Blend:      " & vbTab & Q(RefPSD.LayerBlendMode(I)) & vbTab & _
                  " | " & vbTab & Q(TestPSD.LayerBlendMode(I))
      mo_Log.AppendNL "  Operator:   " & vbTab & CLng(RefPSD.LayerBlendOperator(I)) & vbTab & _
                  " | " & vbTab & CLng(TestPSD.LayerBlendOperator(I))
      mo_Log.AppendNL "  Channels:   " & vbTab & RefPSD.LayerChannelCount(I) & vbTab & _
                  " | " & vbTab & TestPSD.LayerChannelCount(I)
      mo_Log.AppendNL "  Clipping:   " & vbTab & RefPSD.LayerClipping(I) & vbTab & _
                  " | " & vbTab & TestPSD.LayerClipping(I)
      mo_Log.AppendNL "  Flags:      " & vbTab & CLng(RefPSD.LayerFlags(I)) & vbTab & _
                  " | " & vbTab & CLng(TestPSD.LayerFlags(I))
      mo_Log.AppendNL "  ByteSize:   " & vbTab & RefPSD.LayerByteSize(I) & vbTab & _
                  " | " & vbTab & TestPSD.LayerByteSize(I)

      Set RefS = RefPSD.LayerSurface(I)
      Set TestS = TestPSD.LayerSurface(I)

      PrintSurfacePixelDifferences RefS, TestS, "  "
   Next I

   'Repeat after all LayerSurface BindToArray calls.
   Set RefS = RefPSD.RenderAllLayers
   Set TestS = TestPSD.RenderAllLayers
   PostInspectionDiff = SurfacePixelByteDifferences(RefS, TestS)

   mo_Log.AppendNL String$(78, "-")
   mo_Log.AppendNL "Composite after layer inspection: " & vbTab & PostInspectionDiff

   If CStr(FreshCompositeDiff) <> CStr(PostInspectionDiff) Then
      mo_Log.AppendNL "WARNING: composite changed after LayerSurface inspection."
   End If

   If Len(OutputFolder) <> 0 Then
      mo_Log.AppendNL "Wrote (before per-layer inspection):"
      mo_Log.AppendNL "  " & vbTab & OutputFolder & "RC6-cSimplePSD.png"
      mo_Log.AppendNL "  " & vbTab & OutputFolder & "Rebuilt-cSimplePSD.png"
   End If

   mo_Log.AppendNL String$(78, "=")
   Clipboard.Clear
   Clipboard.SetText mo_Log.ToString
End Sub

Private Sub PrintSurfacePixelDifferences(ByVal A As cCairoSurface, _
                                         ByVal B As cCairoSurface, _
                                         Optional ByVal Prefix As String = vbNullString)
   Dim BA() As Byte
   Dim BB() As Byte
   Dim I As Long
   Dim NA As Long
   Dim NB As Long
   Dim N As Long
   Dim Diffs As Double
   Dim DetailCount As Long
   Dim RowByte As Long
   Dim PX As Long
   Dim PY As Long
   Dim ChannelIndex As Long
   Dim ChannelName As String

   If A Is Nothing And B Is Nothing Then
      mo_Log.AppendNL Prefix & "Pixel diff: 0"
      Exit Sub
   End If

   If A Is Nothing Or B Is Nothing Then
      mo_Log.AppendNL Prefix & "Pixel diff: <one surface is Nothing>"
      Exit Sub
   End If

   If A.Width <> B.Width Or A.Height <> B.Height Or A.Stride <> B.Stride Then
      mo_Log.AppendNL Prefix & "Pixel diff: <surface geometry differs>"
      Exit Sub
   End If

   If Not A.BindToArray(BA, True) Then
      mo_Log.AppendNL Prefix & "Pixel diff: <could not bind RC6 surface>"
      Exit Sub
   End If

   If Not B.BindToArray(BB, True) Then
      A.ReleaseArray BA
      mo_Log.AppendNL Prefix & "Pixel diff: <could not bind rebuilt surface>"
      Exit Sub
   End If

   NA = UBound(BA) - LBound(BA) + 1
   NB = UBound(BB) - LBound(BB) + 1
   N = NA
   If NB < N Then N = NB

   For I = 0 To N - 1
      If BA(LBound(BA) + I) <> BB(LBound(BB) + I) Then
         Diffs = Diffs + 1#

         If DetailCount < MAX_DIFF_DETAILS Then
            PY = I \ A.Stride
            RowByte = I Mod A.Stride

            If RowByte < A.Width * 4 Then
               PX = RowByte \ 4
               ChannelIndex = RowByte And 3

               Select Case ChannelIndex
                  Case 0: ChannelName = "B"
                  Case 1: ChannelName = "G"
                  Case 2: ChannelName = "R"
                  Case 3: ChannelName = "A"
               End Select

               mo_Log.AppendNL Prefix & "  diff byte=" & CStr(I) & _
                           " pixel=(" & CStr(PX) & "," & CStr(PY) & ")" & _
                           " channel=" & ChannelName & _
                           " RC6=" & CStr(BA(LBound(BA) + I)) & _
                           " rebuilt=" & CStr(BB(LBound(BB) + I))
            Else
               mo_Log.AppendNL Prefix & "  diff byte=" & CStr(I) & _
                           " row=" & CStr(PY) & " padding+" & _
                           CStr(RowByte - A.Width * 4) & _
                           " RC6=" & CStr(BA(LBound(BA) + I)) & _
                           " rebuilt=" & CStr(BB(LBound(BB) + I))
            End If

            DetailCount = DetailCount + 1
         End If
      End If
   Next I

   If NA <> NB Then Diffs = Diffs + Abs(CDbl(NA) - CDbl(NB))

   A.ReleaseArray BA
   B.ReleaseArray BB

   mo_Log.AppendNL Prefix & "Pixel diff: " & CStr(Diffs)
End Sub

Private Function SurfacePixelByteDifferences(ByVal A As cCairoSurface, ByVal B As cCairoSurface) As Variant
   Dim BA() As Byte
   Dim BB() As Byte
   Dim I As Long
   Dim NA As Long
   Dim NB As Long
   Dim N As Long
   Dim Diffs As Double

   If A Is Nothing And B Is Nothing Then
      SurfacePixelByteDifferences = 0
      Exit Function
   End If

   If A Is Nothing Or B Is Nothing Then
      SurfacePixelByteDifferences = "<one surface is Nothing>"
      Exit Function
   End If

   If A.Width <> B.Width Or A.Height <> B.Height Or A.Stride <> B.Stride Then
      SurfacePixelByteDifferences = "<surface geometry differs>"
      Exit Function
   End If

   If Not A.BindToArray(BA, True) Then
      SurfacePixelByteDifferences = "<could not bind RC6 surface>"
      Exit Function
   End If

   If Not B.BindToArray(BB, True) Then
      A.ReleaseArray BA
      SurfacePixelByteDifferences = "<could not bind rebuilt surface>"
      Exit Function
   End If

   NA = UBound(BA) - LBound(BA) + 1
   NB = UBound(BB) - LBound(BB) + 1
   N = NA
   If NB < N Then N = NB

   For I = 0 To N - 1
      If BA(LBound(BA) + I) <> BB(LBound(BB) + I) Then
         Diffs = Diffs + 1#
      End If
   Next I

   If NA <> NB Then Diffs = Diffs + Abs(CDbl(NA) - CDbl(NB))

   A.ReleaseArray BA
   B.ReleaseArray BB

   SurfacePixelByteDifferences = Diffs
End Function

Private Function Q(ByVal S As String) As String
   Q = Chr$(34) & S & Chr$(34)
End Function


Private Function RenderRC6LayersManually(ByVal PSD As RC6.cSimplePSD) As cCairoSurface
   Dim S As cCairoSurface
   Dim CC As cCairoContext
   Dim LS As cCairoSurface
   Dim I As Long
   Dim A As Double

   Set S = Cairo.CreateSurface(PSD.Width, PSD.Height)
   Set CC = S.CreateContext

   CC.Operator = CAIRO_OPERATOR_CLEAR
   CC.Paint
   CC.Operator = CAIRO_OPERATOR_OVER

   For I = 0 To PSD.LayersCount - 1
      Set LS = PSD.LayerSurface(I)

      If Not LS Is Nothing Then
         If LS.Width > 0 And LS.Height > 0 Then
            'PSD raw flag bit 1 means invisible.
            If (PSD.LayerFlags(I) And &H2) = 0 Then
               A = PSD.LayerAlphaPercent(I)

               If A > 0 Then
                  CC.Operator = PSD.LayerBlendOperator(I)
                  CC.SetSourceSurface LS, PSD.LayerX(I), PSD.LayerY(I)
                  CC.Paint A
               End If
            End If
         End If
      End If
   Next I

   Set RenderRC6LayersManually = S
End Function

