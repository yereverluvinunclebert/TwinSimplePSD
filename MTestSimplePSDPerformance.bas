Attribute VB_Name = "MTestSimplePSDPerformance"
Option Explicit

'===============================================================================
' MTestSimplePSDPerformance
'
' Compiled performance comparison between:
'
'   1. RC6.cSimplePSD
'   2. The clean-room cSimplePSD reimplementation
'
' The PSD file is read into memory once before any timing starts.  This keeps
' disk I/O and Windows file-cache effects out of the comparison and measures the
' PSD implementations themselves.
'
' Two benchmarks are performed:
'
'   Load + Composite
'       Creates a fresh PSD object from the in-memory PSD bytes, obtains
'       AllLayersSurface, touches the result, then destroys the object.
'
'   RenderAllLayers
'       Loads one PSD object for each implementation before timing and repeatedly
'       calls RenderAllLayers.  This isolates the compositing portion as much as
'       possible from parsing/decompression/layer-surface construction.
'
' Each benchmark:
'
'   - warms both implementations before timing
'   - automatically chooses an iteration count that gives reasonably long runs
'   - performs several rounds
'   - alternates which implementation runs first in each round
'   - uses QueryPerformanceCounter for high-resolution timing
'
' For the fairest results:
'
'   - Compile to Native Code.
'   - Use the same Advanced Optimizations you normally ship with.
'   - Run the compiled EXE, not the IDE.
'   - Avoid other CPU-heavy activity during the benchmark.
'   - Ideally run it more than once and compare the results.
'
' Startup Object:
'
'   Set the project's Startup Object to "Sub Main".
'
' PSD file:
'
'   By default the test expects:
'
'       <EXE folder>\tank-clock-mk1.psd
'
'   Alternatively, pass a PSD filename on the EXE command line.  Quoted paths
'   are accepted.
'===============================================================================

Private Declare Function QueryPerformanceCounter Lib "kernel32" ( _
      ByRef lpPerformanceCount As Currency) As Long

Private Declare Function QueryPerformanceFrequency Lib "kernel32" ( _
      ByRef lpFrequency As Currency) As Long

Private Const mc_TargetSecondsPerRound As Double = 1#
Private Const mc_CalibrationSeconds As Double = 0.2
Private Const mc_Rounds As Long = 5
Private Const mc_WarmupIterations As Long = 2
Private Const mc_MaxIterations As Long = 1000

Private m_CounterFrequency As Currency
Private m_BenchmarkSink As Long

Public Sub Main()
   On Error GoTo ErrorHandler

   Dim l_PSDFile As String
   Dim la_PSDData() As Byte
   Dim lv_PSDData As Variant
   Dim l_Result As String

   InitializePerformanceCounter

   l_PSDFile = PSDFileNameFromCommandLine()

   ' Read the PSD exactly once.  File I/O is not part of the timed work.
   la_PSDData = New_c.FSO.ReadByteContent(l_PSDFile)
   lv_PSDData = la_PSDData

   WarmUpLoadBenchmark lv_PSDData

   l_Result = "cSimplePSD PERFORMANCE TEST" & vbCrLf & _
              String$(34, "-") & vbCrLf & _
              "PSD: " & l_PSDFile & vbCrLf & _
              "Size: " & Format$(ByteArrayLength(la_PSDData), "#,##0") & _
              " bytes" & vbCrLf & _
              "Mode: compiled, in-memory PSD bytes" & vbCrLf & vbCrLf

   l_Result = l_Result & BenchmarkLoadAndComposite(lv_PSDData)

   l_Result = l_Result & vbCrLf & vbCrLf & _
              BenchmarkRenderAllLayers(lv_PSDData)

   l_Result = l_Result & vbCrLf & vbCrLf & _
              "Sink: " & CStr(m_BenchmarkSink)

   MsgBox l_Result, vbInformation Or vbOKOnly, _
          "RC6 vs Reimplemented cSimplePSD"

   Exit Sub

ErrorHandler:
   MsgBox "Performance test failed." & vbCrLf & vbCrLf & _
          "Error " & CStr(Err.Number) & vbCrLf & _
          Err.Source & vbCrLf & _
          Err.Description, _
          vbCritical Or vbOKOnly, _
          "cSimplePSD Performance Test"
End Sub

Private Function BenchmarkLoadAndComposite( _
      ByRef pv_PSDData As Variant) As String

   Dim l_Iterations As Long
   Dim l_Round As Long
   Dim l_Operations As Long
   Dim l_RC6Seconds As Double
   Dim l_RebuiltSeconds As Double
   Dim l_RC6Round As Double
   Dim l_RebuiltRound As Double

   l_Iterations = CalibrateLoadIterations(pv_PSDData)
   l_Operations = l_Iterations * mc_Rounds

   For l_Round = 1 To mc_Rounds
      If (l_Round And 1) <> 0 Then
         l_RC6Round = TimeRC6Load(pv_PSDData, l_Iterations)
         l_RebuiltRound = TimeRebuiltLoad(pv_PSDData, l_Iterations)
      Else
         l_RebuiltRound = TimeRebuiltLoad(pv_PSDData, l_Iterations)
         l_RC6Round = TimeRC6Load(pv_PSDData, l_Iterations)
      End If

      l_RC6Seconds = l_RC6Seconds + l_RC6Round
      l_RebuiltSeconds = l_RebuiltSeconds + l_RebuiltRound
   Next l_Round

   BenchmarkLoadAndComposite = _
      BuildResultSection( _
         "LOAD + COMPOSITE", _
         "load", _
         l_Operations, _
         l_RC6Seconds, _
         l_RebuiltSeconds)
End Function

Private Function BenchmarkRenderAllLayers( _
      ByRef pv_PSDData As Variant) As String

   Dim lo_RC6PSD As RC6.cSimplePSD
   Dim lo_RebuiltPSD As cSimplePSD
   Dim l_Iterations As Long
   Dim l_Round As Long
   Dim l_Operations As Long
   Dim l_RC6Seconds As Double
   Dim l_RebuiltSeconds As Double
   Dim l_RC6Round As Double
   Dim l_RebuiltRound As Double

   ' Parsing/layer construction occurs before the timed render-only test.
   Set lo_RC6PSD = New_c.SimplePSD(pv_PSDData)

   Set lo_RebuiltPSD = New cSimplePSD
   lo_RebuiltPSD.Load pv_PSDData

   WarmUpRenderBenchmark lo_RC6PSD, lo_RebuiltPSD

   l_Iterations = CalibrateRenderIterations(lo_RC6PSD, lo_RebuiltPSD)
   l_Operations = l_Iterations * mc_Rounds

   For l_Round = 1 To mc_Rounds
      If (l_Round And 1) <> 0 Then
         l_RC6Round = TimeRC6Render(lo_RC6PSD, l_Iterations)
         l_RebuiltRound = TimeRebuiltRender(lo_RebuiltPSD, l_Iterations)
      Else
         l_RebuiltRound = TimeRebuiltRender(lo_RebuiltPSD, l_Iterations)
         l_RC6Round = TimeRC6Render(lo_RC6PSD, l_Iterations)
      End If

      l_RC6Seconds = l_RC6Seconds + l_RC6Round
      l_RebuiltSeconds = l_RebuiltSeconds + l_RebuiltRound
   Next l_Round

   BenchmarkRenderAllLayers = _
      BuildResultSection( _
         "RENDERALLLAYERS ONLY", _
         "render", _
         l_Operations, _
         l_RC6Seconds, _
         l_RebuiltSeconds)

   Set lo_RebuiltPSD = Nothing
   Set lo_RC6PSD = Nothing
End Function

Private Sub WarmUpLoadBenchmark(ByRef pv_PSDData As Variant)
   Dim l_Index As Long

   For l_Index = 1 To mc_WarmupIterations
      RunOneRC6Load pv_PSDData
      RunOneRebuiltLoad pv_PSDData
   Next l_Index
End Sub

Private Sub WarmUpRenderBenchmark( _
      ByVal po_RC6PSD As RC6.cSimplePSD, _
      ByVal po_RebuiltPSD As cSimplePSD)

   Dim lo_Surface As RC6.cCairoSurface
   Dim l_Index As Long

   For l_Index = 1 To mc_WarmupIterations
      Set lo_Surface = po_RC6PSD.RenderAllLayers
      ConsumeSurface lo_Surface
      Set lo_Surface = Nothing

      Set lo_Surface = po_RebuiltPSD.RenderAllLayers
      ConsumeSurface lo_Surface
      Set lo_Surface = Nothing
   Next l_Index
End Sub

Private Function CalibrateLoadIterations( _
      ByRef pv_PSDData As Variant) As Long

   Dim l_Iterations As Long
   Dim l_RC6Seconds As Double
   Dim l_RebuiltSeconds As Double
   Dim l_SlowerPerIteration As Double
   Dim l_Estimated As Double

   l_Iterations = 1

   Do
      l_RC6Seconds = TimeRC6Load(pv_PSDData, l_Iterations)
      l_RebuiltSeconds = TimeRebuiltLoad(pv_PSDData, l_Iterations)

      If l_RC6Seconds >= mc_CalibrationSeconds And _
         l_RebuiltSeconds >= mc_CalibrationSeconds Then
         Exit Do
      End If

      If l_Iterations >= mc_MaxIterations Then Exit Do

      If l_Iterations > mc_MaxIterations \ 2 Then
         l_Iterations = mc_MaxIterations
      Else
         l_Iterations = l_Iterations * 2
      End If
   Loop

   l_SlowerPerIteration = _
      MaxDouble( _
         l_RC6Seconds / CDbl(l_Iterations), _
         l_RebuiltSeconds / CDbl(l_Iterations))

   If l_SlowerPerIteration <= 0# Then
      CalibrateLoadIterations = l_Iterations
      Exit Function
   End If

   l_Estimated = mc_TargetSecondsPerRound / l_SlowerPerIteration
   CalibrateLoadIterations = ClampIterations(l_Estimated)
End Function

Private Function CalibrateRenderIterations( _
      ByVal po_RC6PSD As RC6.cSimplePSD, _
      ByVal po_RebuiltPSD As cSimplePSD) As Long

   Dim l_Iterations As Long
   Dim l_RC6Seconds As Double
   Dim l_RebuiltSeconds As Double
   Dim l_SlowerPerIteration As Double
   Dim l_Estimated As Double

   l_Iterations = 1

   Do
      l_RC6Seconds = TimeRC6Render(po_RC6PSD, l_Iterations)
      l_RebuiltSeconds = TimeRebuiltRender(po_RebuiltPSD, l_Iterations)

      If l_RC6Seconds >= mc_CalibrationSeconds And _
         l_RebuiltSeconds >= mc_CalibrationSeconds Then
         Exit Do
      End If

      If l_Iterations >= mc_MaxIterations Then Exit Do

      If l_Iterations > mc_MaxIterations \ 2 Then
         l_Iterations = mc_MaxIterations
      Else
         l_Iterations = l_Iterations * 2
      End If
   Loop

   l_SlowerPerIteration = _
      MaxDouble( _
         l_RC6Seconds / CDbl(l_Iterations), _
         l_RebuiltSeconds / CDbl(l_Iterations))

   If l_SlowerPerIteration <= 0# Then
      CalibrateRenderIterations = l_Iterations
      Exit Function
   End If

   l_Estimated = mc_TargetSecondsPerRound / l_SlowerPerIteration
   CalibrateRenderIterations = ClampIterations(l_Estimated)
End Function

Private Function TimeRC6Load( _
      ByRef pv_PSDData As Variant, _
      ByVal p_Iterations As Long) As Double

   Dim l_Start As Currency
   Dim l_Finish As Currency
   Dim l_Index As Long

   CounterNow l_Start

   For l_Index = 1 To p_Iterations
      RunOneRC6Load pv_PSDData
   Next l_Index

   CounterNow l_Finish
   TimeRC6Load = SecondsElapsed(l_Start, l_Finish)
End Function

Private Function TimeRebuiltLoad( _
      ByRef pv_PSDData As Variant, _
      ByVal p_Iterations As Long) As Double

   Dim l_Start As Currency
   Dim l_Finish As Currency
   Dim l_Index As Long

   CounterNow l_Start

   For l_Index = 1 To p_Iterations
      RunOneRebuiltLoad pv_PSDData
   Next l_Index

   CounterNow l_Finish
   TimeRebuiltLoad = SecondsElapsed(l_Start, l_Finish)
End Function

Private Sub RunOneRC6Load(ByRef pv_PSDData As Variant)
   Dim lo_PSD As RC6.cSimplePSD
   Dim lo_Surface As RC6.cCairoSurface

   Set lo_PSD = New_c.SimplePSD(pv_PSDData)
   Set lo_Surface = lo_PSD.AllLayersSurface

   m_BenchmarkSink = m_BenchmarkSink Xor _
                     lo_PSD.Width Xor _
                     lo_PSD.Height Xor _
                     lo_PSD.LayersCount

   ConsumeSurface lo_Surface

   Set lo_Surface = Nothing
   Set lo_PSD = Nothing
End Sub

Private Sub RunOneRebuiltLoad(ByRef pv_PSDData As Variant)
   Dim lo_PSD As cSimplePSD
   Dim lo_Surface As RC6.cCairoSurface

   Set lo_PSD = New cSimplePSD
   lo_PSD.Load pv_PSDData
   Set lo_Surface = lo_PSD.AllLayersSurface

   m_BenchmarkSink = m_BenchmarkSink Xor _
                     lo_PSD.Width Xor _
                     lo_PSD.Height Xor _
                     lo_PSD.LayersCount

   ConsumeSurface lo_Surface

   Set lo_Surface = Nothing
   Set lo_PSD = Nothing
End Sub

Private Function TimeRC6Render( _
      ByVal po_PSD As RC6.cSimplePSD, _
      ByVal p_Iterations As Long) As Double

   Dim lo_Surface As RC6.cCairoSurface
   Dim l_Start As Currency
   Dim l_Finish As Currency
   Dim l_Index As Long

   CounterNow l_Start

   For l_Index = 1 To p_Iterations
      Set lo_Surface = po_PSD.RenderAllLayers
      ConsumeSurface lo_Surface
      Set lo_Surface = Nothing
   Next l_Index

   CounterNow l_Finish
   TimeRC6Render = SecondsElapsed(l_Start, l_Finish)
End Function

Private Function TimeRebuiltRender( _
      ByVal po_PSD As cSimplePSD, _
      ByVal p_Iterations As Long) As Double

   Dim lo_Surface As RC6.cCairoSurface
   Dim l_Start As Currency
   Dim l_Finish As Currency
   Dim l_Index As Long

   CounterNow l_Start

   For l_Index = 1 To p_Iterations
      Set lo_Surface = po_PSD.RenderAllLayers
      ConsumeSurface lo_Surface
      Set lo_Surface = Nothing
   Next l_Index

   CounterNow l_Finish
   TimeRebuiltRender = SecondsElapsed(l_Start, l_Finish)
End Function

Private Sub ConsumeSurface(ByVal po_Surface As RC6.cCairoSurface)
   If po_Surface Is Nothing Then
      m_BenchmarkSink = m_BenchmarkSink Xor &H13579BDF
   Else
      m_BenchmarkSink = m_BenchmarkSink Xor _
                        po_Surface.Width Xor _
                        po_Surface.Height Xor _
                        po_Surface.DataPtr
   End If
End Sub

Private Function BuildResultSection( _
      ByVal p_Title As String, _
      ByVal p_OperationName As String, _
      ByVal p_Operations As Long, _
      ByVal p_RC6Seconds As Double, _
      ByVal p_RebuiltSeconds As Double) As String

   Dim l_RC6Ms As Double
   Dim l_RebuiltMs As Double
   Dim l_RC6PerSecond As Double
   Dim l_RebuiltPerSecond As Double
   Dim l_Ratio As Double

   If p_Operations <= 0 Then
      Err.Raise 5, "BuildResultSection", _
                "Operations must be greater than zero."
   End If

   l_RC6Ms = (p_RC6Seconds / CDbl(p_Operations)) * 1000#
   l_RebuiltMs = (p_RebuiltSeconds / CDbl(p_Operations)) * 1000#

   If p_RC6Seconds > 0# Then
      l_RC6PerSecond = CDbl(p_Operations) / p_RC6Seconds
   End If

   If p_RebuiltSeconds > 0# Then
      l_RebuiltPerSecond = CDbl(p_Operations) / p_RebuiltSeconds
   End If

   If l_RebuiltMs > 0# Then l_Ratio = l_RC6Ms / l_RebuiltMs

   BuildResultSection = _
      p_Title & vbCrLf & _
      String$(Len(p_Title), "-") & vbCrLf & _
      "Iterations: " & Format$(p_Operations, "#,##0") & _
      " (" & CStr(mc_Rounds) & " rounds)" & vbCrLf & _
      "RC6:       " & _
         Format$(l_RC6Ms, "0.000") & " ms/" & p_OperationName & ", " & _
         Format$(l_RC6PerSecond, "#,##0.00") & "/sec" & vbCrLf & _
      "Rebuilt:   " & _
         Format$(l_RebuiltMs, "0.000") & " ms/" & p_OperationName & ", " & _
         Format$(l_RebuiltPerSecond, "#,##0.00") & "/sec" & vbCrLf & _
      RelativePerformanceText(l_Ratio)
End Function

Private Function RelativePerformanceText(ByVal p_RC6OverRebuilt As Double) As String
   If p_RC6OverRebuilt <= 0# Then
      RelativePerformanceText = "Relative: unable to calculate"
   ElseIf Abs(p_RC6OverRebuilt - 1#) < 0.005 Then
      RelativePerformanceText = _
         "Relative: effectively equal (" & _
         Format$(p_RC6OverRebuilt, "0.000") & "x)"
   ElseIf p_RC6OverRebuilt > 1# Then
      RelativePerformanceText = _
         "Relative: rebuilt is " & _
         Format$(p_RC6OverRebuilt, "0.00") & "x faster"
   Else
      RelativePerformanceText = _
         "Relative: RC6 is " & _
         Format$(1# / p_RC6OverRebuilt, "0.00") & "x faster"
   End If
End Function

Private Function ClampIterations(ByVal p_Estimated As Double) As Long
   If p_Estimated < 1# Then
      ClampIterations = 1
   ElseIf p_Estimated > CDbl(mc_MaxIterations) Then
      ClampIterations = mc_MaxIterations
   Else
      ClampIterations = CLng(p_Estimated)
      If ClampIterations < 1 Then ClampIterations = 1
   End If
End Function

Private Function MaxDouble( _
      ByVal p_A As Double, _
      ByVal p_B As Double) As Double

   If p_A >= p_B Then
      MaxDouble = p_A
   Else
      MaxDouble = p_B
   End If
End Function

Private Sub InitializePerformanceCounter()
   If QueryPerformanceFrequency(m_CounterFrequency) = 0 Then
      Err.Raise vbObjectError + 4900, _
                "MTestSimplePSDPerformance", _
                "QueryPerformanceFrequency is not available."
   End If
End Sub

Private Sub CounterNow(ByRef p_Value As Currency)
   If QueryPerformanceCounter(p_Value) = 0 Then
      Err.Raise vbObjectError + 4901, _
                "MTestSimplePSDPerformance", _
                "QueryPerformanceCounter failed."
   End If
End Sub

Private Function SecondsElapsed( _
      ByVal p_Start As Currency, _
      ByVal p_Finish As Currency) As Double

   SecondsElapsed = _
      CDbl(p_Finish - p_Start) / CDbl(m_CounterFrequency)
End Function

Private Function PSDFileNameFromCommandLine() As String
   Dim l_Command As String
   Dim l_Sep As String

   l_Command = Trim$(Command$)

   If Len(l_Command) > 0 Then
      If Len(l_Command) >= 2 Then
         If Left$(l_Command, 1) = """" And _
            Right$(l_Command, 1) = """" Then

            l_Command = Mid$(l_Command, 2, Len(l_Command) - 2)
         End If
      End If

      PSDFileNameFromCommandLine = l_Command
      Exit Function
   End If

   If Right$(App.Path, 1) = "\" Then
      l_Sep = vbNullString
   Else
      l_Sep = "\"
   End If

   PSDFileNameFromCommandLine = _
      App.Path & l_Sep & "tank-clock-mk1.psd"
End Function

Private Function ByteArrayLength(ByRef pa_Data() As Byte) As Long
   On Error GoTo Unallocated

   ByteArrayLength = UBound(pa_Data) - LBound(pa_Data) + 1
   Exit Function

Unallocated:
   ByteArrayLength = 0
End Function


