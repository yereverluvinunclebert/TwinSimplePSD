
# TwinSimplePSD
A direct Replacement for RichClient's Simple PSD Parser for VB6 and TwinBasic. A clean-room, source-compatible reimplementation of RC6.cSimplePSD
works with RC5 and 6, replacing Olaf's SimplePSD parser. Will compile to 32bit or 64bit with TwinBasic. Puts your own UI design extracted directly 
from a Photoshop PSD onto the desktop.

<img width="380" height="348" alt="cpu-gauge" src="https://github.com/user-attachments/assets/e6f1211f-2fd0-4f09-adcd-058be73220e1" />

**Author**

ChatGPT at JBPro's insistence. Inspired by Olaf Schmidt's closed source SimplePSD parser in RC6.

**Intent:**

A deliberately small PSD reader for the same sort of use as RC6.cSimplePSD: extracting ordinary raster layers into cCairoSurface objects and rendering
the visible layer stack with Cairo.

**Supported subset:**

* PSD version 1 (not PSB)
* RGB, 8 bits/channel
* Per-layer R/G/B, transparency and user-mask channels
* RAW, PackBits/RLE, ZIP and ZIP-with-prediction channel compression
* Pascal and Unicode ("luni") layer names
* Photoshop folder/group markers ("lsct"/"lsdk")
* Common Photoshop blend modes that have direct Cairo equivalents

**Deliberately not interpreted:** 

* Text/vector/smart-object rendering
* Adjustment/fill layer semantics
* Layer effects/styles
* Vector masks
* Clipping groups
* Group isolation/knockout and exact Photoshop group-opacity semantics
* CMYK/Lab/Indexed/16-bit/32-bit/PSB

**RichClient is still used within this class for the following:**
  
* zlib           New_c.Crypt.ZLibDecompress
* Cairo surfaces Cairo.CreateSurface / cCairoSurface.BindToArray
* alpha          Cairo.PreMultiplyAlpha
* compositing    cCairoContext

No Win32 API declarations are used.

**Performance notes (V4):**
  
* large byte-array copies use RichClient New_c.MemCopy where practical
* zero-based Variant Byte() input is assigned by the VB runtime directly
* RLE decode loops use linear source/destination positions
* common alpha/no-mask surface packing uses a branch-free linear loop

**Files:**

    cSimplePSD.cls - the Simple Parser class
    modSimplePSDFactory.bas - Factory helper for the clean-room cSimplePSD class.
    CSimplePsdReimpl.vbp - the project file for VB6
    MTestSimplePSDPerformance.bas - Optional Performance tester, not needed in the project
    modSimplePSDCompare.bas - Optional Comparison Tester, not needed in the project

<img width="1078" height="1041" alt="pzCPURC50004" src="https://github.com/user-attachments/assets/637681f5-014b-4d56-93b5-83a5d51cc21c" />
Fig 01. The Panzer CPU gauge RC5/6 widget created using RC6 and this SimplePSD parser.

.
.
.

-o0o-

**Example Usage:**

If you look at this gitlab repo. https://gitlab.com/yereverluvinunclebert/panzer-cpu-gauge-rc5-vb6 , the simplePSD parser is used to render the components 
to a transparent RC form.

Create a design in Photoshop. Save each image element as a simple separate flattened layer, no effects. Place the resulting PSD in a folder that the program can access. 
Use SimplePSD to extract the layers to a collection. SimplePSD will create a Cairo surface and a widget instance for all layers in the PSD each of which that can have properties and events assigned.

**Example 1.**

Private Sub Form_Load()
    Const PSDFile As String = "c:\temp\your_file.psd"
    Dim PSD As cSimplePSD, i As Long
    Dim CC As cCairoContext
    
    ' Create the PSD-parser instance pointing to your file
    Set PSD = New_c.SimplePSD(PSDFile) 
    
    ' Loop over all layers inside the PSD
    For i = 0 To PSD.LayersCount - 1 
        Debug.Print PSD.LayerAlphaPercent(i), PSD.LayerPath(i) 
    Next i
    
    ' Write all stacked layers as a combined alpha-transparent .png
    PSD.AllLayersSurface.WriteContentToPngFile PSDFile & ".png" 
End Sub



**Example 2.**

    'create the Top-Level-Form
    Set gaugeForm = Cairo.WidgetForms.Create(IIf(App.LogMode, AlphaNoTaskbarEntry, AlphaWithTaskbarEntry), gsWidgetName, True, 1, 1)
        gaugeForm.WidgetRoot.BackColor = -1 ' transparent

    'With New_c.SimplePSD(PSD_FileNameOrByteArray)  'create a new PSD-Parser.instance (and load the passed content) only available in RC6
    
    '  Here we use the FOSS/ChatGPT simple PSD parser (replica of Olaf's RC6 SimplePSD parser)
    With SimplePSD(PSD_FileNameOrByteArray)  'create a new PSD-Parser.instance (and load the passed content) available for both RC5 and RC6
        pPSDWidth = .Width
        pPSDHeight = .Height       'store the original Psd-Pixel-Width/Height in Private vars (as the base from which we calculate the zoomed Form-Width/Height)

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

                        Set W = gaugeForm.Widgets.Add(New cwAlphaImg, LCase$(.LayerPath(I)), .LayerX(I), .LayerY(I), .LayerWidth(I), .LayerHeight(I)).Widget

                        W.ImageKey = W.Key 'W.Key equals ImageList-Key, set above - and LayerPath(i) at this point ... set it also as the ImageKey of our new created Widget

                        W.Alpha = 0

                        ' note: the clickable layers characteristics are set in adjustMainControls

                        ' all non-clickable Layer-Widgets will be -1 or "non-hoverable" and "fully click-through"
                        W.HoverColor = -1 ' task: might change this later when new ver or RC6 arrives
                        If gsWidgetTooltips = "1" Then W.ToolTip = "Ctrl + mouse scrollwheel up/down to resize, you can also drag me to a new position."
                        W.MousePointer = IDC_SIZEALL

                    End If
                End If
            End If
        Next I
    End With '<-- the Parser-instance will be destroyed here (freeing the Memory, the internal PSD-Layers have occupied)
