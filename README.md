# PPT

Sub CreateDeScamSlide()
    Dim pptApp As PowerPoint.Application
    Dim pptPres As PowerPoint.Presentation
    Dim pptSlide As PowerPoint.Slide
    Dim titleShape As PowerPoint.Shape
    Dim contentShapes As PowerPoint.Shapes
    Dim leftColumnWidth As Integer, rightColumnWidth As Integer
    Dim verticalSpacing As Integer
    
    ' Initialize parameters
    leftColumnWidth = 300
    rightColumnWidth = 300
    verticalSpacing = 80
    
    ' Create or reference PowerPoint
    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    On Error GoTo 0
    If pptApp Is Nothing Then
        Set pptApp = New PowerPoint.Application
        pptApp.Visible = True
    End If
    
    ' Create new presentation/slide
    Set pptPres = pptApp.Presentations.Add
    Set pptSlide = pptPres.Slides.Add(1, ppLayoutBlank)
    
    ' Add title
    Set titleShape = pptSlide.Shapes.AddTextbox(Orientation:=msoTextOrientationHorizontal, _
        Left:=50, Top:=20, Width:=600, Height:=60)
    titleShape.TextFrame.TextRange.Text = "DeScam - An Interactive Scam Detection Tool"
    titleShape.TextFrame.TextRange.Font.Size = 28
    titleShape.TextFrame.TextRange.Font.Bold = True
    
    ' Add subtitle
    pptSlide.Shapes.AddTextbox(msoTextOrientationHorizontal, 50, 70, 600, 30).TextFrame.TextRange.Text = _
        "First-Line Defense Against Multi-Channel Scams"
    
    ' ===== LEFT COLUMN CONTENT =====
    ' Section 1: What is a Scam?
    AddSection pptSlide, 50, 120, leftColumnWidth, 120, _
        "What is a Scam?", _
        Array( _
            "Fraudulent attempt to deceive individuals", _
            "Example: ""Urgent! Your bank account is frozen. Click [link] to verify""", _
            "(Phishing email with malicious link)" _
        )
    
    ' Section 2: Market Risk (Global)
    AddSection pptSlide, 50, 260, leftColumnWidth, 150, _
        "Market Risk (2023)", _
        Array( _
            "Global Losses: $10.3B+ (FBI IC3)", _
            "30% YoY increase in phishing", _
            "62% via email/SMS/calls", _
            "", _
            "[Internal Data Placeholder 1]", _
            "[Internal Data Placeholder 2]" _
        )
    
    ' ===== RIGHT COLUMN CONTENT =====
    ' Section 3: Problem Statement
    AddSection pptSlide, 400, 120, rightColumnWidth, 150, _
        "Problem Statement", _
        Array( _
            "Current tools lack:", _
            "- Real-time detection", _
            "- Cross-channel coverage", _
            "- Explainability", _
            "", _
            "DeScam provides:", _
            "1. Multi-channel detection", _
            "2. Plain-language explanations", _
            "3. Adaptive learning" _
        )
    
    ' Section 4: Why DeScam?
    AddSection pptSlide, 400, 300, rightColumnWidth, 120, _
        "Why DeScam?", _
        Array( _
            ChrW(10004) & " Multi-channel coverage", _
            ChrW(10004) & " Explainable AI", _
            ChrW(10004) & " Real-time alerts", _
            ChrW(10004) & " Dynamic learning" _
        )
    
    ' Format placeholders
    For Each shp In pptSlide.Shapes
        If InStr(shp.TextFrame.TextRange.Text, "[Internal Data]") > 0 Then
            shp.TextFrame.TextRange.Font.Color = RGB(255, 0, 0)
        End If
    Next
End Sub

' Helper function to create sections
Sub AddSection(slide As Slide, left As Integer, top As Integer, width As Integer, height As Integer, _
    heading As String, bullets As Variant)
    
    Dim shp As Shape
    Set shp = slide.Shapes.AddTextbox(msoTextOrientationHorizontal, left, top, width, height)
    
    With shp.TextFrame.TextRange
        .Text = heading & vbNewLine & Join(bullets, vbNewLine)
        .Font.Size = 14
        .Paragraphs(1).Font.Size = 18
        .Paragraphs(1).Font.Bold = True
        
        ' Add bullets for all lines except heading
        For i = 2 To .Paragraphs.Count
            .Paragraphs(i).IndentLevel = 1
            .Paragraphs(i).Font.Size = 12
        Next i
    End With
End Sub
