''
' Celebrate v1.0.0
' (c) Shannon Whitley - https://github.com/swhitley/Celebrate
'
' Generate slides for typical celebrations such as anniversaries and birthdays.
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Option Explicit
    Dim slideTitle As String
    Dim groupLabel As String
    Dim groupLabelZero As String
    Dim subColor As ColorFormat
Sub DataLoad()

    Dim slideItems() As String
    Dim data As String
    Dim shp As Shape
    Dim tbl As Table
    Dim btnRun As Shape
    Dim itemCount As Integer
    Dim slideNbr As Integer
    Dim group As String
    Dim prevGroup As String
    Dim people As Object
    Dim person As Object
    Dim title As String
  
 
    'Get the slide data from the slide notes
    data = ActivePresentation.Slides(1).NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
    
    'Loop through the shapes to gather input
    For Each shp In ActivePresentation.Slides(1).Shapes
        If shp.AlternativeText = "Run" Then
            Set btnRun = shp
        End If
        If shp.Type = msoTable Then
            Set tbl = shp.Table
        End If
    Next
    
    slideTitle = tbl.Cell(2, 2).Shape.TextFrame.TextRange.Text
    groupLabel = tbl.Cell(3, 2).Shape.TextFrame.TextRange.Text
    groupLabelZero = tbl.Cell(4, 2).Shape.TextFrame.TextRange.Text
    Set subColor = tbl.Cell(5, 2).Shape.TextFrame.TextRange.Font.Color
    
    btnRun.TextFrame.TextRange.Text = "Processing..."

    'Parse the json data.  See TODO: for the required data format.
    Set people = WebHelpers.ParseJson(data)
    On Error Resume Next
        Set people = people.Item("Report_Entry")
    Err.Number = 0
    
    itemCount = 1
    group = 1
    prevGroup = ""
    slideNbr = 1
    ReDim slideItems(1 To 8, 1 To 4)
    For Each person In people
        If prevGroup = "" Then
            prevGroup = person("group")
        End If
        If person("group") <> prevGroup Then
            SlideBuild slideItems, itemCount
            ReDim slideItems(1 To 8, 1 To 4)
            itemCount = 1
            slideNbr = slideNbr + 1
            prevGroup = person("group")
        End If
        If itemCount > 8 Then
            SlideBuild slideItems, itemCount
            ReDim slideItems(1 To 8, 1 To 4)
            itemCount = 1
            slideNbr = slideNbr + 1
        End If
        slideItems(itemCount, 1) = person("group")
        slideItems(itemCount, 2) = person("photo")
        slideItems(itemCount, 3) = person("name")
        title = person("title")
        'Limit the title to 35 characters
        If Len(title) > 35 Then
            title = left(title, 35) & "..."
            title = Replace(title, ",...", "...")
        End If
        slideItems(itemCount, 4) = title
        itemCount = itemCount + 1
        btnRun.TextFrame.TextRange.Text = "Adding Slide " & slideNbr
    Next person
    
    If slideItems(1, 1) <> "" Then
        SlideBuild slideItems, itemCount
    End If
    
   btnRun.TextFrame.TextRange.Text = "Run"
   MsgBox "Your slides have been created!", vbOKOnly

End Sub
Sub UpdateCheck()

    ActivePresentation.FollowHyperlink ""
    
End Sub
Sub SlideBuild(slideItems, itemCount)

    Dim layout(1 To 8) As String
    layout(8) = "1.19,2.12,5.36,2.12,9.7,2.12,13.97,2.12,1.19,6.13,5.36,6.13,9.7,6.13,13.97,6.13"
    layout(7) = "1.19,2.12,5.32,2.12,9.65,2.12,13.92,2.12,3.15,6.13,7.49,6.13,11.76,6.13"
    layout(6) = "1.19,2.12,5.38,2.12,9.72,2.12,13.99,2.12,5.34,6.13,9.67,6.13"
    layout(5) = "1.19,2.12,5.36,2.12,9.7,2.12,13.97,2.12,7.49,6.13"
    layout(4) = "1.19,3.84,5.34,3.84,9.68,3.84,13.95,3.84"
    layout(3) = "2.52,3.84,7.56,3.84,12.61,3.84"
    layout(2) = "5.44,3.84,9.57,3.84"
    layout(1) = "7.64,3.84"
    
    Dim positions() As String
    Dim position As Integer

    Dim sld As Slide
    Dim shp As Shape
    Dim subtitle As Shape
    Dim txtBox As Shape
    Dim subtitleText As String
    Dim group As String
    
    Dim ctr As Integer
    
    Dim imageFile As String
    Dim imageString As String
    Dim tmpFile As String
    tmpFile = ActivePresentation.Path & "/temp.jpg"
    
    'Add a new slide
    Set sld = ActivePresentation.Slides.AddSlide(ActivePresentation.Slides.Count + 1, ActivePresentation.Slides(1).CustomLayout)

    'Title
    sld.Shapes.title.TextFrame.TextRange.Text = slideTitle
    
    'Subtitle
    Set subtitle = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 65, 8, 800, 200)
    
    group = slideItems(1, 1)
    If group = "0" Then
        subtitleText = groupLabelZero
    Else
        If group = "1" Then
            subtitleText = group & " " & groupLabel
        Else
            subtitleText = group & " " & groupLabel & "s"
        End If
    End If
    
    subtitle.TextFrame.TextRange.Text = subtitleText
    subtitle.TextFrame.TextRange.Font.Size = 36
    subtitle.TextFrame.AutoSize = ppAutoSizeNone
    subtitle.TextFrame.TextRange.Font.Color = subColor
    
    positions = Split(layout(itemCount - 1), ",")
    position = 0
    
    For ctr = 1 To itemCount - 1
        'Add the Image
        'If photo contains colon or period, assume url or file path
        'otherwise, assume base64 encoding.
        If InStr(slideItems(ctr, 2), ":") > 0 Or InStr(slideItems(ctr, 2), ".") > 0 Then
            imageFile = slideItems(ctr, 2)
        Else
            imageString = WebHelpers.Base64Decode(slideItems(ctr, 2))
            Open tmpFile For Binary Access Write As #1
            Put #1, , imageString
            Close #1
            imageFile = tmpFile
        End If
        
        ' Oval with photo
        ' 72 points = 1 inch
        Set shp = sld.Shapes.AddPicture(imageFile, _
            msoFalse, msoCTrue, positions(position) * 72, positions(position + 1) * 72, 180, 180)
        shp.AutoShapeType = msoShapeOval
        shp.Line.BackColor.RGB = RGB(255, 255, 255)
        shp.Line.Weight = 6
        
        ' Textbox
        Set txtBox = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, positions(position) * 72, (positions(position + 1) * 72) + 190, 200, 200)
        txtBox.TextFrame.VerticalAnchor = msoAnchorTop
        txtBox.ScaleHeight 3, msoFalse, msoScaleFromTopLeft
        txtBox.ScaleWidth 1, msoFalse, msoScaleFromTopLeft
        txtBox.TextFrame.TextRange.Font.Size = 18
        txtBox.TextFrame.AutoSize = ppAutoSizeNone
        txtBox.TextFrame.TextRange.Text = slideItems(ctr, 3) & vbCr & slideItems(ctr, 4)
        position = position + 2
    Next

End Sub
