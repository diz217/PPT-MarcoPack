Option Explicit
 
Public Sub CropII()
    Dim sel As Selection
    Dim shpc As Shape
    Dim shporig As Shape
    Dim sld As Slide
    Dim shp As Shape
    Dim picCount As Long
    Set sel = ActiveWindow.Selection
    Set sld = ActiveWindow.View.Slide
 
    Select Case sel.Type
        Case ppSelectionShapes
            Set shporig = sel.ShapeRange(1)
            If shporig.Type <> msoPicture And shporig.Type <> msoLinkedPicture Then
                MsgBox "Not a picture", vbExclamation
                Exit Sub
            End If
         Case ppSelectionNone, ppSelectionSlides
            picCount = 0
            For Each shp In sld.Shapes
                If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
                    picCount = picCount + 1
                    Set shporig = shp
                End If
            Next shp
            If picCount = 0 Then
                MsgBox "No pictures on slide", vbExclamation
                Exit Sub
            ElseIf picCount > 1 Then
                MsgBox "More than one pictures in this slide, select one", vbExclamation
                Exit Sub
            End If
         Case Else
            MsgBox "Not a picture", vbExclamation
            Exit Sub
    End Select
    Set shpc = shporig.Duplicate(1)
    shpc.Left = shporig.Left
    shpc.Top = shporig.Top
    shpc.Select
    CommandBars.ExecuteMso "PictureCrop"
End Sub
Public Sub AntiCrop()
    Dim sel As Selection
    Dim sld As Slide
    Dim shp As Shape
    Dim target As Shape
    Dim picCount As Long
    Set sel = ActiveWindow.Selection
    Set sld = ActiveWindow.View.Slide
    Select Case sel.Type
        Case ppSelectionShapes
            Set target = sel.ShapeRange(1)
            If target.Type <> msoPicture And target.Type <> msoLinkedPicture Then
                MsgBox "Not a picture", vbExclamation
                Exit Sub
            End If
         Case ppSelectionNone, ppSelectionSlides
            picCount = 0
            For Each shp In sld.Shapes
                If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
                    picCount = picCount + 1
                    Set target = shp
                End If
            Next shp
            If picCount = 0 Then
                MsgBox "No pictures on slide", vbExclamation
                Exit Sub
            ElseIf picCount > 1 Then
                MsgBox "More than one pictures in this slide, select one", vbExclamation
                Exit Sub
            End If
        Case Else
            MsgBox "Not a picture", vbExclamation
            Exit Sub
    End Select
    On Error Resume Next
    With target.PictureFormat
        .CropLeft = 0
        .CropRight = 0
        .CropTop = 0
        .CropBottom = 0
    End With
    On Error GoTo 0
End Sub
