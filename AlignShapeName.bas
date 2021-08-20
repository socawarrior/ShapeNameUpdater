Attribute VB_Name = "Module1"
Sub AlignShapeName()
    Dim shp As Shape
    Dim sld As Slide
    

    For Each sld In ActivePresentation.Slides
    
        For Each shp In sld.Shapes
            If shp.Name = "Appendix Reference" Then
            shp.Name = shp.TextFrame.TextRange

            End If
        Next shp
    Next sld
End Sub
