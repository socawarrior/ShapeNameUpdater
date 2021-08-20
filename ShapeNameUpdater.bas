Attribute VB_Name = "Module1"
Sub ShapeNameUpdater()
    Dim shp As Shape
    Dim sld As Slide
    

    For Each sld In ActivePresentation.Slides
    
        For Each shp In sld.Shapes
            If shp.Name = "Text Box 5" Then
            shp.Name = "Appendix Reference"

            End If
        Next shp
    Next sld
End Sub
