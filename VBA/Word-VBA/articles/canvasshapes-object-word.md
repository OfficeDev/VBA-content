---
title: CanvasShapes Object (Word)
keywords: vbawd10.chm115
f1_keywords:
- vbawd10.chm115
ms.prod: word
api_name:
- Word.CanvasShapes
ms.assetid: f4b37915-7fde-2a21-0df0-fc3c97983900
ms.date: 06/08/2017
---


# CanvasShapes Object (Word)

Use the  **CanvasItems** property of either a **[Shape](shape-object-word.md)** or **[ShapeRange](shaperange-object-word.md)** object to return a **CanvasShapes** collection.


## Remarks

To add shapes to a drawing canvas, use the following methods of the  **CanvasShapes** collection: **[AddCallout](canvasshapes-addcallout-method-word.md)** , **[AddConnector](canvasshapes-addconnector-method-word.md)** **[AddCurve](canvasshapes-addcurve-method-word.md)** , **[AddLabel](canvasshapes-addlabel-method-word.md)** , **[AddLine](canvasshapes-addline-method-word.md)** , **[AddPicture](canvasshapes-addpicture-method-word.md)** , **[AddPolyline](canvasshapes-addpolyline-method-word.md)** , **[AddShape](canvasshapes-addshape-method-word.md)** , **[AddTextbox](canvasshapes-addtextbox-method-word.md)** , **[AddTextEffect](canvasshapes-addtexteffect-method-word.md)** , or **[BuildFreeform](canvasshapes-buildfreeform-method-word.md)** . The following example adds a drawing canvas to the active document and then adds three shapes to the drawing canvas.


```vb
Sub AddCanvasShapes() 
 Dim shpCanvas As Shape 
 Dim shpCanvasShapes As CanvasShapes 
 Dim shpCnvItem As Shape 
 
 'Adds a new canvas to the document 
 Set shpCanvas = ActiveDocument.Shapes _ 
 .AddCanvas(Left:=100, Top:=75, _ 
 Width:=50, Height:=75) 
 Set shpCanvasShapes = shpCanvas.CanvasItems 
 
 'Adds shapes to the CanvasShapes collection 
 With shpCanvasShapes 
 .AddShape Type:=msoShapeRectangle, _ 
 Left:=0, Top:=0, Width:=50, Height:=50 
 .AddShape Type:=msoShapeOval, _ 
 Left:=5, Top:=5, Width:=40, Height:=40 
 .AddShape Type:=msoShapeIsoscelesTriangle, _ 
 Left:=0, Top:=25, Width:=50, Height:=50 
 End With 
End Sub
```

Use  **CanvasItems** (index), where index is the name or the index number, to return a single shape in the **CanvasShapes** collection. The following example sets the **Line** and **Fill** properties and vertically flips the third shape in a drawing canvas.




```vb
Sub CanvasShapeThree() 
 With ActiveDocument.Shapes(1).CanvasItems(3) 
 .Line.ForeColor.RGB = RGB(50, 0, 255) 
 .Fill.ForeColor.RGB = RGB(50, 0, 255) 
 .Flip msoFlipVertical 
 End With 
End Sub
```

Each shape is assigned a default name when it is created. For example, if you add three different shapes to a document, they might be named Rectangle 2, TextBox 3, and Oval 4. Use the  **Name** property to reference the default name or to assign a more meaningful name to a shape.


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


