---
title: Shapes Object (Word)
keywords: vbawd10.chm2463
f1_keywords:
- vbawd10.chm2463
ms.prod: word
ms.assetid: 0907eed3-886e-8e73-0e5e-71f4b37ddd5b
ms.date: 06/08/2017
---


# Shapes Object (Word)

A collection of  **Shape** objects that represent all the shapes in a document or all the shapes in all the headers and footers in a document. Each **[Shape](shape-object-word.md)** object represents an object in the drawing layer, such as an AutoShape, freeform, OLE object, or picture.


## Remarks

If you want to work with a subset of the shapes on a document — for example, to do something to only the AutoShapes on the document or to only the selected shapes — you must construct a  **[ShapeRange](shaperange-object-word.md)** collection that contains the shapes you want to work with.

Use the  **Shapes** property to return the **Shapes** collection. The following example selects all the shapes on the active document.




```
ActiveDocument.Shapes.SelectAll
```


 **Note**  If you want to do something (like delete or set a property) to all the shapes on a document at the same time, use the  **Range** method to create a **ShapeRange** object that contains all the shapes in the **Shapes** collection, and then apply the appropriate property or method to the **ShapeRange** object.

Use one of the following methods of the  **Shapes** collection: **AddCallout**, **AddCurve**, **AddLabel**, **AddLine**, **AddOleControl**, **AddOleObject**, **AddPolyline**, **AddShape**, **AddTextbox**, **AddTextEffect**, or **BuildFreeForm** to add a shape to a document return a **Shape** object that represents the newly created shape The following example adds a rectangle to the active document.




```
ActiveDocument.Shapes.AddShape msoShapeRectangle, 50, 50, 100, 200
```

Use  **Shapes** (Index), where Index is the name or the index number, to return a single **Shape** object. The following example horizontally flips shape one on the active document.




```
ActiveDocument.Shapes(1).Flip msoFlipHorizontal
```

This example horizontally flips the shape named "Rectangle 1" on the active document.




```
ActiveDocument.Shapes("Rectangle 1").Flip msoFlipHorizontal
```

Each shape is assigned a default name when it is created. For example, if you add three different shapes to a document, they might be named "Rectangle 2," "TextBox 3," and "Oval 4." To give a shape a more meaningful name, set the  **Name** property.

The  **Shapes** collection does not include **[InlineShape](inlineshape-object-word.md)** objects. **InlineShape** objects are treated like characters and are positioned as characters within a line of text. **Shape** objects are anchored to a range of text but are free-floating and can be positioned anywhere on the page. You can use the **ConvertToInlineShape** method and the **ConvertToShape** method to convert shapes from one type to the other. You can convert only pictures, OLE objects, and ActiveX controls to inline shapes.

The  **Count** property for this collection in a document returns the number of items in the main story only. To count the shapes in all the headers and footers, use the **Shapes** collection with any **HeaderFooter** object.


## Methods



|**Name**|
|:-----|
|[AddCallout](shapes-addcallout-method-word.md)|
|[AddCanvas](shapes-addcanvas-method-word.md)|
|[AddChart2](shapes-addchart2-method-word.md)|
|[AddCurve](shapes-addcurve-method-word.md)|
|[AddLabel](shapes-addlabel-method-word.md)|
|[AddLine](shapes-addline-method-word.md)|
|[AddOLEControl](shapes-addolecontrol-method-word.md)|
|[AddOLEObject](shapes-addoleobject-method-word.md)|
|[AddPicture](shapes-addpicture-method-word.md)|
|[AddPolyline](shapes-addpolyline-method-word.md)|
|[AddShape](shapes-addshape-method-word.md)|
|[AddSmartArt](shapes-addsmartart-method-word.md)|
|[AddTextbox](shapes-addtextbox-method-word.md)|
|[AddTextEffect](shapes-addtexteffect-method-word.md)|
|[AddWebVideo](shapes-addwebvideo-method-word.md)|
|[BuildFreeform](shapes-buildfreeform-method-word.md)|
|[Item](shapes-item-method-word.md)|
|[Range](shapes-range-method-word.md)|
|[SelectAll](shapes-selectall-method-word.md)|

## Properties



|**Name**|
|:-----|
|[Application](shapes-application-property-word.md)|
|[Count](shapes-count-property-word.md)|
|[Creator](shapes-creator-property-word.md)|
|[Parent](shapes-parent-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
