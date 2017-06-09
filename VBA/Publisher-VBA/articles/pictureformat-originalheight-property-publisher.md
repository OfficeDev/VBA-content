---
title: PictureFormat.OriginalHeight Property (Publisher)
keywords: vbapb10.chm3604774
f1_keywords:
- vbapb10.chm3604774
ms.prod: publisher
api_name:
- Publisher.PictureFormat.OriginalHeight
ms.assetid: 0bf97bb1-d333-a7ed-686c-da2f3cce97c5
ms.date: 06/08/2017
---


# PictureFormat.OriginalHeight Property (Publisher)

Returns a  **Variant** representing the height, in points, of the specified linked picture or OLE object. Read-only.


## Syntax

 _expression_. **OriginalHeight**

 _expression_A variable that represents an  **PictureFormat** object.


### Return Value

Variant


## Remarks

This property applies only to linked pictures or OLE objects. Returns "Permission Denied" for shapes representing embedded or pasted pictures.

To determine whether a shape represents a linked picture, use either the  **[Type](shape-type-property-publisher.md)** property of the **[Shape](shape-object-publisher.md)** object, or the **[IsLinked](pictureformat-islinked-property-publisher.md)** property of the **[PictureFormat](pictureformat-object-publisher.md)** object.


## Example

The following example tests each picture in the active publication, and returns selected image properties for pictures that are linked.


```vb
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 
 Debug.Print "File Name: " &; .Filename 
 Debug.Print "Horizontal Scaling: " &; .HorizontalScale &; "%" 
 Debug.Print "Original Image Height: " &; .OriginalHeight &; " points" 
 Debug.Print "Height in publication: " &; .Height &; " points" 
 
 End With 
 End If 
 Next shpLoop 
Next pgLoop 

```


