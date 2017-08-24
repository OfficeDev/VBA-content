---
title: PictureFormat.OriginalColorsInPalette Property (Publisher)
keywords: vbapb10.chm3604771
f1_keywords:
- vbapb10.chm3604771
ms.prod: publisher
api_name:
- Publisher.PictureFormat.OriginalColorsInPalette
ms.assetid: 87c67430-1a5a-47f7-822f-6af8783f73b3
ms.date: 06/08/2017
---


# PictureFormat.OriginalColorsInPalette Property (Publisher)

Returns a  **Long** that represents the number of colors in the specified linked picture's palette. Read-only.


## Syntax

 _expression_. **OriginalColorsInPalette**

 _expression_A variable that represents an  **PictureFormat** object.


### Return Value

Long


## Remarks

This property only applies to linked pictures or OLE objects that are not TrueColor (that is, they contain color data of less than 24 bits per channel.) Returns "Permission Denied" for shapes representing embedded or pasted pictures and OLE objects, or linked pictures that are TrueColor.

Use either of the following properties to determine whether a shape represents a linked picture:


-  The **[Type](shape-type-property-publisher.md)** property of the **[Shape](shape-object-publisher.md)** object
    
- The  **[IsLinked](pictureformat-islinked-property-publisher.md)** property of the **[PictureFormat](pictureformat-object-publisher.md)** object
    


Use the  **[OriginalIsTrueColor](pictureformat-originalistruecolor-property-publisher.md)** property to determine whether a linked picture contains color data of 24 bits per channel or greater.


## Example

The following example returns a list of all pictures in the active publication that are not TrueColor. The number of colors in each picture's palette is returned, and if the picture is linked and the linked picture is not TrueColor, the number of colors in its palette is also returned.


```vb
Sub PictureColorInformation() 
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbLinkedPicture Or shpLoop.Type = pbPicture Then 
 
 With shpLoop.PictureFormat 
 If .IsEmpty = msoFalse Then 
 
 If .IsTrueColor = msoFalse Then 
 Debug.Print .Filename 
 Debug.Print "This picture has " &; .ColorsInPalette &; " colors." 
 If .IsLinked = msoTrue Then 
 If .OriginalIsTrueColor = msoFalse Then 
 Debug.Print "The linked picture has " &; _ 
 .OriginalColorsInPalette &; " colors." 
 End If 
 End If 
 End If 
 
 End If 
 End With 
 
 End If 
 Next shpLoop 
Next pgLoop 
 
End Sub
```


