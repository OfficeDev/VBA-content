---
title: PictureFormat.OriginalHasAlphaChannel Property (Publisher)
keywords: vbapb10.chm3604773
f1_keywords:
- vbapb10.chm3604773
ms.prod: publisher
api_name:
- Publisher.PictureFormat.OriginalHasAlphaChannel
ms.assetid: e58a97d2-4ced-d3cf-56b2-6a89df02bcdf
ms.date: 06/08/2017
---


# PictureFormat.OriginalHasAlphaChannel Property (Publisher)

Returns an  **MsoTriState** constant depending on whether the original, linked picture contains an alpha channel. Read-only.


## Syntax

 _expression_. **OriginalHasAlphaChannel**

 _expression_A variable that represents an  **PictureFormat** object.


### Return Value

MsoTriState


## Remarks

This property only applies to linked pictures. Returns "Permission Denied" for shapes representing embedded or pasted pictures.

Use either of the following properties to determine whether a shape represents a linked picture:


-  The **[Type](shape-type-property-publisher.md)** property of the **[Shape](shape-object-publisher.md)** object
    
- The  **[IsLinked](pictureformat-islinked-property-publisher.md)** property of the **[PictureFormat](pictureformat-object-publisher.md)** object
    


An alpha channel is a special 8-bit channel used by some image processing software to contain additional data, such as masking information or transparency information.

The  **OriginalHasAlphaChannel** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**| The original, linked picture does not contain an alpha channel.|
| **msoTriStateMixed**| Indicates a combination of **msoTrue** and **msoFalse** for the specified shape range.|
| **msoTrue**|The original, linked picture contains an alpha channel.|

## Example

The following example returns whether the first shape on the first page of the active publication contains an alpha channel. If the picture is linked, and the original picture contains an alpha channel, that is also returned. This example assumes the shape is a picture.


```vb
With ActiveDocument.Pages(1).Shapes(1).PictureFormat 
 If .HasAlphaChannel = msoTrue Then 
 Debug.Print .Filename 
 Debug.Print "This picture contains an alpha channel." 
 
 If .IsLinked = msoTrue Then 
 If .OriginalHasAlphaChannel = msoTrue Then 
 Debug.Print "The linked picture " &; _ 
 "also contains an alpha channel." 
 End If 
 End If 
 End If 
End With 

```


