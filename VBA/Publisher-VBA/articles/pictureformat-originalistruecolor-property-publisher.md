---
title: PictureFormat.OriginalIsTrueColor Property (Publisher)
keywords: vbapb10.chm3604775
f1_keywords:
- vbapb10.chm3604775
ms.prod: publisher
api_name:
- Publisher.PictureFormat.OriginalIsTrueColor
ms.assetid: 837109d4-3479-2500-a1fa-b4c00e0f8672
ms.date: 06/08/2017
---


# PictureFormat.OriginalIsTrueColor Property (Publisher)

Returns an  **MsoTriState** constant indicating whether the specified linked picture or OLE object contains color data of 24 bits per channel or greater. Read-only.


## Syntax

 _expression_. **OriginalIsTrueColor**

 _expression_A variable that represents an  **PictureFormat** object.


### Return Value

MsoTriState


## Remarks

This property only applies to linked pictures or OLE objects. It returns "Permission Denied" for shapes representing embedded or pasted pictures and OLE objects.

To determine whether a shape represents a linked picture, use either the  **[Type](shape-type-property-publisher.md)** property of the **[Shape](shape-object-publisher.md)** object, or the **[IsLinked](pictureformat-islinked-property-publisher.md)** property of the **[PictureFormat](pictureformat-object-publisher.md)** object.

The  **OriginalIsTrueColor** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|The specified linked picture does not contain color data of 24 bits per channel or greater.|
| **msoTriStateMixed**|Indicates a combination of  **msoTrue** and **msoFalse** for the specified shape range.|
| **msoTrue**|The specified linked picture contains color data of 24 bits per channel or greater.|

## Example

The following example returns a list of pictures in the active document that are TrueColor. If a picture is linked, and the linked picture is also TrueColor, that information is also returned.


```vb
Sub PictureColorInformation() 
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbLinkedPicture Or shpLoop.Type = pbPicture Then 
 
 With shpLoop.PictureFormat 
 If .IsEmpty = msoFalse Then 
 
 If .IsTrueColor = msoTrue Then 
 Debug.Print .Filename 
 Debug.Print "This picture is TrueColor" 
 If .IsLinked = msoTrue Then 
 If .OriginalIsTrueColor = msoTrue Then 
 Debug.Print "The linked picture is also TrueColor." 
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


