---
title: PictureFormat.IsTrueColor Property (Publisher)
keywords: vbapb10.chm3604770
f1_keywords:
- vbapb10.chm3604770
ms.prod: publisher
api_name:
- Publisher.PictureFormat.IsTrueColor
ms.assetid: 63708d40-996a-67ca-b4eb-dd53c83d1764
ms.date: 06/08/2017
---


# PictureFormat.IsTrueColor Property (Publisher)

Returns an  **MsoTriState** constant indicating whether the specified picture or OLE object contains color data of 24 bits per channel or greater. Read-only.


## Syntax

 _expression_. **IsTrueColor**

 _expression_A variable that represents an  **PictureFormat** object.


### Return Value

MsoTriState


## Remarks

For pictures that are not TrueColor, use the  **[ColorsInPalette](pictureformat-colorsinpalette-property-publisher.md)** property of the **[PictureFormat](pictureformat-object-publisher.md)** object to determine the number of colors in the picture's palette.

The  **IsTrueColor** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|The specified picture does not contain color data of 24 bits per channel or greater.|
| **msoTriStateMixed**|Return value indicating a combination of  **msoTrue** and **msoFalse** for the specified shape range.|
| **msoTrue**| The specified picture contains color data of 24 bits per channel or greater.|

## Example

The following example tests each picture in the active document and prints whether the picture is TrueColor. If it is not TrueColor, the example prints how many colors are in the picture's palette.


```vb
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbLinkedPicture Or shpLoop.Type = pbPicture Then 
 
 With shpLoop.PictureFormat 
 If .IsEmpty = msoFalse Then 
 Debug.Print .Filename 
 If .IsTrueColor = msoTrue Then 
 Debug.Print "This picture is TrueColor" 
 Else 
 Debug.Print "This picture contains " &; .ColorsInPalette &; " colors." 
 End If 
 End If 
 End With 
 
 End If 
 Next shpLoop 
Next pgLoop 

```


