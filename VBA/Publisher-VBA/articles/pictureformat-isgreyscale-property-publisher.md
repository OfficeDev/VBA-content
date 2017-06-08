---
title: PictureFormat.IsGreyScale Property (Publisher)
keywords: vbapb10.chm3604768
f1_keywords:
- vbapb10.chm3604768
ms.prod: publisher
api_name:
- Publisher.PictureFormat.IsGreyScale
ms.assetid: 1f8308c1-353e-2aac-9b4b-fad300a89b97
ms.date: 06/08/2017
---


# PictureFormat.IsGreyScale Property (Publisher)

Returns a  **MsoTriState** constant that indicates whether the picture is a greyscale image. Read-only.


## Syntax

 _expression_. **IsGreyScale**

 _expression_A variable that represents an  **PictureFormat** object.


### Return Value

MsoTriState


## Remarks

The  **IsGreyScale** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|The picture is not a greyscale image.|
| **msoTriStateMixed**|Indicates a combination of  **msoTrue** and **msoFalse** for the specified shape range.|
| **msoTrue**|The specified picture is a greyscale image.|

## Example

The following example returns a list of the greyscale pictures contained in the active publication.


```vb
Sub ListGreyScalePictures() 
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
 For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 
 If shpLoop.Type = pbPicture Or shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 If .IsEmpty = msoFalse And .IsGreyScale = msoCTrue Then 
 
 Debug.Print .Filename 
 Debug.Print "Page " &; pgLoop.PageNumber 
 
 End If 
 End With 
 
 End If 
 
 Next shpLoop 
 Next pgLoop 
 
End Sub
```


