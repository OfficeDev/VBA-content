---
title: PictureFormat.VerticalPictureLocking Property (Publisher)
keywords: vbapb10.chm3604745
f1_keywords:
- vbapb10.chm3604745
ms.prod: publisher
api_name:
- Publisher.PictureFormat.VerticalPictureLocking
ms.assetid: 0575d733-b515-2256-7136-6ec07532ab67
ms.date: 06/08/2017
---


# PictureFormat.VerticalPictureLocking Property (Publisher)

Returns or sets a  **PbVerticalPictureLocking** constant indicating where newly inserted pictures appear in relation to the specified frame. Read/write.


## Syntax

 _expression_. **VerticalPictureLocking**

 _expression_A variable that represents a  **PictureFormat** object.


### Return Value

PbVerticalPictureLocking


## Remarks

The  **Vertical PictureLocking** property value can be one of the **PbVerticalPictureLocking** constants declared in the Microsoft Publisher type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **pbVerticalLockingBottom**|New pictures are inserted along the bottom edge of the frame.|
| **pbVerticalLockingNone**|New pictures are inserted in the center between the top and bottom edges of the frame.|
| **pbVerticalLockingStretch**|New pictures are vertically stretched to the full height of the frame.|
| **pbVerticalLockingTop**|New pictures are inserted along the top edge of the frame.|

## Example

The following example locks the specified picture to the upper-left corner of the picture frame. Shape one on page one of the active publication must be a picture frame for this example to work.


```vb
With ActiveDocument.Pages(1).Shapes(1).PictureFormat 
 .HorizontalPictureLocking = pbHorizontalLockingLeft 
 .VerticalPictureLocking = pbVerticalLockingTop 
End With
```


