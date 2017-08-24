---
title: ShapeRange.HorizontalFlip Property (Publisher)
keywords: vbapb10.chm2293824
f1_keywords:
- vbapb10.chm2293824
ms.prod: publisher
api_name:
- Publisher.ShapeRange.HorizontalFlip
ms.assetid: c0dd2f4a-0baf-3720-113a-b929193f2b1d
ms.date: 06/08/2017
---


# ShapeRange.HorizontalFlip Property (Publisher)

Indicates whether the specified shape has been flipped around its horizontal axis. Read-only.


## Syntax

 _expression_. **HorizontalFlip**

 _expression_A variable that represents a  **ShapeRange** object.


## Remarks

The  **HorizontalFlip** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|The shape has not been flipped around its horizontal axis.|
| **msoTriStateMixed**|Indicates a combination of  **msoTrue** and **msoFalse** for the specified shape range.|
| **msoTrue**|The shape has been flipped around its horizontal axis.|

## Example

This example restores each shape on the active publication to its original state if it has been flipped horizontally or vertically.


```vb
Sub Flipper() 
 
 Dim shpS As Shape 
 
 For Each shpS In ActiveDocument.MasterPages.Item(1).Shapes 
 If shpS.HorizontalFlip = msoTrue Then shpS.Flip msoFlipHorizontal 
 If shpS.VerticalFlip = msoTrue Then shpS.Flip msoFlipVertical 
 Next 
 
End Sub
```


