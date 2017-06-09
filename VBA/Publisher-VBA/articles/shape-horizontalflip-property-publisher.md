---
title: Shape.HorizontalFlip Property (Publisher)
keywords: vbapb10.chm2228288
f1_keywords:
- vbapb10.chm2228288
ms.prod: publisher
api_name:
- Publisher.Shape.HorizontalFlip
ms.assetid: 5a940631-c63a-efdf-6cfb-dc6b82594028
ms.date: 06/08/2017
---


# Shape.HorizontalFlip Property (Publisher)

Indicates whether the specified shape has been flipped around its horizontal axis. Read-only.


## Syntax

 _expression_. **HorizontalFlip**

 _expression_A variable that represents a  **Shape** object.


## Remarks

The  **HorizontalFlip** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**|The shape has not been flipped around its horizontal axis.|
| **msoTriStateMixed**|Indicates a combination of  **msoTrue** and **msoFalse** for the specified shape.|
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


