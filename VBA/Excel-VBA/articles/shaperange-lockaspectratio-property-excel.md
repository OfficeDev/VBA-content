---
title: ShapeRange.LockAspectRatio Property (Excel)
keywords: vbaxl10.chm640109
f1_keywords:
- vbaxl10.chm640109
ms.prod: excel
api_name:
- Excel.ShapeRange.LockAspectRatio
ms.assetid: 58b33bc9-de5c-1fb2-7369-7f4f8dedde58
ms.date: 06/08/2017
---


# ShapeRange.LockAspectRatio Property (Excel)

 **True** if the specified shape retains its original proportions when you resize it. **False** if you can change the height and width of the shape independently of one another when you resize it. Read/write **[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** .


## Syntax

 _expression_ . **LockAspectRatio**

 _expression_ A variable that represents a **ShapeRange** object.


## Remarks





| **MsoTriState** can be one of these **MsoTriState** constants.|
| **msoCTrue**|
| **msoFalse** . You can change the height and width of the shape independently of one another when you resize it.|
| **msoTriStateMixed**|
| **msoTriStateToggle**|
| **msoTrue** . The specified shape retains its original proportions when you resize it.|

## Example

This example adds a cube to  `myDocument`. The cube can be moved and resized, but not reproportioned.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes.AddShape(msoShapeCube, _ 
    50, 50, 100, 200).LockAspectRatio = msoTrue
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-excel.md)

