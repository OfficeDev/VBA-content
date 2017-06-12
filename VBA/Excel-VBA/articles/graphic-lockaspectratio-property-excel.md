---
title: Graphic.LockAspectRatio Property (Excel)
keywords: vbaxl10.chm694082
f1_keywords:
- vbaxl10.chm694082
ms.prod: excel
api_name:
- Excel.Graphic.LockAspectRatio
ms.assetid: d38851e4-7ca6-bb1f-4b16-03fe78fae726
ms.date: 06/08/2017
---


# Graphic.LockAspectRatio Property (Excel)

 **True** if the specified shape retains its original proportions when you resize it. **False** if you can change the height and width of the shape independently of one another when you resize it. Read/write **[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** .


## Syntax

 _expression_ . **LockAspectRatio**

 _expression_ A variable that represents a **Graphic** object.


## Remarks





| **MsoTriState** can be one of these **MsoTriState** constants.|
| **msoCTrue**|
| **msoFalse** . You can change the height and width of the shape independently of one another when you resize it.|
| **msoTriStateMixed**|
| **msoTriStateToggle**|
| **msoTrue** . The specified shape retains its original proportions when you resize it.|

## See also


#### Concepts


[Graphic Object](graphic-object-excel.md)

