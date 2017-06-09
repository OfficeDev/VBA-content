---
title: Shape.HasTable Property (PowerPoint)
keywords: vbapp10.chm547059
f1_keywords:
- vbapp10.chm547059
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.HasTable
ms.assetid: fa38891a-e915-3a5c-4169-3c14e5e7136e
ms.date: 06/08/2017
---


# Shape.HasTable Property (PowerPoint)

Returns whether the specified shape is a table. Read-only.


## Syntax

 _expression_. **HasTable**

 _expression_ A variable that represents a **Shape** object.


### Return Value

MsoTriState


## Remarks

The value of the  **HasTable** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified shape is not a table.|
|**msoTrue**| The specified shape is a table.|

## Example

This example checks the currently selected shape to see if it is a table. If it is, the code sets the width of column one to one inch (72 points).


```vb
With ActiveWindow.Selection.ShapeRange

    If .HasTable = msoTrue Then

       .Table.Columns(1).Width = 72

    End If

End With
```


## See also


#### Concepts


[Shape Object](shape-object-powerpoint.md)

