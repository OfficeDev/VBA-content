---
title: Shape.LayoutInCell Property (Word)
keywords: vbawd10.chm161480849
f1_keywords:
- vbawd10.chm161480849
ms.prod: word
api_name:
- Word.Shape.LayoutInCell
ms.assetid: 6a80b806-2a7b-aced-4601-774d8937ee2f
ms.date: 06/08/2017
---


# Shape.LayoutInCell Property (Word)

Returns a  **Long** that represents whether a shape in a table is displayed inside or outside the table.


## Syntax

 _expression_ . **LayoutInCell**

 _expression_ Required. A variable that represents a **[Shape](shape-object-word.md)** object.


## Remarks

 **True** indicates that a specified picture is displayed inside the table. **False** indicates that a specified picture is displayed outside the table.

The  **LayoutInCell** property corresponds to the **Layout in table cell** option in the **Advanced Layout** dialog box for picture formatting.


 **Note**  Setting the  **LayoutInCell** property will take effect only if the **Type** property of the **WrapFormat** object is set to something other than **wdWrapTypeInline** or **wdWrapTypeNone** .


## Example

The following example disables the  **Layout in table cell** option for the first shape in the active document. This example assumes that the specified shape is within a table and is not an inline shape.


```vb
ActiveDocument.Shapes(1).LayoutInCell = False
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

