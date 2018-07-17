---
title: Story.HasTextFrame Property (Publisher)
keywords: vbapb10.chm5832708
f1_keywords:
- vbapb10.chm5832708
ms.prod: publisher
api_name:
- Publisher.Story.HasTextFrame
ms.assetid: 10c3a002-05ae-1167-784c-d62066de802d
ms.date: 06/08/2017
---


# Story.HasTextFrame Property (Publisher)

Indicates whether the specified shape has a  **TextFrame** object associated with it. Read-only.


## Syntax

 _expression_. **HasTextFrame**

 _expression_A variable that represents a  **Story** object.


## Remarks

If the  **HasTextFrame** property is true, clients must check the value of the **HasText** property of the **TextFrame** object to determine if there is any text on the shape.

The  **HasTextFrame** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**| The specified shape does not have a **TextFrame** object associated with it.|
| **msoTriStateMixed**| Indicates a combination of **msoTrue** and **msoFalse** for the specified shape range.|
| **msoTrue**| The specified shape has a **TextFrame** object associated with it.|

## Example

This example tests all the shapes in the selection and if none have text frames associated with them, they are left aligned.


```vb
Sub MoveLeft() 
 
 Dim shpAll As ShapeRange 
 
 Set shpAll = Application.ActiveDocument.Selection.ShapeRange 
 If shpAll.HasTextFrame = msoFalse Then 
 shpAll.Align msoAlignLefts, msoTrue 
 End If 
 
End Sub
```


