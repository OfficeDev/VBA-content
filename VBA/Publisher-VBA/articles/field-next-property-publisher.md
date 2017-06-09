---
title: Field.Next Property (Publisher)
keywords: vbapb10.chm6094854
f1_keywords:
- vbapb10.chm6094854
ms.prod: publisher
api_name:
- Publisher.Field.Next
ms.assetid: a8f0a246-c55e-715e-3f97-a2f08c383e87
ms.date: 06/08/2017
---


# Field.Next Property (Publisher)

Returns a  **[Field](field-object-publisher.md)** object that represents the next field in a text range.


## Syntax

 _expression_. **Next**

 _expression_A variable that represents a  **Field** object.


### Return Value

Field


## Example

This example makes the field next to the first field in the specified text range bold. This assumes that there are at least two fields in the specified text range.


```vb
Sub GoToNextField() 
 ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange _ 
 .Fields(1).Next.TextRange.Font.Bold = msoTrue 
End Sub
```


