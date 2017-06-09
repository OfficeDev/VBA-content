---
title: TextFrame2.HasText Property (Excel)
ms.prod: excel
api_name:
- Excel.TextFrame2.HasText
ms.assetid: b9c7d9f4-22d3-5a45-e03b-8e06e87a2af9
ms.date: 06/08/2017
---


# TextFrame2.HasText Property (Excel)

Returns whether the specified text frame has text. Read-only  **[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** .


## Syntax

 _expression_ . **HasText**

 _expression_ A variable that represents a **TextFrame2** object.


## Example

This example formats the text in the text frame, if the text frame contains text.


```vb
With ActiveSheet.Shapes(1).TextFrame2 
If .HasText Then 
.TextRange2.Font.Name = "Arial" 

```


## See also


#### Concepts


[TextFrame2 Object](textframe2-object-excel.md)

