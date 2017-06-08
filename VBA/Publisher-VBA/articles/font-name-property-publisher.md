---
title: Font.Name Property (Publisher)
keywords: vbapb10.chm5373952
f1_keywords:
- vbapb10.chm5373952
ms.prod: publisher
api_name:
- Publisher.Font.Name
ms.assetid: 03561991-5456-aee3-4c04-56a2520a4d6e
ms.date: 06/08/2017
---


# Font.Name Property (Publisher)

Indicates the name of the specified font. Read/write.


## Syntax

 _expression_. **Name**

 _expression_An expression that returns a  **Font** object.


### Return Value

String


## Example

This example formats a text frame on page one as Arial bold.


```vb
With ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Font 
 .Name = "Arial" 
 .Bold = True 
End With 

```


