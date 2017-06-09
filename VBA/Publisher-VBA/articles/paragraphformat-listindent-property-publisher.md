---
title: ParagraphFormat.ListIndent Property (Publisher)
keywords: vbapb10.chm5439522
f1_keywords:
- vbapb10.chm5439522
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.ListIndent
ms.assetid: b42000ea-0636-88cf-b7ed-c71384a2b0d5
ms.date: 06/08/2017
---


# ParagraphFormat.ListIndent Property (Publisher)

Returns or sets a  **Single** that represents the list indent value (in points) for the specified **ParagraphFormat** object. Read/write.


## Syntax

 _expression_. **ListIndent**

 _expression_A variable that represents a  **ParagraphFormat** object.


### Return Value

Single


## Example

This example sets the  **ListIndent** property of a **ParagraphFormat** object to 0.25 inches. The **InchesToPoints** method is used to convert inches to points.


```vb
Dim objParaForm As ParagraphFormat 
 
Set objParaForm = ActiveDocument.Pages(1).Shapes(1) _ 
.TextFrame.TextRange.ParagraphFormat 
 
With objParaForm 
 .ListIndent = InchesToPoints(0.25) 
End With 

```


