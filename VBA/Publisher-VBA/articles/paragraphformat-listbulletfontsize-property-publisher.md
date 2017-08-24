---
title: ParagraphFormat.ListBulletFontSize Property (Publisher)
keywords: vbapb10.chm5439524
f1_keywords:
- vbapb10.chm5439524
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.ListBulletFontSize
ms.assetid: 1ff1de0f-afcc-cc9c-bf45-d745695db89b
ms.date: 06/08/2017
---


# ParagraphFormat.ListBulletFontSize Property (Publisher)

Sets or retrieves a  **Single** that represents the list bullet font size from the specified paragraphs. Read/write.


## Syntax

 _expression_. **ListBulletFontSize**

 _expression_A variable that represents a  **ParagraphFormat** object.


### Return Value

Single


## Remarks

Returns an "Access Denied" message if the list is not a bulleted list.


## Example

This example tests to see if the list type is a bulleted list. If it is, the  **ListFontSize** is set to 24 and the **ListBulletFontName** is set to "Verdana".


```vb
Dim objParaForm As ParagraphFormat 
 
Set objParaForm = ActiveDocument.Pages(1).Shapes(1) _ 
.TextFrame.TextRange.ParagraphFormat 
 
With objParaForm 
 If .ListType = pbListTypeBullet Then 
 .ListBulletFontSize = 24 
 .ListBulletFontName = "Verdana" 
 End If 
End With 
 
 

```


