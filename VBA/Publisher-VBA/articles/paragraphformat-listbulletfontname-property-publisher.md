---
title: ParagraphFormat.ListBulletFontName Property (Publisher)
keywords: vbapb10.chm5439525
f1_keywords:
- vbapb10.chm5439525
ms.prod: publisher
api_name:
- Publisher.ParagraphFormat.ListBulletFontName
ms.assetid: aa0269a1-c5a8-1705-551f-6b1b849701e9
ms.date: 06/08/2017
---


# ParagraphFormat.ListBulletFontName Property (Publisher)

Sets or retrieves a  **String** representing the list bullet font name from the specified paragraphs. Read/write.


## Syntax

 _expression_. **ListBulletFontName**

 _expression_A variable that represents a  **ParagraphFormat** object.


### Return Value

String


## Remarks

Returns an "Access Denied" message if the list is not a bulleted list.


## Example

This example tests to see if the list type is a bulleted list. If it is, the  **ListBulletFontName** is set to "Verdana" and the **ListFontSize** is set to 24.


```vb
Dim objParaForm As ParagraphFormat 
 
Set objParaForm = ActiveDocument.Pages(1).Shapes(1) _ 
.TextFrame.TextRange.ParagraphFormat 
 
With objParaForm 
 If .ListType = pbListTypeBullet Then 
 .ListBulletFontName = "Verdana" 
 .ListBulletFontSize = 24 
 End If 
End With 

```


