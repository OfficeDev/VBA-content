---
title: Document.ViewVerticalBaseLineGuides Property (Publisher)
keywords: vbapb10.chm196729
f1_keywords:
- vbapb10.chm196729
ms.prod: publisher
api_name:
- Publisher.Document.ViewVerticalBaseLineGuides
ms.assetid: 711335ab-237b-65a2-534a-7635cfba474e
ms.date: 06/08/2017
---


# Document.ViewVerticalBaseLineGuides Property (Publisher)

Sets or returns a  **Boolean** that represents whether or not the vertical baseline guides are visible in the specified **Document** object. **True** if they are visible. **False** if they are not visible. Read/write.


## Syntax

 _expression_. **ViewVerticalBaseLineGuides**

 _expression_A variable that represents a  **Document** object.


### Return Value

Boolean


## Remarks

The default setting for this property is  **False**.


## Example

The following example makes the vertical baseline guides visible in the active document.


```vb
Dim objDocument As Document 
Set objDocument = ActiveDocument 
objDocument.ViewVerticalBaseLineGuides = True 

```


