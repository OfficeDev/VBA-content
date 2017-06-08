---
title: Document.ViewHorizontalBaseLineGuides Property (Publisher)
keywords: vbapb10.chm196728
f1_keywords:
- vbapb10.chm196728
ms.prod: publisher
api_name:
- Publisher.Document.ViewHorizontalBaseLineGuides
ms.assetid: e5471313-38e0-9454-04af-4c85d976b312
ms.date: 06/08/2017
---


# Document.ViewHorizontalBaseLineGuides Property (Publisher)

Sets or returns a  **Boolean** that represents whether or not the horizontal baseline guides are visible in the specified **Document** object. **True** if they are visible. **False** if they are not visible. Read/write.


## Syntax

 _expression_. ****

 _expression_A variable that represents a  **Document** object.


### Return Value

Boolean


## Remarks

The default setting for this property is  **False**.


## Example

The following example makes the horizontal baseline guides visible in the active document.


```vb
Dim objDocument As Document 
Set objDocument = ActiveDocument 
objDocument.ViewHorizontalBaseLineGuides = True 

```


