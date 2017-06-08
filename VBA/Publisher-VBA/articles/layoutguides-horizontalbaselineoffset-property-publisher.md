---
title: LayoutGuides.HorizontalBaseLineOffset Property (Publisher)
keywords: vbapb10.chm1114131
f1_keywords:
- vbapb10.chm1114131
ms.prod: publisher
api_name:
- Publisher.LayoutGuides.HorizontalBaseLineOffset
ms.assetid: b80d2114-8132-db13-a50d-ce904dbe5919
ms.date: 06/08/2017
---


# LayoutGuides.HorizontalBaseLineOffset Property (Publisher)

Returns a  **Single** that represents the horizontal baseline offset of the specified **LayoutGuides** object. Read/Write.


## Syntax

 _expression_. **HorizontalBaseLineOffset**

 _expression_A variable that represents a  **LayoutGuides** object.


### Return Value

Single


## Remarks

When setting the layout guide properties of a  **Page** object it must be returned from the **MasterPages** collection.


## Example

This example sets the horizontal baseline offset of the layout guides object to 12 for the second master page in the active document.


```vb
Dim objLayout As LayoutGuides 
Set objLayout = ActiveDocument.MasterPages(2).LayoutGuides 
objLayout.HorizontalBaseLineSpacing = 12 

```


