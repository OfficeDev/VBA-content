---
title: LayoutGuides.VerticalBaseLineSpacing Property (Publisher)
keywords: vbapb10.chm1114134
f1_keywords:
- vbapb10.chm1114134
ms.prod: publisher
api_name:
- Publisher.LayoutGuides.VerticalBaseLineSpacing
ms.assetid: 49391fbd-86c0-b53f-ff57-009af9341e74
ms.date: 06/08/2017
---


# LayoutGuides.VerticalBaseLineSpacing Property (Publisher)

Returns a  **Single** that represents the vertical baseline spacing of the specified **LayoutGuides** object. Read/write.


## Syntax

 _expression_. **VerticalBaseLineSpacing**

 _expression_A variable that represents a  **LayoutGuides** object.


### Return Value

Single


## Remarks

When setting the layout guide properties of a  **Page** object it must be returned from the **MasterPages** collection.


## Example

This example sets the vertical baseline spacing of the  **LayoutGuides** object to 12 for the second master page in the active document.


```vb
Dim objLayout As LayoutGuides 
Set objLayout = ActiveDocument.MasterPages(2).LayoutGuides 
objLayout.VerticalBaseLineSpacing = 12 

```


