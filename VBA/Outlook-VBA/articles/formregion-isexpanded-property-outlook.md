---
title: FormRegion.IsExpanded Property (Outlook)
keywords: vbaol11.chm2389
f1_keywords:
- vbaol11.chm2389
ms.prod: outlook
api_name:
- Outlook.FormRegion.IsExpanded
ms.assetid: 6b2a033c-c852-d669-d641-098f9b6c8e35
ms.date: 06/08/2017
---


# FormRegion.IsExpanded Property (Outlook)

Returns a  **Boolean** that indicates if the form region is expanded. Read-only.


## Syntax

 _expression_ . **IsExpanded**

 _expression_ A variable that represents a **FormRegion** object.


## Remarks

This property applies to adjoining form regions only and is ignored for separate form regions.

Outlook always first loads a form region in an expanded state and sets  **IsExpanded** to **True** . If the initial state of the form region is to be collapsed, then Outlook immediately closes the form region, fires the **[Expanded](formregion-expanded-event-outlook.md)** event with the _Expand_ parameter being **False** , and sets **IsExpanded** to **False** .


## See also


#### Concepts


[FormRegion Object](formregion-object-outlook.md)

