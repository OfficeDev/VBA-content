---
title: FormRegion.Select Method (Outlook)
keywords: vbaol11.chm3216
f1_keywords:
- vbaol11.chm3216
ms.prod: outlook
api_name:
- Outlook.FormRegion.Select
ms.assetid: b0a16d61-6c6f-7eb5-d9e2-7f095fba11cf
ms.date: 06/08/2017
---


# FormRegion.Select Method (Outlook)

Makes the form region the active form region such that it becomes visible.


## Syntax

 _expression_ . **Select**

 _expression_ A variable that represents a **FormRegion** object.


## Remarks

If the form region is an adjoining form region, then  **Select** will expand the form region (if it is not already expanded) and set focus on the first control on that page. If the form region is a separate form region and is not already the active page, then **Select** will swtich to the form region page and set focus on the first control on that page. If the form region is a separate form region and is already the active page, then nothing happens.


## See also


#### Concepts


[FormRegion Object](formregion-object-outlook.md)

