---
title: OlActionShowOn Enumeration (Outlook)
keywords: vbaol11.chm3051
f1_keywords:
- vbaol11.chm3051
ms.prod: outlook
api_name:
- Outlook.OlActionShowOn
ms.assetid: 6a6e4156-d593-b5c7-8ed1-e133d61332df
ms.date: 06/08/2017
---


# OlActionShowOn Enumeration (Outlook)

Identifies where an  **[Action](action-object-outlook.md)** is displayed as an available action.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **olDontShow**|0|Indicates that the action will not be displayed on the menu or toolbar.|
| **olMenu**|1|Indicates that the action will be displayed as an available action on the menu.|
| **olMenuAndToolbar**|2|Indicates that the action will be displayed as an available action on the menu and the toolbar.|

## Remarks

Displaying an action on a toolbar is only supported in versions of Outlook without the Office Fluent ribbon, before Microsoft Office Outlook 2007. In versions of Outlook that contain the Ribbon, custom actions are displayed only on the  **Custom Actions** menu on the ribbon for an inspector, and on the context menu of an item.


