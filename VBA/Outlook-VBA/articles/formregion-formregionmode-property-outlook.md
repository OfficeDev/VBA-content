---
title: FormRegion.FormRegionMode Property (Outlook)
keywords: vbaol11.chm2395
f1_keywords:
- vbaol11.chm2395
ms.prod: outlook
api_name:
- Outlook.FormRegion.FormRegionMode
ms.assetid: 8c6971a0-eddc-7e98-5f32-1a27b44d56ed
ms.date: 06/08/2017
---


# FormRegion.FormRegionMode Property (Outlook)

Returns an  **OlFormRegionMode** constant that indicates whether the form region is in a read page, compose page, or Reading Pane. Read-only.


## Syntax

 _expression_ . **FormRegionMode**

 _expression_ A variable that represents a **FormRegion** object.


## Remarks

If the user has a mail item in the Reading Pane, you can use the  **[MailItem.Sent](mailitem-sent-property-outlook.md)** property to further determine if the user is in the edit mode or the read mode of the Reading Pane. A mail item is displayed differently in the Reading Pane if it is in the edit mode (the mail item is in the Draft folder) than if it is in the read mode (the mail item is in the Inbox or Sent folder).


## See also


#### Concepts


[FormRegion Object](formregion-object-outlook.md)

