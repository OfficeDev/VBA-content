---
title: DocumentItem.Open Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.DocumentItem.Open
ms.assetid: e7d95148-9fa2-3f0f-cbfc-f835c9017c3b
ms.date: 06/08/2017
---


# DocumentItem.Open Event (Outlook)

Occurs when an instance of the parent object is being opened in an  **[Inspector](inspector-object-outlook.md)** .


## Syntax

 _expression_ . **Open**( **_Cancel_** )

 _expression_ A variable that represents a **DocumentItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True** , the open operation is not completed and the inspector is not displayed.|

## Remarks

When this event occurs, the  **Inspector** object is initialized but not yet displayed. The **Open** event differs from the **[Read](documentitem-read-event-outlook.md)** event in that **Read** occurs whenever the user selects the item in a view that supports in-cell editing as well as when the item is being opened in an inspector.

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False** , the open operation is not completed and the inspector is not displayed.


## See also


#### Concepts


[DocumentItem Object](documentitem-object-outlook.md)

