---
title: RemoteItem.Write Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.RemoteItem.Write
ms.assetid: a38eef6b-23da-ba10-ad94-cc63e2bf60c2
ms.date: 06/08/2017
---


# RemoteItem.Write Event (Outlook)

Occurs when an instance of the parent object is saved, either explicitly (for example, using the  **[Save](remoteitem-save-method-outlook.md)** or **[SaveAs](remoteitem-saveas-method-outlook.md)** methods) or implicitly (for example, in response to a prompt when closing the item's inspector).


## Syntax

 _expression_ . **Write**( **_Cancel_** )

 _expression_ A variable that represents a **RemoteItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| (Not used in VBScript). **False** when the event occurs. If the event procedure sets this argument to **True** , the save operation is not completed.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False** , the save operation is not completed.


## See also


#### Concepts


[RemoteItem Object](remoteitem-object-outlook.md)

