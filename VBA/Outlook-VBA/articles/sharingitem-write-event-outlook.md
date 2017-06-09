---
title: SharingItem.Write Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.SharingItem.Write
ms.assetid: 22cfb332-d9e9-005a-fb6c-e77ff098a444
ms.date: 06/08/2017
---


# SharingItem.Write Event (Outlook)

Occurs when an instance of the parent object is saved, either explicitly (for example, using the  **[Save](sharingitem-save-method-outlook.md)** or **[SaveAs](sharingitem-saveas-method-outlook.md)** methods) or implicitly (for example, in response to a prompt when closing the item's inspector).


## Syntax

 _expression_ . **Write**( **_Cancel_** )

 _expression_ An expression that returns a **SharingItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| (Not used in VBScript). **False** when the event occurs. If the event procedure sets this argument to **True** , the save operation is not completed.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False** , the save operation is not completed.


## See also


#### Concepts


[SharingItem Object](sharingitem-object-outlook.md)

