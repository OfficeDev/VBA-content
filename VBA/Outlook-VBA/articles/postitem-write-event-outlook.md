---
title: PostItem.Write Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.PostItem.Write
ms.assetid: 27ab5442-2ce2-c40e-b95c-6e23f29e124b
ms.date: 06/08/2017
---


# PostItem.Write Event (Outlook)

Occurs when an instance of the parent object is saved, either explicitly (for example, using the  **[Save](postitem-save-method-outlook.md)** or **[SaveAs](postitem-saveas-method-outlook.md)** methods) or implicitly (for example, in response to a prompt when closing the item's inspector).


## Syntax

 _expression_ . **Write**( **_Cancel_** )

 _expression_ A variable that represents a **PostItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| (Not used in VBScript). **False** when the event occurs. If the event procedure sets this argument to **True** , the save operation is not completed.|

## Remarks

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False** , the save operation is not completed.


## See also


#### Concepts


[PostItem Object](postitem-object-outlook.md)

