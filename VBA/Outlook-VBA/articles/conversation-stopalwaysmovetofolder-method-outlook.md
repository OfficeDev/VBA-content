---
title: Conversation.StopAlwaysMoveToFolder Method (Outlook)
keywords: vbaol11.chm3433
f1_keywords:
- vbaol11.chm3433
ms.prod: outlook
api_name:
- Outlook.Conversation.StopAlwaysMoveToFolder
ms.assetid: 3be830e9-ceea-369c-1f7b-966c68cfb8fd
ms.date: 06/08/2017
---


# Conversation.StopAlwaysMoveToFolder Method (Outlook)

Stops the action of always moving conversation items in the specified store to a specific folder.


## Syntax

 _expression_ . **StopAlwaysMoveToFolder**( **_Store_** )

 _expression_ A variable that represents a **[Conversation](conversation-object-outlook.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Store_|Required| **[Store](store-object-outlook.md)**|The store where the conversation items to be cleaned up reside.|

## Remarks

If the always-move action has not been turned on,  **StopAlwaysMoveToFolder** does not carry out any action.

If the  _Store_ parameter represents a non-delivery store such as an archive .pst store, the stop-always-move action will apply to conversation items in the default delivery store.

After you call the  **StopAlwaysMoveToFolder** method, calling the **[GetAlwaysMoveToFolder](conversation-getalwaysmovetofolder-method-outlook.md)** method returns **Null** ( **Nothing** in Visual Basic).


## See also


#### Concepts


[Conversation Object](conversation-object-outlook.md)

