---
title: TaskRequestItem.BeforeCheckNames Event (Outlook)
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem.BeforeCheckNames
ms.assetid: 6093de6d-a706-3798-8f7e-a9395ef9f776
ms.date: 06/08/2017
---


# TaskRequestItem.BeforeCheckNames Event (Outlook)

Occurs just before Microsoft Outlook starts resolving names in the recipient collection for an item (which is an instance of the parent object).


## Syntax

 _expression_ . **BeforeCheckNames**( **_Cancel_** )

 _expression_ A variable that represents a **TaskRequestItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the name resolution process is not completed.|

## Remarks

You use the  **BeforeCheckNames** event in VBScript, but the event does not fire when an e-mail name is resolved on the form.

The event does not fire under the following circumstances:


- You customized a Journal Entry form and then resolved a contact in the  **Contacts** field.
    
- You customized a Contact form and then resolved a contact in the  **Contacts** field.
    
- You customized any type of form and Outlook automatically resolved the name in the background.
    
- You programmatically created and resolved a recipient.
    



## See also


#### Concepts


[TaskRequestItem Object](taskrequestitem-object-outlook.md)

