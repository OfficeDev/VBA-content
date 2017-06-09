---
title: Conversation.GetChildren Method (Outlook)
keywords: vbaol11.chm3391
f1_keywords:
- vbaol11.chm3391
ms.prod: outlook
api_name:
- Outlook.Conversation.GetChildren
ms.assetid: bc68ccd6-9d3c-c404-72b0-a21dbc99ed63
ms.date: 06/08/2017
---


# Conversation.GetChildren Method (Outlook)

Returns a  **[SimpleItems](simpleitems-object-outlook.md)** collection that contains all items under the specified conversation node.


## Syntax

 _expression_ . **GetChildren**( **_Item_** )

 _expression_ A variable that represents a **[Conversation](conversation-object-outlook.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|A conversation node that is part of a conversation.|

### Return Value

A  **SimpleItems** collection that represents the set of child items in the conversation under the node specified by the _Item_ parameter.


## Remarks

The returned  **SimpleItems** collection contains immediate child items of the conversation node specified by the _Item_ parameter. If the specified node does not exist in the conversation, the **GetChildren** method returns an error.

If no child items exist under that node, the  **GetChildren** method returns a **SimpleItems** collection with zero objects, in which case the **[Count](simpleitems-count-property-outlook.md)** property of the **SimpleItems** collection returns 0.


## See also


#### Concepts


[Conversation Object](conversation-object-outlook.md)

