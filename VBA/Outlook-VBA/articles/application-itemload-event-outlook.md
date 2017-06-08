---
title: Application.ItemLoad Event (Outlook)
keywords: vbaol11.chm446
f1_keywords:
- vbaol11.chm446
ms.prod: outlook
api_name:
- Outlook.Application.ItemLoad
ms.assetid: aed0656d-4e5a-550a-1116-76773215a897
ms.date: 06/08/2017
---


# Application.ItemLoad Event (Outlook)

Occurs when an Outlook item is loaded into memory.


## Syntax

 _expression_ . **ItemLoad**( **_Item_** , )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|A weak object reference for the loaded Outlook item.|

## Remarks

This event occurs when the Outlook item begins to load into memory. Data for the item is not yet available, other than the values for the  **Class** and **MessageClass** properties of the Outlook item, so an error occurs when calling any property other than **Class** or **MessageClass** for the Outlook item returned in _Item_. Similarly, an error occurs if you attempt to call any method from the Outlook item, or if you call the  **[GetObjectReference](application-getobjectreference-method-outlook.md)** method of the **[Application](application-object-outlook.md)** object on the Outlook item returned in _Item_.

The  **ItemLoad** event should typically be implemented as a means to hook up item-level event handlers such as **BeforeRead**,  **Open**,  **Send**, and  **Write**.

This event is not raised when the following conditions occur:


- An Outlook item is synchronized with a folder.
    
- A server-side rule is triggered for an Outlook item.
    
- A reminder is triggered for an Outlook item.
    
- A Desktop Alert is displayed for an Outlook item.
    

## See also


#### Concepts


[Application Object](application-object-outlook.md)

