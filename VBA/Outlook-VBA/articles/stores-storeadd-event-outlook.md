---
title: Stores.StoreAdd Event (Outlook)
keywords: vbaol11.chm2755
f1_keywords:
- vbaol11.chm2755
ms.prod: outlook
api_name:
- Outlook.Stores.StoreAdd
ms.assetid: 26e7eddc-9c5a-ffff-d574-afa48e5953d8
ms.date: 06/08/2017
---


# Stores.StoreAdd Event (Outlook)

Occurs when a  **[Store](store-object-outlook.md)** has been added to the current session either programmatically or through user action.


## Syntax

 _expression_ . **StoreAdd**( **_Store_** )

 _expression_ A variable that represents a **Stores** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Store_|Required| **Store**|The  **Store** to be added to the current session.|

## Remarks

Outlook must be running in order for this event to fire. This event will fire when any of the following occur:


- A store is added through the  **Open Outlook Data File** dialog box, by selecting **Open** and then **Outlook Data File** on the **File** menu.
    
- A store is added through the  **Data Files** tab of the **Account Manager** dialog box.
    
- A store is added successfully by calling the  **[Namespace.AddStore](namespace-addstore-method-outlook.md)** method.
    


This event will not fire when any of the following occurs:


- When Outlook starts and opens a primary or delegate store. 
    
- If a store is added through the  **Mail** applet in the Microsoft Windows Control Panel and Outlook is not running.
    
- A delegate store is added through the  **Advanced** tab of the **Microsoft Exchange Server** dialog box.
    


You can use this event to determine whether a store has been added and take appropriate actions on items in that store. Otherwise, you would have to resort to polling the  **[Stores](stores-object-outlook.md)** collection.


## See also


#### Concepts


[Stores Object](stores-object-outlook.md)

