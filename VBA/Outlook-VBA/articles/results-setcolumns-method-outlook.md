---
title: Results.SetColumns Method (Outlook)
keywords: vbaol11.chm510
f1_keywords:
- vbaol11.chm510
ms.prod: outlook
api_name:
- Outlook.Results.SetColumns
ms.assetid: 119ea78f-f61e-a95e-e9df-440499af962a
ms.date: 06/08/2017
---


# Results.SetColumns Method (Outlook)

Caches certain properties for extremely fast access to those particular properties of an item within the collection. 


## Syntax

 _expression_ . **SetColumns**( **_Columns_** )

 _expression_ A variable that represents a **Results** object.


## Remarks

The  **SetColumns** method is useful for iterating through the **[Results](results-object-outlook.md)** object. If you don't use this method, Microsoft Outlook must open each item to access the property. With the **SetColumns** method, Outlook only checks the properties that you have cached. Properties which are not cached are returned empty.


## See also


#### Concepts


[Results Object](results-object-outlook.md)

