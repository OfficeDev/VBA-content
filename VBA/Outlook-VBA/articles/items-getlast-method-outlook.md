---
title: Items.GetLast Method (Outlook)
keywords: vbaol11.chm65
f1_keywords:
- vbaol11.chm65
ms.prod: outlook
api_name:
- Outlook.Items.GetLast
ms.assetid: d02a20be-19fc-fb6e-feff-b66ca0273beb
ms.date: 06/08/2017
---


# Items.GetLast Method (Outlook)

Returns the last object in the collection. 


## Syntax

 _expression_ . **GetLast**

 _expression_ A variable that represents an **Items** object.


### Return Value

An  **Object** value that represents the last object contained by the collection.


## Remarks

It returns  **Nothing** if no last object exists, for example, if the collection is empty. To ensure correct operation of the **[GetFirst](items-getfirst-method-outlook.md)** , **GetLast** , **[GetNext](items-getnext-method-outlook.md)** , and **[GetPrevious](items-getprevious-method-outlook.md)** methods in a large collection, call **GetFirst** before calling **GetNext** on that collection, and call **GetLast** before calling **GetPrevious** . To ensure that you are always making the calls on the same collection, create an explicit variable that refers to that collection before entering the loop.


## See also


#### Concepts


[Items Object](items-object-outlook.md)

