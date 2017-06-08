---
title: AddressEntries.GetLast Method (Outlook)
keywords: vbaol11.chm34
f1_keywords:
- vbaol11.chm34
ms.prod: outlook
api_name:
- Outlook.AddressEntries.GetLast
ms.assetid: 22b54c0f-5167-ac76-0cff-7ee4a142e1b3
ms.date: 06/08/2017
---


# AddressEntries.GetLast Method (Outlook)

Returns the last object in the  **[AddressEntries](addressentries-object-outlook.md)** collection.


## Syntax

 _expression_ . **GetLast**

 _expression_ A variable that represents an **AddressEntries** object.


### Return Value

An  **[AddressEntry](addressentry-object-outlook.md)** object that represents the last object contained by the collection.


## Remarks

It returns  **Nothing** if no last object exists, for example, if the collection is empty. To ensure correct operation of the **[GetFirst](addressentries-getfirst-method-outlook.md)** , **GetLast** , **[GetNext](addressentries-getnext-method-outlook.md)** , and **[GetPrevious](addressentries-getprevious-method-outlook.md)** methods in a large collection, call **GetFirst** before calling **GetNext** on that collection, and call **GetLast** before calling **GetPrevious** . To ensure that you are always making the calls on the same collection, create an explicit variable that refers to that collection before entering the loop.


## See also


#### Concepts


[AddressEntries Object](addressentries-object-outlook.md)

