---
title: Store.IsOpen Property (Outlook)
keywords: vbaol11.chm808
f1_keywords:
- vbaol11.chm808
ms.prod: outlook
api_name:
- Outlook.Store.IsOpen
ms.assetid: 05e93457-2d17-39ac-404c-c78c76d2ef72
ms.date: 06/08/2017
---


# Store.IsOpen Property (Outlook)

Returns a  **Boolean** that indicates if the **[Store](store-object-outlook.md)** is open. Read-only.


## Syntax

 _expression_ . **IsOpen**

 _expression_ A variable that represents a **Store** object.


## Remarks

This property supports both Exchange and non-Exchange stores.

 **IsOpen** only indicates if the store is open. It does not indicate if the store is offline, or if it is an Exchange mailbox or an Exchange Public Folder and the store server is down.

Because opening a store can impose a performance overhead, and  **[Store.GetRootFolder](store-getrootfolder-method-outlook.md)** and **[Store.GetSearchFolders](store-getsearchfolders-method-outlook.md)** will open a store if it is not already open, you can use **IsOpen** before deciding to call **GetRootFolder** or **GetSearchFolders** to minimize performance overhead.


## See also


#### Concepts


[Store Object](store-object-outlook.md)

