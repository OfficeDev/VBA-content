---
title: NavigationFolders.Add Method (Outlook)
keywords: vbaol11.chm2897
f1_keywords:
- vbaol11.chm2897
ms.prod: outlook
api_name:
- Outlook.NavigationFolders.Add
ms.assetid: f88fd69a-8684-bfc4-bc20-1cff5c44974e
ms.date: 06/08/2017
---


# NavigationFolders.Add Method (Outlook)

Adds the specified  **[Folder](folder-object-outlook.md)** , as a **[NavigationFolder](navigationfolder-object-outlook.md)** object, to the end of the **[NavigationFolders](navigationfolders-object-outlook.md)** collection.


## Syntax

 _expression_ . **Add**( **_Folder_** )

 _expression_ A variable that represents a **NavigationFolders** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Folder_|Required| **Folder**|The folder to add.|

### Return Value

A  **NavigationFolder** object that represents the new navigation folder.


## Remarks

A folder can only appear in one navigation group. When adding a  **Folder** object to a new navigation group, any references to that **Folder** are removed from any other navigation group of which it was previously a member.


## See also


#### Concepts


[NavigationFolders Object](navigationfolders-object-outlook.md)

