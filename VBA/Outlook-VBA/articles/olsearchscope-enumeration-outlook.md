---
title: OlSearchScope Enumeration (Outlook)
keywords: vbaol11.chm3248
f1_keywords:
- vbaol11.chm3248
ms.prod: outlook
api_name:
- Outlook.OlSearchScope
ms.assetid: 13d19f0e-88f3-07d8-b048-87fc586e2e0c
ms.date: 06/08/2017
---


# OlSearchScope Enumeration (Outlook)

Specifies the scope in terms of folders for the search. 



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **olSearchScopeAllFolders**|1|The search scope is across all folders that have the same folder type as the current folder ( **[Folder.DefaultItemType](folder-defaultitemtype-property-outlook.md)** ), and all stores that have been selected for search.|
| **olSearchScopeAllOutlookItems**|2|The search scope is all Outlook items in all folders in stores that have been selected for search.|
| **olSearchScopeCurrentFolder**|0|The search scope is the folder represented by  **[Explorer.CurrentFolder](explorer-currentfolder-property-outlook.md)** , and only that folder.|
| **olSearchScopeCurrentStore**|4|The search scope is the store for the current folder, which contains the item displayed in the active explorer. |
| **olSearchScopeSubfolders**|3|The search scope is the folder represented by  **Explorer.CurrentFolder** and its subfolders.|

## Remarks

You can select stores to search in the  **Locations to Search** menu by clicking **Search Tools** in the **Options** group of the **Search** contextual tab in the Microsoft Office Fluent ribbon.

By default, search does not include the Deleted Items folder. To search the Deleted Items folder, set that folder as your current folder and search by  **olSearchScopeCurrentFolder**.


