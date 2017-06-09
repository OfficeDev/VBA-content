---
title: RecentFiles.Add Method (Word)
keywords: vbawd10.chm157483011
f1_keywords:
- vbawd10.chm157483011
ms.prod: word
api_name:
- Word.RecentFiles.Add
ms.assetid: 6d20df76-9a7a-be22-2c11-44f328dee13a
ms.date: 06/08/2017
---


# RecentFiles.Add Method (Word)

Returns a  **RecentFile** object that represents a file added to the list of recently used files.


## Syntax

 _expression_ . **Add**( **_Document_** , **_ReadOnly_** )

 _expression_ Required. A variable that represents a **[RecentFiles](recentfiles-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Document_|Required| **Variant**|The document you want to add to the list of recently used files. You can specify this argument by using either the string name for the document or a  **Document** object.|
| _ReadOnly_|Optional| **Variant**| **True** to make the document read-only.|

### Return Value

RecentFile


## Example

This example adds the active document to the list of recently used files.


```vb
If ActiveDocument.Saved = True Then 
 RecentFiles.Add Document:=ActiveDocument.Name 
End If
```


## See also


#### Concepts


[RecentFiles Collection Object](recentfiles-object-word.md)

