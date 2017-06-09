---
title: Bookmarks.Exists Method (Word)
keywords: vbawd10.chm157745158
f1_keywords:
- vbawd10.chm157745158
ms.prod: word
api_name:
- Word.Bookmarks.Exists
ms.assetid: 7a9df80d-1a52-022f-f234-336369b73fca
ms.date: 06/08/2017
---


# Bookmarks.Exists Method (Word)

Determines whether the specified bookmark exists. Returns  **True** if the bookmark exists.


## Syntax

 _expression_ . **Exists**( **_Name_** )

 _expression_ A variable that represents a **[Bookmarks](bookmarks-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|A bookmark name than can not include more than 40 characters or more than one word.|

## Example

This example determines whether the bookmark named "start" exists in the active document. If the bookmark exists, it is deleted.


```vb
If ActiveDocument.Bookmarks.Exists("start") = True Then 
 ActiveDocument.Bookmarks("start").Delete 
End If
```


## See also


#### Concepts


[Bookmarks Collection Object](bookmarks-object-word.md)

