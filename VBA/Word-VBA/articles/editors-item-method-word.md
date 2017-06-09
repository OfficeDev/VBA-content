---
title: Editors.Item Method (Word)
keywords: vbawd10.chm9175040
f1_keywords:
- vbawd10.chm9175040
ms.prod: word
api_name:
- Word.Editors.Item
ms.assetid: 58fee673-6162-37e3-803d-5fd0ce1fb144
ms.date: 06/08/2017
---


# Editors.Item Method (Word)

Returns an  **Editor** object that represents a specific user or a group of users who have been given permission to edit a portion of a document.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ Required. A variable that represents an **[Editors](editors-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**| Can be either a **String** that represents the user's e-mail alias (if in the same domain), an e-mail address, or a **WdEditorType** constant that represents a group of users.|

### Return Value

Editor


## See also


#### Concepts


[Editors Collection](editors-object-word.md)

