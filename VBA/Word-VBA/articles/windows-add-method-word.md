---
title: Windows.Add Method (Word)
keywords: vbawd10.chm157351946
f1_keywords:
- vbawd10.chm157351946
ms.prod: word
api_name:
- Word.Windows.Add
ms.assetid: ce201ef7-db0a-b697-815d-bb2cd670f4ad
ms.date: 06/08/2017
---


# Windows.Add Method (Word)

Returns a  **Window** object that represents a new window of a document.


## Syntax

 _expression_ . **Add**( **_Window_** )

 _expression_ Required. A variable that represents a **[Windows](windows-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Window_|Optional| **Variant**|The  **Window** object you want to open another window for. If this argument is omitted, a new window is opened for the active document.|

### Return Value

Window


## Remarks

A colon (:) and a number appear in the window caption when more than one window is open for the document.


## Example

This example opens a new window for the document that's displayed in the active window.


```
Windows.Add
```

This example opens a new window for MyDoc.doc.




```
Windows.Add Window:=Documents("MyDoc.doc").Windows(1)
```


## See also


#### Concepts


[Windows Collection Object](windows-object-word.md)

