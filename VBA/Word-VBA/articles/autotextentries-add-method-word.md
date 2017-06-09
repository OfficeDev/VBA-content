---
title: AutoTextEntries.Add Method (Word)
keywords: vbawd10.chm154599525
f1_keywords:
- vbawd10.chm154599525
ms.prod: word
api_name:
- Word.AutoTextEntries.Add
ms.assetid: 7ffa87f9-a23c-1847-3907-84c95f2b7f73
ms.date: 06/08/2017
---


# AutoTextEntries.Add Method (Word)

Returns an  **AutoTextEntry** object that represents an AutoText entry added to the list of available AutoText entries.


## Syntax

 _expression_ . **Add**( **_Name_** , **_Range_** )

 _expression_ Required. A variable that represents an **[AutoTextEntries](autotextentries-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The text that, when typed, initiates an AutoText entry.|
| _Range_|Required| **Range**|A range of text that will be inserted whenever Name is typed.|

## Example

This example adds an AutoText entry named Sample Text that contains the text in the selection. This example assumes you have text selected in the active document.


```vb
Sub AutoTxt() 
 NormalTemplate.AutoTextEntries.Add Name:="Sample Text", _ 
 Range:=Selection.Range 
End Sub
```


## See also


#### Concepts


[AutoTextEntries Collection Object](autotextentries-object-word.md)

