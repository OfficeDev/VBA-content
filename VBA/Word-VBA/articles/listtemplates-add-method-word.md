---
title: ListTemplates.Add Method (Word)
keywords: vbawd10.chm160432228
f1_keywords:
- vbawd10.chm160432228
ms.prod: word
api_name:
- Word.ListTemplates.Add
ms.assetid: cb5ad343-fbcc-22f0-6a05-83f1480da691
ms.date: 06/08/2017
---


# ListTemplates.Add Method (Word)

Returns a  **ListTemplate** object that represents a new list template.


## Syntax

 _expression_ . **Add**( **_OutlineNumbered_** , **_Name_** )

 _expression_ Required. A variable that represents a **[ListTemplates](listtemplates-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _OutlineNumbered_|Optional| **Variant**| **True** to apply outline numbering to the new list template.|
| _Name_|Optional| **Variant**|An optional name used for linking the list template to a LISTNUM field. You can use this name to index the list template in the collection.|

### Return Value

ListTemplate


## Remarks

You cannot use the  **Add** method on **ListTemplates** objects returned from a **ListGallery** object. You can, however, modify the existing list templates in the galleries.


## Example

This example adds a new, single-level list template to the active document. The example changes the numbering style for the new list template and then applies the list template to the selection.


```vb
Set myList = _ 
 ActiveDocument.ListTemplates.Add(OutlineNumbered:=False) 
myList.ListLevels(1).NumberStyle = wdListNumberStyleUpperCaseLetter 
Selection.Range.ListFormat.ApplyListTemplate ListTemplate:=myList
```


## See also


#### Concepts


[ListTemplates Collection Object](listtemplates-object-word.md)

