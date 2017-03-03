---
title: Documents.Item Method (Word)
keywords: vbawd10.chm158072832
f1_keywords:
- vbawd10.chm158072832
ms.prod: WORD
api_name:
- Word.Documents.Item
ms.assetid: 0777c075-b466-3ac9-312a-4e1da7c1a732
---


# Documents.Item Method (Word)

Returns an individual  **Document** object in a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ Required. A variable that represents a **[Documents](documents-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The individual object to be returned. Can be a  **Long** indicating the ordinal position of the individual object.|

### Return Value

Document


## Example

This example displays the name of the first document in the  **Documents** collection.


```vb
Sub DocumentItem() 
 If Documents.Count >= 1 Then 
 MsgBox Documents.Item(1).Name 
 End If 
End Sub
```


## See also


#### Concepts


[Documents Collection Object](documents-object-word.md)

