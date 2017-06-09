---
title: Application.CheckGrammar Method (Word)
keywords: vbawd10.chm158335299
f1_keywords:
- vbawd10.chm158335299
ms.prod: word
api_name:
- Word.Application.CheckGrammar
ms.assetid: 4675bda9-c31d-efdc-4def-38bfdeb200e4
ms.date: 06/08/2017
---


# Application.CheckGrammar Method (Word)

Checks a string for grammatical errors. Returns a  **Boolean** to indicate whether the string contains grammatical errors. **True** if the string contains no errors.


## Syntax

 _expression_ . **CheckGrammar**( **_String_** )

 _expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _String_|Required| **String**|The string you want to check for grammatical errors.|

### Return Value

Boolean


## Example

This example displays the result of a grammar check on the selection.


```
strPass = Application.CheckGrammar(String:=Selection.Text) 
MsgBox "Selection is grammatically correct: " &; strPass
```


## See also


#### Concepts


[Application Object](application-object-word.md)

