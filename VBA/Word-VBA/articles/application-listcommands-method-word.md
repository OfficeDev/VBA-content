---
title: Application.ListCommands Method (Word)
keywords: vbawd10.chm158335322
f1_keywords:
- vbawd10.chm158335322
ms.prod: word
api_name:
- Word.Application.ListCommands
ms.assetid: 425abd0f-c9c4-c4ab-b308-e7876ace5778
ms.date: 06/08/2017
---


# Application.ListCommands Method (Word)

Creates a new document and then inserts a table of Word commands along with their associated shortcut keys and menu assignments.


## Syntax

 _expression_ . **ListCommands**( **_ListAllCommands_** )

 _expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ListAllCommands_|Required| **Boolean**| **True** to include all Word commands and their assignments (whether customized or built-in). **False** to include only commands with customized assignments.|

## Example

This example creates a new document that lists all Word commands along with their associated shortcut keys and menu assignments. The example then prints and closes the new document without saving changes.


```vb
Application.ListCommands ListAllCommands:=True 
With ActiveDocument 
 .PrintOut 
 .Close SaveChanges:=wdDoNotSaveChanges 
End With
```


## See also


#### Concepts


[Application Object](application-object-word.md)

