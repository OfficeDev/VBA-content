---
title: Application.OrganizerRename Method (Word)
keywords: vbawd10.chm158335296
f1_keywords:
- vbawd10.chm158335296
ms.prod: word
api_name:
- Word.Application.OrganizerRename
ms.assetid: abbe323c-b882-e497-608f-80004e166c8a
ms.date: 06/08/2017
---


# Application.OrganizerRename Method (Word)

Renames the specified style, AutoText entry, toolbar, or macro project item in a document or template.


## Syntax

 _expression_ . **OrganizerRename**( **_Source_** , **_Name_** , **_NewName_** , **_Object_** )

 _expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Source_|Required| **String**|The file name of the document or template that contains the item you want to rename.|
| _Name_|Required| **String**|The name of the style, AutoText entry, toolbar, or macro you want to rename.|
| _NewName_|Required| **String**|The new name for the item.|
| _Object_|Required| **WdOrganizerObject**|The kind of item you want to copy.|

## Example

This example changes the name of the style named "SubText" in the active document to "SubText2."


```vb
Dim styleLoop as Style 
 
For Each styleLoop In ActiveDocument.Styles 
 If styleLoop.NameLocal = "SubText" Then 
 Application.OrganizerRename _ 
 Source:=ActiveDocument.Name, Name:="SubText", _ 
 NewName:="SubText2", _ 
 Object:=wdOrganizerObjectStyles 
 End If 
Next styleLoop
```

This example changes the name of the macro module named "Module1" in the attached template to "Macros1."




```vb
Dim dotTemp As Template 
 
dotTemp = ActiveDocument.AttachedTemplate.Name 
Application.OrganizerRename Source:=dotTemp, Name:="Module1", _ 
 NewName:="Macros1", Object:=wdOrganizerObjectProjectItems
```


## See also


#### Concepts


[Application Object](application-object-word.md)

