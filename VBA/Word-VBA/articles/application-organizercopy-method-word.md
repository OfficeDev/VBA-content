---
title: Application.OrganizerCopy Method (Word)
keywords: vbawd10.chm158335294
f1_keywords:
- vbawd10.chm158335294
ms.prod: word
api_name:
- Word.Application.OrganizerCopy
ms.assetid: a23452aa-7372-ca58-291f-164e6000162d
ms.date: 06/08/2017
---


# Application.OrganizerCopy Method (Word)

Copies the specified AutoText entry, toolbar, style, or macro project item from the source document or template to the destination document or template.


## Syntax

 _expression_ . **OrganizerCopy**( **_Source_** , **_Destination_** , **_Name_** , **_Object_** )

 _expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Source_|Required| **String**|The document or template file name that contains the item you want to copy.|
| _Destination_|Required| **String**|The document or template file name to which you want to copy an item.|
| _Name_|Required| **String**|The name of the AutoText entry, toolbar, style, or macro you want to copy.|
| _Object_|Required| **WdOrganizerObject**|The kind of item you want to copy.|

## Example

This example copies all the AutoText entries in the template attached to the active document to the Normal template.


```vb
Dim atEntry As AutoTextEntry 
 
For Each atEntry In _ 
 ActiveDocument.AttachedTemplate.AutoTextEntries 
 Application.OrganizerCopy _ 
 Source:=ActiveDocument.AttachedTemplate.FullName, _ 
 Destination:=NormalTemplate.FullName, Name:=atEntry.Name, _ 
 Object:=wdOrganizerObjectAutoText 
Next atEntry
```

If the style named "SubText" exists in the active document, this example copies the style to C:\Templates\Template1.dot.




```vb
Dim styleLoop As Style 
 
For Each styleLoop In ActiveDocument.Styles 
 If styleLoop = "SubText" Then 
 Application.OrganizerCopy Source:=ActiveDocument.Name, _ 
 Destination:="C:\Templates\Template1.dot", _ 
 Name:="SubText", _ 
 Object:=wdOrganizerObjectStyles 
 End If 
Next styleLoop
```


## See also


#### Concepts


[Application Object](application-object-word.md)

