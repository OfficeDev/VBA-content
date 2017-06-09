---
title: Application.OrganizerDelete Method (Word)
keywords: vbawd10.chm158335295
f1_keywords:
- vbawd10.chm158335295
ms.prod: word
api_name:
- Word.Application.OrganizerDelete
ms.assetid: 45b394fc-cdd5-18ff-f30d-7339237a1b41
ms.date: 06/08/2017
---


# Application.OrganizerDelete Method (Word)

Deletes the specified style, AutoText entry, toolbar, or macro project item from a document or template.


## Syntax

 _expression_ . **OrganizerDelete**( **_Source_** , **_Name_** , **_Object_** )

 _expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Source_|Required| **String**|The file name of the document or template that contains the item you want to delete.|
| _Name_|Required| **String**|The name of the style, AutoText entry, toolbar, or macro you want to delete.|
| _Object_|Required| **WdOrganizerObject**|The kind of item you want to copy.|

## Example

This example deletes the toolbar named "Custom 1" from the Normal template.


```vb
Dim cbLoop As CommandBar 
 
For Each cbLoop In CommandBars 
 If cbLoop.Name = "Custom 1" Then 
 Application.OrganizerDelete Source:=NormalTemplate.Name, _ 
 Name:="Custom 1", _ 
 Object:=wdOrganizerObjectCommandBars 
 End If 
Next cbLoop
```

This example prompts the user to delete each AutoText entry in the template attached to the active document. If the user clicks the Yes button, the AutoText entries are deleted.




```vb
Dim atEntry As AutoTextEntry 
Dim intResponse As Integer 
 
For Each atEntry In _ 
 ActiveDocument.AttachedTemplate.AutoTextEntries 
 intResponse = _ 
 MsgBox("Do you want to delete the " &; atEntry.Name _ 
 &; " AutoText entry?", vbYesNoCancel) 
 If intResponse = vbYes Then 
 With ActiveDocument.AttachedTemplate 
 Application.OrganizerDelete _ 
 Source:= .Path &; "\" &; .Name, _ 
 Name:=atEntry.Name, _ 
 Object:=wdOrganizerObjectAutoText 
 End With 
 ElseIf intResponse = vbCancel Then 
 Exit For 
 End If 
Next atEntry
```


## See also


#### Concepts


[Application Object](application-object-word.md)

