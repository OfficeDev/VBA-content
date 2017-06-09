---
title: List.ApplyListTemplateWithLevel Method (Word)
keywords: vbawd10.chm160563307
f1_keywords:
- vbawd10.chm160563307
ms.prod: word
api_name:
- Word.List.ApplyListTemplateWithLevel
ms.assetid: 19c380d0-0599-72e2-b8ee-56ac7536d16c
ms.date: 06/08/2017
---


# List.ApplyListTemplateWithLevel Method (Word)

Applies a set of list-formatting characteristics, optionally for a specified level.


## Syntax

 _expression_ . **ApplyListTemplateWithLevel**( **_ListTemplate_** , **_ContinuePreviousList_** , **_DefaultListBehavior_** , **_ApplyLevel_** )

 _expression_ A variable that represents a **[List](list-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ListTemplate_|Required| **[ListTemplate](listtemplate-object-word.md)**|The list template to be applied.|
| _ContinuePreviousList_|Optional| **Variant**| **True** to continue the numbering from the previous list; **False** to start a new list.|
| _DefaultListBehavior_|Optional| **Variant**|Sets a value that specifies whether Microsoft Word uses new Web-oriented formatting for better list display. Can be either of the following  **[WdDefaultListBehavior](wddefaultlistbehavior-enumeration-word.md)** constants: **wdWord8ListBehavior** (use formatting compatible with Microsoft Word 97) or **wdWord9ListBehavior** (use Web-oriented formatting). For compatibility reasons, the default constant is **wdWord8ListBehavior** , but in new procedures you should use **wdWord9ListBehavior** to take advantage of improved Web-oriented formatting for indenting and multiple-level lists.|
| _ApplyLevel_|Optional| **Variant**|The level to which the list template is to be applied.|

## Example

The following example sets the variable  `myRange` to a range in the active document, and then it verifies whether the range has list formatting. If no list formatting has been applied, the fourth outline-numbered list template is applied to the range.


```vb
Set myDoc = ActiveDocument 
Set myRange = myDoc.Range( _ 
 Start:= myDoc.Paragraphs(3).Range.Start, _ 
 End:=myDoc.Paragraphs(6).Range.End) 
If myRange.ListFormat.ListType = wdListNoNumbering Then 
 myRange.ListFormat.ApplyListTemplate _ 
 ListTemplate:=ListGalleries(wdOutlineNumberGallery) _ 
 .ListTemplates(4) 
End If
```


## See also


#### Concepts


[List Object](list-object-word.md)

