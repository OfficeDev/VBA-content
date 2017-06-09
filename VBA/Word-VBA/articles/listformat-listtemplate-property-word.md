---
title: ListFormat.ListTemplate Property (Word)
keywords: vbawd10.chm163577926
f1_keywords:
- vbawd10.chm163577926
ms.prod: word
api_name:
- Word.ListFormat.ListTemplate
ms.assetid: 778f4b21-575c-b6b1-768a-735c4730ae13
ms.date: 06/08/2017
---


# ListFormat.ListTemplate Property (Word)

Returns a  **ListTemplate** object that represents the list formatting for the specified **ListFormat** object.


## Syntax

 _expression_ . **ListTemplate**

 _expression_ Required. A variable that represents a **[ListFormat](listformat-object-word.md)** object.


## Remarks

A list template includes all the formatting that defines a particular list. Each of the seven formats (excluding  **None**) found on each of the tabs in the  **Bullets and Numbering** dialog box corresponds to a list template. Documents and templates can also contain collections of list templates.

If the first paragraph in the range for the  **ListFormat** object is not formatted as a list, the **ListTemplate** property returns **Nothing** .


## Example

This example checks to see which list template is used for the second paragraph in the active document, and then it applies that list template to the selection.


```vb
Set myltemp = ActiveDocument.Paragraphs(2).Range. _ 
 ListFormat.ListTemplate 
Selection.Range.ListFormat.ApplyListTemplate ListTemplate:=myltemp
```


## See also


#### Concepts


[ListFormat Object](listformat-object-word.md)

