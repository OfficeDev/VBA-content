---
title: MailMergeFields.AddFillIn Method (Word)
keywords: vbawd10.chm153026663
f1_keywords:
- vbawd10.chm153026663
ms.prod: word
api_name:
- Word.MailMergeFields.AddFillIn
ms.assetid: aefd78e5-3439-473c-1b9b-7f58a3a45d55
ms.date: 06/08/2017
---


# MailMergeFields.AddFillIn Method (Word)

Adds a FILLIN field to a mail merge main document. Returns a  **MailMergeField** object.


## Syntax

 _expression_ . **AddFillIn**( **_Range_** , **_Prompt_** , **_DefaultFillInText_** , **_AskOnce_** )

 _expression_ Required. A variable that represents a **[MailMergeFields](mailmergefields-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range object**|The location for the FILLIN field.|
| _Prompt_|Optional| **Variant**|The text that's displayed in the dialog box.|
| _DefaultFillinText_|Optional| **Variant**|The default response, which appears in the text box when the dialog box is displayed. Corresponds to the \d switch for an FILLIN field.|
| _AskOnce_|Optional| **Variant**| **True** to display the prompt only once instead of each time a new record is merged. Corresponds to the \o switch for a FILLIN field. The default value is **False** .|

### Return Value

MailMergeField


## Remarks

When updated, a FILLIN field displays a dialog box that prompts you for text to insert into the document at the location of the FILLIN field. Use the  **Add** method with the **Fields** collection object to add a FILLIN field to a document other than a mail merge main document.


## Example

This example adds a FILLIN field that prompts you for a name to insert after "Name:".


```vb
With Selection 
 .Collapse Direction:=wdCollapseStart 
 .InsertAfter "Name: " 
 .Collapse Direction:=wdCollapseEnd 
End With 
ActiveDocument.MailMerge.Fields.AddFillin Range:=Selection.Range, _ 
 Prompt:="Your name?", DefaultFillInText:="Joe", AskOnce:=True
```


## See also


#### Concepts


[MailMergeFields Collection Object](mailmergefields-object-word.md)

