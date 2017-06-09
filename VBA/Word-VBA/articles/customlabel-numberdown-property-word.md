---
title: CustomLabel.NumberDown Property (Word)
keywords: vbawd10.chm152371210
f1_keywords:
- vbawd10.chm152371210
ms.prod: word
api_name:
- Word.CustomLabel.NumberDown
ms.assetid: d2257e2f-2641-764c-d5a1-72a1fddb6f22
ms.date: 06/08/2017
---


# CustomLabel.NumberDown Property (Word)

Returns or sets the number of custom mailing labels down the length of a page. Read/write  **Long** .


## Syntax

 _expression_ . **NumberDown**

 _expression_ An expression that returns a **[CustomLabel](customlabel-object-word.md)** object.


## Remarks

If this property is changed to a value that isn't valid for the specified mailing label layout, an error occurs.


## Example

This example displays the number of labels across and down the page for the first custom label in the CustomLabels collection.


```
numAcr = Application.MailingLabel.CustomLabels(1).NumberAcross 
numDwn = Application.MailingLabel.CustomLabels(1).NumberDown 
MsgBox Prompt:= "Number of labels across " &; numAcr &; vbCr _ 
 &; "Number of labels down " &; numDwn &; vbCr , _ 
 Title:="Label Page Configuration"
```


## See also


#### Concepts


[CustomLabel Object](customlabel-object-word.md)

