---
title: CustomLabel.Valid Property (Word)
keywords: vbawd10.chm152371213
f1_keywords:
- vbawd10.chm152371213
ms.prod: word
api_name:
- Word.CustomLabel.Valid
ms.assetid: dbc744fc-bf6d-bc1e-5cd2-a5dd97619593
ms.date: 06/08/2017
---


# CustomLabel.Valid Property (Word)

 **True** if the various properties (for example, **Height** , **Width** , and **NumberDown** ) for the specified custom label work together to produce a valid mailing label. Read-only **Boolean** .


## Syntax

 _expression_ . **Valid**

 _expression_ A variable that represents a **[CustomLabel](customlabel-object-word.md)** object.


## Example

If the settings for the custom label named "My Labels" are valid, this example creates a new document of labels using the My Labels settings.


```vb
addr = "James Allard" &; vbCr &; "123 Main St." &; vbCr _ 
 &; "Seattle, WA 98040" 
If Application.MailingLabel.CustomLabels("My Labels") _ 
 .Valid = True Then Application.MailingLabel.CreateNewDocument _ 
 Name:="My Labels", Address:=addr 
End If
```


## See also


#### Concepts


[CustomLabel Object](customlabel-object-word.md)

