---
title: Section.ProtectedForForms Property (Word)
keywords: vbawd10.chm156827771
f1_keywords:
- vbawd10.chm156827771
ms.prod: word
api_name:
- Word.Section.ProtectedForForms
ms.assetid: f87ef960-9ef3-f5a8-c3e0-325c263e985b
ms.date: 06/08/2017
---


# Section.ProtectedForForms Property (Word)

 **True** if the specified section is protected for forms. Read/write **Boolean** .


## Syntax

 _expression_ . **ProtectedForForms**

 _expression_ An expression that returns a **[Section](section-object-word.md)** object.


## Remarks

When a section is protected for forms, you can select and modify text only in form fields. To protect an entire document, use the  **[Protect](http://msdn.microsoft.com/library/8269051e-7952-7dc0-c7d8-cbf5ff711e38%28Office.15%29.aspx)** method of the **[Document](document-object-word.md)** object.


## Example

This example turns on form protection for the second section in the active document.


```vb
If ActiveDocument.Sections.Count >= 2 Then _ 
 ActiveDocument.Sections(2).ProtectedForForms = True
```

This example unprotects the first section in the selection.




```vb
Selection.Sections(1).ProtectedForForms = False
```

This example toggles the protection for the first section in the selection.




```vb
Selection.Sections(1).ProtectedForForms = Not _ 
 Selection.Sections(1).ProtectedForForms
```


## See also


#### Concepts


[Section Object](section-object-word.md)

