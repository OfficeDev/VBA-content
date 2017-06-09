---
title: Document.VBProject Property (Word)
keywords: vbawd10.chm158007395
f1_keywords:
- vbawd10.chm158007395
ms.prod: word
api_name:
- Word.Document.VBProject
ms.assetid: bf9d4c60-8e7a-b076-b20c-0021e9352273
ms.date: 06/08/2017
---


# Document.VBProject Property (Word)

Returns the  **VBProject** object for the specified template or document.


## Syntax

 _expression_ . **VBProject**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

Use this property to gain access to code modules and user forms.

To view the  **VBProject** object in the object browser, you must select the **Microsoft Visual Basic for Applications Extensibility** check box in the **References** dialog box ( **Tools** menu) in the Visual Basic Editor.


## Example

This example displays the name of the Visual Basic project for the active document.


```vb
Set currProj = ActiveDocument.VBProject 
MsgBox currProj.Name
```

This example adds a standard code module to the active document and renames it "MyModule."




```vb
Set newModule = ActiveDocument.VBProject.VBComponents _ 
 .Add(vbext_ct_StdModule) 
NewModule.Name = "MyModule"
```


## See also


#### Concepts


[Document Object](document-object-word.md)

