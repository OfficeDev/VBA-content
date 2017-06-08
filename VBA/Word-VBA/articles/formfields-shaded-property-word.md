---
title: FormFields.Shaded Property (Word)
keywords: vbawd10.chm153681922
f1_keywords:
- vbawd10.chm153681922
ms.prod: word
api_name:
- Word.FormFields.Shaded
ms.assetid: 816b0d24-7558-4e19-c390-791aefb29c65
ms.date: 06/08/2017
---


# FormFields.Shaded Property (Word)

 **True** if shading is applied to form fields. Read/write **Boolean** .


## Syntax

 _expression_ . **Shaded**

 _expression_ An expression that returns a **FormFields** collection object.


## Remarks

Shading makes form fields easier to locate in a document and doesn't affect the printed output.


## Example

This example removes shading from form fields in Employment Form.doc.


```vb
Documents("Employment Form.doc").FormFields.Shaded = False
```

This example adds shading to the form fields in the active document and protects the document for forms.




```vb
With ActiveDocument 
 .FormFields.Shaded = True 
 .Protect Type:=wdAllowOnlyFormFields, NoReset:=True 
End With
```


## See also


#### Concepts


[FormFields Collection Object](formfields-object-word.md)

