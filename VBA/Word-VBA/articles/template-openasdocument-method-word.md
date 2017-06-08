---
title: Template.OpenAsDocument Method (Word)
keywords: vbawd10.chm157941860
f1_keywords:
- vbawd10.chm157941860
ms.prod: word
api_name:
- Word.Template.OpenAsDocument
ms.assetid: 3e73bddd-a5af-5c58-cd66-3271271633dd
ms.date: 06/08/2017
---


# Template.OpenAsDocument Method (Word)

Opens the specified template as a document and returns a  **Document** object.


## Syntax

 _expression_ . **OpenAsDocument**

 _expression_ Required. A variable that represents a **[Template](template-object-word.md)** object.


### Return Value

Document


## Remarks

Opening a template as a document allows the user to edit the contents of the template. This may be necessary if a property or method (the  **Styles** property, for example) isn't available from the **Template** object.


## Example

This example opens the template attached to the active document, displays a message box if the template contains anything more than a single paragraph mark, and then closes the template.


```vb
Dim docNew As Document 
 
Set docNew = ActiveDocument.AttachedTemplate.OpenAsDocument 
 
If docNew.Content.Text <> Chr(13) Then 
 MsgBox "Template is not empty" 
Else 
 MsgBox "Template is empty" 
End If 
docNew.Close SaveChanges:=wdDoNotSaveChanges
```

This example saves a copy of the Normal template as "Backup.dot."




```vb
Dim docNew As Document 
 
Set docNew = NormalTemplate.OpenAsDocument 
 
With docNew 
 .SaveAs FileName:="Backup.dot" 
 .Close SaveChanges:=wdDoNotSaveChanges 
End With
```

This example changes the formatting of the Heading 1 style in the template attached to the active document. The  **UpdateStyles** method updates the styles in the active document.




```vb
Dim docNew As Document 
 
Set docNew = ActiveDocument.AttachedTemplate.OpenAsDocument 
 
With docNew.Styles(wdStyleHeading1).Font 
 .Name = "Arial" 
 .Size = 16 
 .Bold = False 
End With 
docNew.Close SaveChanges:=wdSaveChanges 
ActiveDocument.UpdateStyles
```


## See also


#### Concepts


[Template Object](template-object-word.md)

