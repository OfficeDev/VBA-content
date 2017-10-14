---
title: Document.IsInAutosave Property (Word)
keywords: vbawd10.chm158007915
f1_keywords:
- vbawd10.chm158007915
ms.prod: word
ms.assetid: 89438dfd-3b5a-e90b-5059-a62f1e47afeb
ms.date: 06/08/2017
---


# Document.IsInAutosave Property (Word)

Returns  **False** if the most recent firing of the[Application.DocumentBeforeSave Event (Word)](application-documentbeforesave-event-word.md) event was the result of a manual save by the user, and not an automatic save. Read-only.


## Syntax

 _expression_ . **IsInAutosave**

 _expression_ A variable that represents a **Document** object.


## Property value

 **BOOLEAN**


## Remarks

The  **IsInAutosave** property is designed to be used in an event handler for the **Application.DocumentBeforeSave** event. Using it for other purposes is not recommended.


 **Note**  In Visual Basic for Applications (VBA), the  **True** keyword has a value of negative one (-1). The **IsInAutosave** property, however, returns positive one (1) for **True**, rather than -1. As a result,  **IsInAutosave** never returns the VBA **True** keyword. To determine if the property is true, use code similar to that shown in the following code sample. If you determine that **IsInAutosave** is true, you can safely avoid doing any of the additional processing you might normally do for a manual save operation.


## Example

Use the following code to determine if  **IsInAutosave** is true:


```vb
If Word.ActiveDocument.IsInAutosave = False Then
   Debug.Print "Manual save."
Else
   Debug.Print "Automatic save."
End If
```


## See also


#### Concepts


[Document Object](document-object-word.md)

