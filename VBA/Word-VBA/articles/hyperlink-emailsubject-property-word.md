---
title: Hyperlink.EmailSubject Property (Word)
keywords: vbawd10.chm161285106
f1_keywords:
- vbawd10.chm161285106
ms.prod: word
api_name:
- Word.Hyperlink.EmailSubject
ms.assetid: 8b019ae2-40da-b69c-8f0b-554724a770bd
ms.date: 06/08/2017
---


# Hyperlink.EmailSubject Property (Word)

Returns or sets the text string for the specified hyperlink's subject line. Read/write  **String** .


## Syntax

 _expression_ . **EmailSubject**

 _expression_ A variable that represents a **[Hyperlink](hyperlink-object-word.md)** object.


## Remarks

The subject line is appended to the hyperlink's Internet address, or URL. This property is commonly used with e-mail hyperlinks. The value of this property takes precedence over any e-mail subject specified in the  **[Address](hyperlink-address-property-word.md)** property of the same **Hyperlink** object.


## Example

This example checks the active document for e-mail hyperlinks; if it finds any that have a blank subject line, it adds the subject "NewProducts".


```vb
Dim hypLoop As Hyperlink 
 
For Each hypLoop In ActiveDocument.Hyperlinks 
 If hypLoop.Address Like "mailto*" And _ 
 hypLoop.Address = hypLoop.EmailSubject Then 
 hypLoop.EmailSubject = "NewProducts" 
 End If 
Next hypLoop
```


## See also


#### Concepts


[Hyperlink Object](hyperlink-object-word.md)

