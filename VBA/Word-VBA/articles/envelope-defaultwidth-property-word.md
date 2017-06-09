---
title: Envelope.DefaultWidth Property (Word)
keywords: vbawd10.chm152567815
f1_keywords:
- vbawd10.chm152567815
ms.prod: word
api_name:
- Word.Envelope.DefaultWidth
ms.assetid: 2b593322-0959-a4a4-8607-65e2f9e91f7b
ms.date: 06/08/2017
---


# Envelope.DefaultWidth Property (Word)

Returns or sets the default envelope width, in points. Read/write  **Single** .


## Syntax

 _expression_ . **DefaultWidth**

 _expression_ A variable that represents a **[Envelope](envelope-object-word.md)** object.


## Remarks

If you set the  **[DefaultHeight](envelope-defaultheight-property-word.md)** or **DefaultWidth** property, the envelope size is automatically changed to **Custom Size** in the **Envelope Options** dialog box ( **Tools** menu). Use the **[DefaultSize](envelope-defaultsize-property-word.md)** property to set the default size to a predefined size.


## Example

This example sets the default custom envelope width and height and adds an envelope to the active document.


```vb
Dim strAddress As String 
Dim strReturn As String 
 
strAddress = "Tim O' Brien " &; vbCr &; "123 Skye St." _ 
 &; vbCr &; "Bellevue, WA 98004" 
strReturn = "Dave Edson" &; vbCr &; "123 West Main" _ 
 &; vbCr &; "Seattle, WA 98004" 
 
With ActiveDocument.Envelope 
 .DefaultWidth = InchesToPoints(9) 
 .DefaultHeight = InchesToPoints(3.85) 
End With 
 
ActiveDocument.Envelope.Insert _ 
 Address:=strAddress, ReturnAddress:=strReturn
```


## See also


#### Concepts


[Envelope Object](envelope-object-word.md)

