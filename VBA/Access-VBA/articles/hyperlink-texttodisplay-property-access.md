---
title: Hyperlink.TextToDisplay Property (Access)
keywords: vbaac10.chm10120
f1_keywords:
- vbaac10.chm10120
ms.prod: access
api_name:
- Access.Hyperlink.TextToDisplay
ms.assetid: 61417274-e124-be4c-1b80-9d4600021326
ms.date: 06/08/2017
---


# Hyperlink.TextToDisplay Property (Access)

You can use the  **TextToDisplay** property to specify or determine the display text for a hyperlink. Read/write **String**.


## Syntax

 _expression_. **TextToDisplay**

 _expression_ A variable that represents a **Hyperlink** object.


## Example

The following example displays the words "Go to Home page" as an active hyperlink in the label named "Label20" on the "Suppliers" form. Clicking the hyperlink takes the user to the address specified in the label's  **HyperlinkAddress** property.


```vb
Forms.Item("Suppliers").Controls.Item("Label20").Hyperlink. _ 
 TextToDisplay = "Go to Home page"
```


## See also


#### Concepts


[Hyperlink Object](hyperlink-object-access.md)

