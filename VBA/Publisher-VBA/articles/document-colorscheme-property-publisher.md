---
title: Document.ColorScheme Property (Publisher)
keywords: vbapb10.chm196614
f1_keywords:
- vbapb10.chm196614
ms.prod: publisher
api_name:
- Publisher.Document.ColorScheme
ms.assetid: b7748b48-eff3-bdf0-e6ce-a9a2e788d0f7
ms.date: 06/08/2017
---


# Document.ColorScheme Property (Publisher)

Returns or sets the  **[ColorScheme](colorscheme-object-publisher.md)** object that represents the scheme colors for the specified publication. Read/write.


## Syntax

 _expression_. **ColorScheme**

 _expression_A variable that represents a  **Document** object.


### Return Value

ColorScheme


## Example

This example displays the name of the current color scheme for the active publication.


```vb
With ActiveDocument.ColorScheme 
 MsgBox "The current color scheme is " &; .Name &; "." 
End With
```

This example sets the color scheme of the active publication to "Alpine."




```vb
ActiveDocument.ColorScheme _ 
 = Application.ColorSchemes("Alpine")
```


