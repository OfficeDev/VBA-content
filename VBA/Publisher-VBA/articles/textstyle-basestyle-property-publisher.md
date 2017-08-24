---
title: TextStyle.BaseStyle Property (Publisher)
keywords: vbapb10.chm5963783
f1_keywords:
- vbapb10.chm5963783
ms.prod: publisher
api_name:
- Publisher.TextStyle.BaseStyle
ms.assetid: c8d1665c-c232-ecdf-3c1c-f614c7374c1e
ms.date: 06/08/2017
---


# TextStyle.BaseStyle Property (Publisher)

Returns or sets a  **String** that represents the style upon which the formatting of another style is based. Read/write.


## Syntax

 _expression_. **BaseStyle**

 _expression_A variable that represents a  **TextStyle** object.


### Return Value

String


## Example

This example sets the base formatting of the style named Body Text to the formatting of the Normal style.


```vb
Sub SetBaseStyle() 
 With ActiveDocument.TextStyles 
 .Add "Body Text" 
 .Item("Body Text").BaseStyle = "Normal" 
 End With 
End Sub
```


