---
title: Document.MasterPages Property (Publisher)
keywords: vbapb10.chm196629
f1_keywords:
- vbapb10.chm196629
ms.prod: publisher
api_name:
- Publisher.Document.MasterPages
ms.assetid: 26e5342b-94f0-4fd5-2743-92cfd2d43a01
ms.date: 06/08/2017
---


# Document.MasterPages Property (Publisher)

Returns the  **[MasterPages](masterpages-object-publisher.md)** collection for the specified publication.


## Syntax

 _expression_. **MasterPages**

 _expression_A variable that represents a  **Document** object.


### Return Value

MasterPages


## Example

The following example sets the text in the first text frame on the master page to Second Quarter.


```vb
Dim mp As MasterPages 
 
Set mp = ActiveDocument.MasterPages 
 
With mp.Item(1) 
 .Shapes(1).TextFrame.TextRange.Text = "Second Quarter" 
End With
```


