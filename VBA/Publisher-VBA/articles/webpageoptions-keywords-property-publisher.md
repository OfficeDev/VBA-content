---
title: WebPageOptions.Keywords Property (Publisher)
keywords: vbapb10.chm544772
f1_keywords:
- vbapb10.chm544772
ms.prod: publisher
api_name:
- Publisher.WebPageOptions.Keywords
ms.assetid: 8dd7b073-747e-a6f6-a20d-0b3e3d9a27b8
ms.date: 06/08/2017
---


# WebPageOptions.Keywords Property (Publisher)

Returns or sets a  **String** that represents the keywords for a Web page within a Web publication. Read/write.


## Syntax

 _expression_. **Keywords**

 _expression_A variable that represents a  **WebPageOptions** object.


### Return Value

String


## Example

The following example sets the keywords for page four of the active publication.


```vb
Dim theWPO As WebPageOptions 
 
Set theWPO = ActiveDocument.Pages(4).WebPageOptions 
 
With theWPO 
 .Keywords = "software, hardware, computers" 
End With
```


