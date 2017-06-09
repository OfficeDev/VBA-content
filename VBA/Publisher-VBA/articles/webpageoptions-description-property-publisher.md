---
title: WebPageOptions.Description Property (Publisher)
keywords: vbapb10.chm544771
f1_keywords:
- vbapb10.chm544771
ms.prod: publisher
api_name:
- Publisher.WebPageOptions.Description
ms.assetid: dfd18427-c70d-7232-191e-a6332a89c3fe
ms.date: 06/08/2017
---


# WebPageOptions.Description Property (Publisher)

Returns or sets a  **String** that represents the description of a Web page within a Web publication. Read/write.


## Syntax

 _expression_. **Description**

 _expression_A variable that represents a  **WebPageOptions** object.


## Example

This example sets the description for page two of the active Web publication.


```vb
Dim theWPO As WebPageOptions 
 
Set theWPO = ActiveDocument.Pages(2).WebPageOptions 
 
With theWPO 
 .Description = "Company Profile" 
End With
```


