---
title: Page.WebPageOptions Property (Publisher)
keywords: vbapb10.chm393264
f1_keywords:
- vbapb10.chm393264
ms.prod: publisher
api_name:
- Publisher.Page.WebPageOptions
ms.assetid: c2e3ee01-5b49-e83c-a68b-a4d526da0215
ms.date: 06/08/2017
---


# Page.WebPageOptions Property (Publisher)

Returns a  **[WebPageOptions](webpageoptions-object-publisher.md)** object, which represents the properties of a single Web page within a Web publication. Read-only.


## Syntax

 _expression_. **WebPageOptions**

 _expression_A variable that represents a  **Page** object.


### Return Value

WebPageOptions


## Example

The following example sets the description and the background sound for the fourth page of the active Web publication.


```vb
With ActiveDocument.Pages(4).WebPageOptions 
 .Description = "Company Profile" 
 .BackgroundSound = "C:\CompanySounds\corporate_jingle.wav" 
End With 

```


