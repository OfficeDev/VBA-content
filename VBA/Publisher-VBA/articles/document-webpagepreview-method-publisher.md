---
title: Document.WebPagePreview Method (Publisher)
keywords: vbapb10.chm196724
f1_keywords:
- vbapb10.chm196724
ms.prod: publisher
api_name:
- Publisher.Document.WebPagePreview
ms.assetid: 44083fae-d21d-9cd3-3553-a4d4346141f5
ms.date: 06/08/2017
---


# Document.WebPagePreview Method (Publisher)

Generates a Web page preview of the specified publication in Internet Explorer.


## Syntax

 _expression_. **WebPagePreview**

 _expression_A variable that represents a  **Document** object.


## Remarks

A Web preview can be generated for print publications. However, the appearance of the Web preview may differ from the printed publication.

The Web preview opens with the active page displayed. Preview Web pages are generated for each page in the publication. However, if the publication is a print publication or otherwise lacks a navigation bar, there may be no way to navigate to those pages.

Use the  **[PublicationType](document-publicationtype-property-publisher.md)** property to determine if a publication is a print publication or a Web publication.

This method corresponds to the  **Web Page Preview** command on the **File** menu.


## Example

The following example sets the active page of the publication and generates a Web preview of the publication.


```vb
 
With ActiveDocument 
 .ActiveView.ActivePage = .Pages(2) 
 .WebPagePreview 
End With
```


