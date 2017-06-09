---
title: Document.SurplusShapes Property (Publisher)
keywords: vbapb10.chm196754
f1_keywords:
- vbapb10.chm196754
ms.prod: publisher
api_name:
- Publisher.Document.SurplusShapes
ms.assetid: 8c1c5fee-bea0-1660-a4a5-b465879d6ec9
ms.date: 06/08/2017
---


# Document.SurplusShapes Property (Publisher)

Returns a  **ShapeRange** object that represents the collection of surplus shapes that Microsoft Publisher places under **Extra Content**in the  **Format Publication** task pane after the document template (wizard) is changed by using the ** [Document.ChangeDocument](document-changedocument-method-publisher.md)** method or by using the **Change Template** command in the user interface. Read-only.


## Syntax

 _expression_. **SurplusShapes**

 _expression_A variable that represents a  **Document** object.


### Return Value

ShapeRange


## Remarks

Publisher classifies a shape as surplus if it does not fit neatly into the new template after the template is changed.


