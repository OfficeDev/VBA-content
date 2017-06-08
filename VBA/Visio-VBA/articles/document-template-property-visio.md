---
title: Document.Template Property (Visio)
keywords: vis_sdr.chm10514505
f1_keywords:
- vis_sdr.chm10514505
ms.prod: visio
api_name:
- Visio.Document.Template
ms.assetid: c9e579d7-4448-4dc7-0130-1b38d41cbf1a
ms.date: 06/08/2017
---


# Document.Template Property (Visio)

Returns the name of the template from which the document was created. Read-only.


## Syntax

 _expression_ . **Template**

 _expression_ A variable that represents a **Document** object.


### Return Value

String


## Remarks

If the document is based on no template, the  **Template** property returns an empty string (''").


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Template** property to get the name of the template from which the document was created.


```vb
 
Public Sub Template_Example() 
 
 Dim strTemplateName As String 
 
 strTemplateName = ThisDocument.Template 
 
 'Verify that the proper string was returned. 
 Debug.Print strTemplateName 
 
End Sub
```


