---
title: Document.Styles Property (Visio)
keywords: vis_sdr.chm10514455
f1_keywords:
- vis_sdr.chm10514455
ms.prod: visio
api_name:
- Visio.Document.Styles
ms.assetid: 41434c49-3306-78b5-2126-0320fc05825a
ms.date: 06/08/2017
---


# Document.Styles Property (Visio)

Returns the  **Styles** collection for a document. Read-only.


## Syntax

 _expression_ . **Styles**

 _expression_ A variable that represents a **Document** object.


### Return Value

Styles


## Example

The following macro shows how to use the  **Styles** property to add **Style** objects to the **Styles** collection. It shows how to add a new style based on an existing style as well as how to add a new style created from scratch.


```vb
 
Public Sub Styles_Example() 
 
 Dim vsoDocument As Visio.Document 
 Dim vsoStyles As Visio.Styles 
 Dim vsoStyle As Visio.Style 
 
 'Add a document based on the Basic Diagram template. 
 Set vsoDocument = Documents.Add("Basic Diagram.vst") 
 
 'Add a style named "My FillStyle" to the Styles collection. 
 'This style is based on the style Basic Shadow and includes 
 'only a Fill style. 
 Set vsoStyles = vsoDocument.Styles 
 Set vsoStyle = vsoStyles.Add("My FillStyle", "Basic Shadow", 0, 0, 1) 
 
 'Add a style named "My NoStyle" to the Styles collection. 
 'This style is not based on an existing style and includes 
 'Text, Line, and Fill styles. 
 Set vsoStyle = vsoStyles.Add("My NoStyle", "", 1, 1, 1) 
 
End Sub
```


