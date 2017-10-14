---
title: Hyperlink.Address Property (Visio)
keywords: vis_sdr.chm15013065
f1_keywords:
- vis_sdr.chm15013065
ms.prod: visio
api_name:
- Visio.Hyperlink.Address
ms.assetid: 289cc2d6-d147-0ef9-e09e-12127c74072a
ms.date: 06/08/2017
---


# Hyperlink.Address Property (Visio)

Gets or sets the address for a shape's  **Hyperlink** object, the address to which the hyperlink navigates. Read/write.


## Syntax

 _expression_ . **Address**

 _expression_ A variable that represents a **Hyperlink** object.


### Return Value

String


## Remarks

Setting the  **Address** property for a **Hyperlink** object is equivalent to entering information in the **Address** box in the **Hyperlinks** dialog box ( **Insert** menu), or setting the result of the Address cell in the shape's Hyperlink. _name_ row in the ShapeSheet window.

The  **Address** property value can be a DOS, UNC, or URL path, for example, _"driveletter"_ :\ _"foldername"_ \ _"drawingname"_ , \\ _"servername"_ \ _"foldername"_ \ _"drawingname"_ , or http:// _address_ , respectively.

If the  **Address** property is relative, for example, "..\ _"drawingname"_ ", it is composed against the **HyperlinkBase** property, if supplied, or the hyperlink's document path. If the document is not saved, the hyperlink is undefined.

If the  **Address** property is empty, you can assume the address points to a page in the document that contains the page. In this case, the **SubAddress** property contains the name of the drawing page to which the hyperlink navigates.


## Example

This Microsoft Visual Basic for Applications (VBA) macro sets the  **Address** property of a **Hyperlink** object. Before running this example, replace _address_ with a valid Internet or intranet address.


```vb
 
Public Sub Address_Example() 
 
 Dim vsoShape As Visio.Shape 
 Dim vsoHyperlink As Visio.Hyperlink 
 
 'Create a new shape to receive the hyperlink 
 Set vsoShape = ActivePage.DrawRectangle(1, 2, 2, 1) 
 Set vsoHyperlink = vsoShape.AddHyperlink 
 
 vsoHyperlink.Description = "Web site" 
 vsoHyperlink.Address = "http://address /" 
 
End Sub
```


