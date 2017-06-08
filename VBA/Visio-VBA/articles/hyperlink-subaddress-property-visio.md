---
title: Hyperlink.SubAddress Property (Visio)
keywords: vis_sdr.chm15014460
f1_keywords:
- vis_sdr.chm15014460
ms.prod: visio
api_name:
- Visio.Hyperlink.SubAddress
ms.assetid: e384fd34-7696-042d-12a3-a2aae949ce43
ms.date: 06/08/2017
---


# Hyperlink.SubAddress Property (Visio)

Gets or sets the subaddress in a shape's  **Hyperlink** object. Read/write.


## Syntax

 _expression_ . **SubAddress**

 _expression_ A variable that represents a **Hyperlink** object.


### Return Value

String


## Remarks

Setting the  **SubAddress** property of a shape's **Hyperlink** object is optional unless the **Address** property is blank. In this case the **SubAddress** must contain the name of the drawing page.

Setting a hyperlink's  **SubAddress** property is equivalent to entering information in the **Sub-address** box in the **Hyperlinks** dialog box (on the **Insert** tab, click **Hyperlink**). This is also equivalent to setting the result of the SubAddress cell in the shape's Hyperlink. _name_ row in the ShapeSheet window.

The  **SubAddress** property for a **Hyperlink** object specifies a sublocation within the hyperlink's address. For Microsoft Visio files, this can be a page name. For Microsoft Excel, this can be a worksheet or a range within a worksheet. For HTML pages, this can be a subanchor.

The hyperlink address for which a subaddress is being supplied must support subaddress linking.


## Example

The following example shows how to use the  **SubAddress** property to set the subaddress of a hyperlink. Before running this macro, replace _drive\ folder\subfolder_ with a valid path on your computer, replace _address_ with a valid Internet or intranet address, replace _subaddress_ with a valid subaddress for the Internet or intranet address, replace _drawing.vsd_ with a valid file on your computer, and replace _anchor_ with a valid page and shape in the file.


```vb
 
Sub SubAddress_Example() 
 
 Dim vsoShape As Visio.Shape 
 Dim vsoHyperlink As Visio.Hyperlink 
 
 'Draw a rectangle shape on the active page. 
 Set vsoShape = ActivePage.DrawRectangle(1, 2, 2, 1) 
 
 'Add a hyperlink to the shape. 
 Set vsoHyperlink = vsoShape.AddHyperlink 
 
 'Allow relative hyperlink addresses. 
 ActiveDocument.HyperlinkBase = "drive:\folder\subfolder " 
 
 'Return a relative address. 
 vsoHyperlink.Address = "..\drawing.vsd " 
 
 'Return a relative subaddress. 
 vsoHyperlink.SubAddress = "anchor " 
 
 'Print the resulting URLs to the Debug window 
 'to show how the relative path is derived 
 'from the base path and the difference 
 'between canonical and noncanonical forms. 
 Debug.Print vsoHyperlink.CreateURL(False) 
 Debug.Print vsoHyperlink.CreateURL(True) 
 
 'Return an absolute address. 
 vsoHyperlink.Address = "http://address " 
 
 'Return an absolute subaddress. 
 vsoHyperlink.SubAddress = "../subaddress " 
 
 'Print the resulting URL to the Debug window 
 Debug.Print vsoHyperlink.CreateURL(False) 
 
End Sub
```


