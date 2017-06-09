---
title: Hyperlink.NewWindow Property (Visio)
keywords: vis_sdr.chm15013945
f1_keywords:
- vis_sdr.chm15013945
ms.prod: visio
api_name:
- Visio.Hyperlink.NewWindow
ms.assetid: a86cb7c6-c1e5-eb54-09ce-6f111c3a42ce
ms.date: 06/08/2017
---


# Hyperlink.NewWindow Property (Visio)

Determines whether Microsoft Visio opens a window in a new location when it follows a hyperlink to open a Web page or another Visio document. Read/write.


## Syntax

 _expression_ . **NewWindow**

 _expression_ A variable that represents a **Hyperlink** object.


### Return Value

Integer


## Remarks

Setting the  **NewWindow** property of a **Hyperlink** object is equivalent to setting the NewWindow cell in the shape's Hyperlink. _name_ row.

When  **NewWindow** is set to **False** (0) and the hyperlink's target is a Web page or a document that will open in a browser, the browser will be in the same position and of the same size as the Visio window. If **NewWindow** is **True** (non-zero), a browser window will appear at another location (unless the Visio document is maximized).

When the hyperlink's target is a Visio document, the value of  **NewWindow** determines whether the linked document will open in a window on top of the existing document window, or in another location.


## Example

The following example draws a rectangle shape, adds a  **Hyperlink** object to the shape, sets its **Address** and **NewWindow** properties, and then uses the **Follow** method to navigate the hyperlink. To better observe the effect this property has, before running this macro, size and position the Visio document window so that it is not in the fully maximized position.


```vb
 
Public Sub NewWindow_Example() 
 
 Dim vsoHyperlink As Visio.Hyperlink 
 
 Set vsoHyperlink = ActivePage.DrawRectangle(0,0,5,5).AddHyperlink 
 
 vsoHyperlink.Address = "http://www.microsoft.com/" 
 vsoHyperlink.NewWindow = True 
 vsoHyperlink.Follow 
 
End Sub
```


