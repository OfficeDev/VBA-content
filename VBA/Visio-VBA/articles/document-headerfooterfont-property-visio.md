---
title: Document.HeaderFooterFont Property (Visio)
keywords: vis_sdr.chm10550640
f1_keywords:
- vis_sdr.chm10550640
ms.prod: visio
api_name:
- Visio.Document.HeaderFooterFont
ms.assetid: cd4b1f35-c3a2-d48c-fc0d-37f9626ecdab
ms.date: 06/08/2017
---


# Document.HeaderFooterFont Property (Visio)

Specifies the font used for the header and footer text. Read/write.


## Syntax

 _expression_ . **HeaderFooterFont**

 _expression_ A variable that represents a **Document** object.


### Return Value

IFontDisp


## Remarks

COM provides a standard implementation of a font object with the  **IFontDisp** interface on top of the underlying system font support. The **IFontDisp** interface exposes a font object's properties and is implemented in the stdole type library as a **StdFont** object that can be created in Microsoft Visual Basic. The stdole type library is automatically referenced from all Visual Basic for Applications (VBA) projects in Microsoft Visio.

To get information about the  **StdFont** object that supports the **IFontDisp** interface:


1. In the  **Code** group on the[Developer](http://msdn.microsoft.com/library/1bdc55f5-8fc7-7257-03d5-c049eceb29ff%28Office.15%29.aspx) tab, click **Visual Basic**.
    
2. On the  **View** menu, click **Object Browser**.
    
3. In the  **Project/Library** list, click **stdole**.
    
4. Under  **Classes**, examine the class named  **StdFont** .
    
For details about the  **IFontDisp** interface, see the Microsoft Platform SDK on MSDN, the Microsoft Developer Network.

Setting the  **HeaderFooterFont** property is the equivalent of setting values in the **Font** box in the **Choose Font** dialog box (click the **File** tab, click **Print**, click  **Print Preview**, in the  **Preview** group, click **Header &; Footer**, and then click  **Choose Font**).


## Example

The following sample code shows how to use the  **HeaderFooterFont** property to get a reference to the current **Font** object and set the document's text font to non-bold Arial.


```vb
 
Public Sub HeaderFooterFont_Example()  
 
    Dim objStdFont As StdFont 
 
    Set objStdFont = ThisDocument.HeaderFooterFont  
    objStdFont.Name = "Arial"  
    objStdFont.Bold = False 
    Set ThisDocument.HeaderFooterFont = objStdFont  
 
End Sub
```


