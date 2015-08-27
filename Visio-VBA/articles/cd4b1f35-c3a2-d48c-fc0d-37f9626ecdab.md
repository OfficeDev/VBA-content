
# Document.HeaderFooterFont Property (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


Specifies the font used for the header and footer text. Read/write.

## Syntax
<a name="sectionSection0"> </a>

 _expression_. **HeaderFooterFont**

 _expression_A variable that represents a  **Document** object.


### Return Value

IFontDisp


## Remarks
<a name="sectionSection1"> </a>

COM provides a standard implementation of a font object with the  **IFontDisp** interface on top of the underlying system font support. The **IFontDisp** interface exposes a font object's properties and is implemented in the stdole type library as a **StdFont** object that can be created in Microsoft Visual Basic. The stdole type library is automatically referenced from all Visual Basic for Applications (VBA) projects in Microsoft Visio.

To get information about the  **StdFont** object that supports the **IFontDisp** interface:


1. In the  **Code** group on the [Developer](1bdc55f5-8fc7-7257-03d5-c049eceb29ff.md) tab, click **Visual Basic**.
    
2. On the  **View** menu, click **Object Browser**.
    
3. In the  **Project/Library** list, click **stdole**.
    
4. Under  **Classes**, examine the class named  **StdFont**.
    
For details about the  **IFontDisp** interface, see the Microsoft Platform SDK on MSDN, the Microsoft Developer Network.

Setting the  **HeaderFooterFont** property is the equivalent of setting values in the **Font** box in the **Choose Font** dialog box (click the **File** tab, click **Print**, click  **Print Preview**, in the  **Preview** group, click **Header &amp; Footer**, and then click  **Choose Font**).


## Example
<a name="sectionSection2"> </a>

The following sample code shows how to use the  **HeaderFooterFont** property to get a reference to the current **Font** object and set the document's text font to non-bold Arial.


```
 
Public Sub HeaderFooterFont_Example()  
 
    Dim objStdFont As StdFont 
 
    Set objStdFont = ThisDocument.HeaderFooterFont  
    objStdFont.Name = "Arial"  
    objStdFont.Bold = False 
    Set ThisDocument.HeaderFooterFont = objStdFont  
 
End Sub 

```

