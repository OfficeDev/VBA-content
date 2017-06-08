---
title: Document.HeaderFooterColor Property (Visio)
keywords: vis_sdr.chm10550635
f1_keywords:
- vis_sdr.chm10550635
ms.prod: visio
api_name:
- Visio.Document.HeaderFooterColor
ms.assetid: d1f3887f-d6b5-feb5-b119-ddf3d9eb3542
ms.date: 06/08/2017
---


# Document.HeaderFooterColor Property (Visio)

Specifies the color of the header and footer text. Read/write.


## Syntax

 _expression_ . **HeaderFooterColor**

 _expression_ A variable that represents a **Document** object.


### Return Value

OLE_COLOR


## Remarks

Valid values for  **OLE_COLOR** within Microsoft Visio can be one of the following:


- &;H00 _bbggrr,_ where _bb_ is the blue value between 0 and 0xFF (255), _gg_ the green value, and _rr_ the red value.
    
- &;H800000 _xx_ , where _xx_ is a valid **GetSysColor** index.
    
For details about the  **GetSysColor** function, search for " **GetSysColor** " in the Microsoft Platform SDK on MSDN, the Microsoft Developer Network.

The  **OLE_COLOR** data type is used for properties that return colors. When a property is declared as **OLE_COLOR** , the **Properties** window will display a color-picker dialog box that allows the user to select the color for the property visually, rather than having to remember the numeric equivalent.

You can also set this value in the  **Color** box in the **Choose Font** dialog box (click the **File** tab, click **Print**, click  **Print Preview**, in the  **Preview** group, click **Header &; Footer**, and then click  **Choose Font**).


## Example

The following macro shows how to assign the color blue to text in the header and footer.


```vb
 
Sub HeaderFooterColor_Example() 
  
    'Set color of the header of this document to blue.  
    ThisDocument.HeaderFooterColor = &;H00FF0000  
 
End Sub
```


