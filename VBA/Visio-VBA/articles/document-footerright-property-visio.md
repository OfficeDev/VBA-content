---
title: Document.FooterRight Property (Visio)
keywords: vis_sdr.chm10550595
f1_keywords:
- vis_sdr.chm10550595
ms.prod: visio
api_name:
- Visio.Document.FooterRight
ms.assetid: 17db938c-6b1b-6cd1-7f4e-65aca275f30b
ms.date: 06/08/2017
---


# Document.FooterRight Property (Visio)

Gets or sets the text string that appears in the right portion of a document's footer. Read/write.


## Syntax

 _expression_ . **FooterRight**

 _expression_ A variable that represents a **Document** object.


### Return Value

String


## Remarks

You can also set this value in the  **Right** box under **Footer** in the **Header and Footer** dialog box (click the **File** tab, click **Print**, click  **Print Preview**, and then in the  **Preview** group, click **Header &; Footer**).

Both the string returned by  **FooterRight** and the string to which you set **FooterRight** can contain escape codes that represent data. These escape codes can be concatenated with other text. For a list of valid escape codes you can use with the **FooterRight** property, see the **[FooterLeft](document-footerleft-property-visio.md)** property topic.


## Example

The following macro shows how to place a string containing the current date into the right portion of a document's footer. After you run this macro, if the date is May 4, 2007, the right portion of the footer contains "The date is Thursday, May 4, 2007".


```vb
 
Sub FooterRight_Example()  
 
    Dim strFooter as String 
 
    'Build the footer string.  
    strFooter = "The date is " &; "&;D"  
 
    'Set the footer of the current document.  
    ThisDocument.FooterRight = strFooter  
 
End Sub
```


