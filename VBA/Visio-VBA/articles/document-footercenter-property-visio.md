---
title: Document.FooterCenter Property (Visio)
keywords: vis_sdr.chm10550580
f1_keywords:
- vis_sdr.chm10550580
ms.prod: visio
api_name:
- Visio.Document.FooterCenter
ms.assetid: 7abdcd6c-39ed-ad05-bfef-cbd979f3a8d6
ms.date: 06/08/2017
---


# Document.FooterCenter Property (Visio)

Gets or sets the text string that appears in the center portion of a document's footer. Read/write.


## Syntax

 _expression_ . **FooterCenter**

 _expression_ A variable that represents a **Document** object.


### Return Value

String


## Remarks

You can also set this value in the  **Center** box under **Footer** in the **Header and Footer** dialog box (click the **File** tab, click **Print**, click  **Print Preview**, and then in the  **Preview** group, click **Header &; Footer**).

Both the string returned by the property and the string you pass to the property can contain escape codes that represent data. These escape codes can be concatenated with other text. For a list of valid escape codes you can use with the  **FooterCenter** property, see the **[FooterLeft](document-footerleft-property-visio.md)** property topic.


## Example

The following macro shows how to place a string containing the current page number and total number of pages into the center portion of a document's footer. After you run this macro on a one-page document, the center portion of the footer contains "Page 1 of 1".


```vb
 
Sub FooterCenter_Example()  
 
    Dim strFooter as String 
 
    'Build the footer string.  
    strFooter = "Page &;p of &;P"  
 
    'Set the footer of the current document.  
     ThisDocument.FooterCenter = strFooter 
  
End Sub
```


