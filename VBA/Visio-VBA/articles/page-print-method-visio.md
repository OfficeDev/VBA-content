---
title: Page.Print Method (Visio)
keywords: vis_sdr.chm10916445
f1_keywords:
- vis_sdr.chm10916445
ms.prod: visio
api_name:
- Visio.Page.Print
ms.assetid: 021cdd78-1699-4345-5b32-c2c0a300ca00
ms.date: 06/08/2017
---


# Page.Print Method (Visio)

Prints the contents of an object to the default printer.


## Syntax

 _expression_ . **Print**

 _expression_ A variable that represents a **Page** object.


### Return Value

Nothing


## Remarks

For a  **Document** object, this method prints all of the document's pages. Background pages are printed on the same sheet of paper as the foreground pages to which they are assigned.

For a  **Page** object, this method prints the page and its background page (if any) on the same sheet of paper.

If you're using Microsoft Visual Basic for Applications (VBA) or Visual Basic, you must assign the method result to a dummy variable and you must apply the method to a variable of type  **Object** , not type **Visio.Document** or **Visio.Page** . For example, to print a document, use the following code.




```vb
 
 Dim vsoDocument As Visio.Document 
 Dim vsoDocumentTemp as Object 
 Dim strDummy As String 
 
 Set vsoDocument = ThisDocument 
 Set vsoDocumentTemp = vsoDocument 
 strDummy = vsoDocumentTemp.Print 

```


