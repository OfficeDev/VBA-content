---
title: Document.FooterLeft Property (Visio)
keywords: vis_sdr.chm10550585
f1_keywords:
- vis_sdr.chm10550585
ms.prod: visio
api_name:
- Visio.Document.FooterLeft
ms.assetid: e832c09d-3ddb-4351-43ad-e1c5633b7bc9
ms.date: 06/08/2017
---


# Document.FooterLeft Property (Visio)

Gets or sets the text string that appears on the left side of a document's footer. Read/write.


## Syntax

 _expression_ . **FooterLeft**

 _expression_ A variable that represents a **Document** object.


### Return Value

String


## Remarks

You can also set this value in the  **Left** box under **Footer** in the **Header and Footer** dialog box (click the **File** tab, click **Print**, click  **Print Preview**, and then in the  **Preview** group, click **Header &; Footer**).

Both the string returned by  **FooterLeft** and the string to which you set **FooterLeft** can contain escape codes that represent data. These escape codes can be concatenated with other text.

Following is a list of valid escape codes for document footers and headers.



|** Escape code**|** Description**|
|:-----|:-----|
| &;p| Page number|
| &;t or &;T| Current time|
| &;d (short version) or &;D (long version)| Current date|
| &;&;| Ampersand|
| &;e| File name extension|
| &;f| File name|
| &;f&;e| File name and extension|
| &;n| Page name|
| &;P| Total printed pages|

## Example

The following macro shows how to place a string containing the current date into the left portion of a document's footer. After you run this macro, if the date is May 4, 2007, the left portion of the footer contains "The date is Thursday, May 4, 2007".


```vb
 
Sub FooterLeft_Example()  
 
    Dim strFooter as String 
 
    'Build the footer string.  
    strFooter = "The date is " &; "&;D"  
 
    'Set the footer of the current document.  
    ThisDocument.FooterLeft = strFooter  
 
End Sub
```


