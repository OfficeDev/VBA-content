---
title: Document.HeaderCenter Property (Visio)
keywords: vis_sdr.chm10550630
f1_keywords:
- vis_sdr.chm10550630
ms.prod: visio
api_name:
- Visio.Document.HeaderCenter
ms.assetid: 8695883a-8b00-eef4-aecd-81ad47581a82
ms.date: 06/08/2017
---


# Document.HeaderCenter Property (Visio)

Contains the text string that appears in the center portion of a document's header. Read/write.


## Syntax

 _expression_ . **HeaderCenter**

 _expression_ A variable that represents a **Document** object.


### Return Value

String


## Remarks

You can also set this value in the  **Center** box under **Header** in the **Header and Footer** dialog box (click the **File** tab, click **Print**, click  **Print Preview**, and then in the  **Preview** group, click **Header &; Footer**).

Both the string that  **HeaderCenter** returns and the string to which you set it can contain escape codes that represent data. These escape codes can be concatenated with other text. For a list of valid escape codes you can use with the **HeaderCenter** property, see the **[FooterLeft](document-footerleft-property-visio.md)** property.


## Example

The following macro shows how to place the string containing "Document Title" into the center portion of the document's header.


```vb
 
Sub HeaderCenter_Example() 
  
    'Set header of current document.  
    ThisDocument.HeaderCenter = "Document Title"  
 
End Sub
```


