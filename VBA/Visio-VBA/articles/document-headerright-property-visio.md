---
title: Document.HeaderRight Property (Visio)
keywords: vis_sdr.chm10550655
f1_keywords:
- vis_sdr.chm10550655
ms.prod: visio
api_name:
- Visio.Document.HeaderRight
ms.assetid: 3d702cb7-9b70-5f00-c2ea-b619cbfed37f
ms.date: 06/08/2017
---


# Document.HeaderRight Property (Visio)

Gets or sets the text string that appears in the right portion of a document's header. Read/write.


## Syntax

 _expression_ . **HeaderRight**

 _expression_ A variable that represents a **Document** object.


### Return Value

String


## Remarks

You can also set this value in the  **Right** box under **Header** in the **Header and Footer** dialog box (click the **File** tab, click **Print**, click  **Print Preview**, and then in the  **Preview** group, click **Header &; Footer**).

Both the string that  **HeaderRight** returns and the string to which you set it can contain escape codes that represent data. These escape codes can be concatenated with other text. For a list of valid escape codes you can use with the **HeaderRight** property, see the **[FooterLeft](document-footerleft-property-visio.md)** property.


## Example

The following macro shows how to place a string containing the current date into the right portion of a document's header. After you run this macro, if the date is October 1, 2009, the right portion of the header contains "The date is Thursday, October 1, 2009".


```vb
 
Sub HeaderRight_Example() 
  
    Dim strHeader as String 
 
    'Build header string.  
    strHeader = "The date is " &; "&;D"  
 
    'Set header of the current document.  
    ThisDocument.HeaderRight = strHeader  
 
End Sub
```


