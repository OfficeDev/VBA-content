---
title: PageSetup.PublicationLayout Property (Publisher)
keywords: vbapb10.chm6946839
f1_keywords:
- vbapb10.chm6946839
ms.prod: publisher
ms.assetid: 6c476789-577d-2088-37dc-bcaed25cd219
ms.date: 06/08/2017
---


# PageSetup.PublicationLayout Property (Publisher)

Returns or sets a  [PbPublicationLayout Enumeration (Publisher)](pbpublicationlayout-enumeration-publisher.md) constant that indicates the layout of a publication. Read/write.


## Syntax

 _expression_. **PublicationLayout**

 _expression_A variable that represents a  **PageSetup** object.


## Return value

 **PBPUBLICATIONLAYOUT**


## Remarks

Using the  **PublicationLayout** property to set the layout of a publication is equivalent to setting the layout from the list box in the **Page Setup** dialog box.


## Example

The following example sets the layout of the active publication to  **pbLayoutBusinessCardUS**, which, by default, specifies a page width of 3.5 inches and a page height of 2 inches.


```vb
With ActiveDocument.PageSetup
    .PublicationLayout = pbLayoutBusinessCardUS
End With

```


## See also


#### Concepts


 [PageSetup Object (Publisher)](pagesetup-object-publisher.md)

