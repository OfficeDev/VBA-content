---
title: WebNavigationBarSet.Links Property (Publisher)
keywords: vbapb10.chm8519697
f1_keywords:
- vbapb10.chm8519697
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSet.Links
ms.assetid: 9f155781-390b-ad77-8db7-5099be1409ce
ms.date: 06/08/2017
---


# WebNavigationBarSet.Links Property (Publisher)

Returns a  **WebNavigationBarHyperlinks** collection containing all of the hyperlinks in the specified Web navigation bar set. Read/write.


## Syntax

 _expression_. **Links**

 _expression_A variable that represents a  **WebNavigationBarSet** object.


### Return Value

WebNavigationBarHyperlinks


## Example

Use the  **Links** property to return a **WebNavigationBarHyperlinks** property. This example returns the Web navigation bar hyperlinks of the first Web navigation bar set of the active document.


```vb
ActiveDocument.WebNavigationBarSets(1).Links
```

The following example adds a new Web navigation bar set to the active document, adds a hyperlink to the navigation bar, and then adds the navigation bar to every page of the publication that has the  **AddHyperlinkToWebNavbar** property set to **True** or the **Page.WebPageOptions.IncludePageOnNewWebNavigationBars** property set to **True**.




```vb
With ActiveDocument.WebNavigationBarSets.AddSet(Name:="WebNavigationBarSet1") 
 With .Links 
 .Add Address:="www.microsoft.com", TextToDisplay:="Microsoft", Index:=1 
 End With 
 .AddToEveryPage Left:=10, Top:=10 
End With
```


