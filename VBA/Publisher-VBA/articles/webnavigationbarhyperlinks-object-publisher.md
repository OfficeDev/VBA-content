---
title: WebNavigationBarHyperlinks Object (Publisher)
keywords: vbapb10.chm540671
f1_keywords:
- vbapb10.chm540671
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarHyperlinks
ms.assetid: 4dfa7273-4770-d77c-275c-6b7eeae04aa5
ms.date: 06/08/2017
---


# WebNavigationBarHyperlinks Object (Publisher)

The  **WebNavigationBarHyperlinks** represents a collection of all the **Hyperlink** objects of the specified **WebNavigationBarSet** object.
 


## Example

Use the  **Links** property of the **WebNavigationBarSets** collection to return a **WebNavigationBarHyperlinks** object. The following example adds a hyperlink to the first **WebNavigationBarSet** of the active document.
 

 

```
Dim objWebNavLinks As WebNavigationBarHyperlinks 
Set objWebNavLinks = ActiveDocument.WebNavigationBarSets(1).Links 
objWebNavLinks.Add Address:="www.microsoft.com", _ 
 TextToDisplay:="Microsoft"
```

Use  **WebNavigationBarHyperlinks** **.Count** to return a Long representing the number of hyperlinks in the **WebNavigationBarHyperlinks** collection of the specified **WebNavigationBarSet** object. The following example displays the number of hyperlinks in the first **WebNavigationBarSet** of the active document.
 

 



```
MsgBox ActiveDocument.WebNavigationBarSets(1).Links.Count
```

Use  **WebNavigationBarHyperlinks**.Item(index), where index is the index number, to return a specific **Hyperlink** object from the collection. This example displays the displayed text of the first item in the **WebNavigationBarHyperlinks** collection of the first **WebNavigationBarSet** of the active document.
 

 



```
MsgBox ActiveDocument.WebNavigationBarSets(1).Links.Item(1).TextToDisplay
```


## Methods



|**Name**|
|:-----|
|[Add](webnavigationbarhyperlinks-add-method-publisher.md)|
|[Item](webnavigationbarhyperlinks-item-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](webnavigationbarhyperlinks-application-property-publisher.md)|
|[Count](webnavigationbarhyperlinks-count-property-publisher.md)|
|[Parent](webnavigationbarhyperlinks-parent-property-publisher.md)|

