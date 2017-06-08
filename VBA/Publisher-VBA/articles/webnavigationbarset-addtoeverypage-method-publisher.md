---
title: WebNavigationBarSet.AddToEveryPage Method (Publisher)
keywords: vbapb10.chm8519698
f1_keywords:
- vbapb10.chm8519698
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSet.AddToEveryPage
ms.assetid: d36a3281-a313-084c-0ae9-7a981a7d9713
ms.date: 06/08/2017
---


# WebNavigationBarSet.AddToEveryPage Method (Publisher)

Adds a  **ShapeRange** of type **pbWebNavigationBar** to each page of the current document.


## Syntax

 _expression_. **AddToEveryPage**( **_Left_**,  **_Top_**,  **_Width_**)

 _expression_A variable that represents a  **WebNavigationBarSet** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Left|Required| **Variant**|The position of the left edge of the shape representing the Web navigation bar set.|
|Top|Required| **Variant**|The position of the top edge of the shape representing the Web navigation bar set.|
|Width|Optional| **Variant**|The width of the shape representing the Web navigation bar set.|

### Return Value

ShapeRange


## Remarks

The specified Web navigation bar set must exist before calling this method. 


## Example

The following example adds a Web navigation bar set named "WebNavBarSet1" to the top of every page in the active document.


```vb
ActiveDocument.WebNavigationBarSets("WebNavBarSet1") _ 
 .AddToEveryPage Left:=10, Top:=20 

```

The following example adds a new Web navigation bar set to the active document and adds it to every page of the publication.




```vb
Dim objWebNavBarSet As WebNavigationBarSet 
 
Set objWebNavBarSet = ActiveDocument.WebNavigationBarSets.AddSet( _ 
 Name:="WebNavBarSet1", _ 
 Design:=pbnbDesignTopLine, _ 
 AutoUpdate:=True) 
 
objWebNavBarSet.AddToEveryPage Left:=50, Top:=10, Width:=500
```


