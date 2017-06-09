---
title: WebNavigationBarSets.AddSet Method (Publisher)
keywords: vbapb10.chm8454148
f1_keywords:
- vbapb10.chm8454148
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarSets.AddSet
ms.assetid: 5b998e14-b1eb-2a4a-2ed5-9a1ef16d69c1
ms.date: 06/08/2017
---


# WebNavigationBarSets.AddSet Method (Publisher)

Adds a new  **WebNavigationBarSet** object representing a Web navigation bar set to the specified **WebNavigationBarSets** collection. .


## Syntax

 _expression_. **AddSet**( **_Name_**,  **_Design_**,  **_AutoUpdate_**)

 _expression_A variable that represents a  **WebNavigationBarSets** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Name|Required| **String**|The name of the Web navigation bar to be added. This parameter must be unique.|
|Design|Optional| **PbWizardNavBarDesign**|Specifies the navigation bar design scheme.|
|AutoUpdate|Optional| **Boolean**| **True** if all pages with the **AddHyperlinkToWebNavBar** property set to **True**are added as links to the navigation bar and the navigation bar is kept updated.|

### Return Value

WebNavigationBarSet


## Remarks

The  **Name** parameter must be unique to avoid a run time error.


## Example

The following example adds a  **WebNavigationBarSet** object to the **WebNavigationBarSets** collection of the active document then sets some properties.


```vb
Dim objWebNavBarSet As WebNavigationBarSet 
 
Set objWebNavBarSet = ActiveDocument.WebNavigationBarSets.AddSet( _ 
 Name:="WebNavBarSet1", _ 
 Design:=pbnbDesignAmbient, _ 
 AutoUpdate:=True) 
 
With objWebNavBarSet 
 .AddToEveryPage Left:=50, Top:=10 
 .ButtonStyle = pbnbDesignTopLine 
 .ChangeOrientation pbNavBarOrientHorizontal 
End With
```


