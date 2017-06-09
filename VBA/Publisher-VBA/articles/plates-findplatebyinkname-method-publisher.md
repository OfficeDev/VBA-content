---
title: Plates.FindPlateByInkName Method (Publisher)
keywords: vbapb10.chm2818053
f1_keywords:
- vbapb10.chm2818053
ms.prod: publisher
api_name:
- Publisher.Plates.FindPlateByInkName
ms.assetid: 4ebbc826-468b-7cd7-806e-056e4cbb488c
ms.date: 06/08/2017
---


# Plates.FindPlateByInkName Method (Publisher)

Returns a  **Plate** object that represents the plate of the specified ink name.


## Syntax

 _expression_. **FindPlateByInkName**( **_InkName_**)

 _expression_An expression that returns a  **Plates** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|InkName|Required| **PbInkName**|Specifies the plate to return.|

### Return Value

Plate


## Remarks

The InkName parameter can be one of the  ** [PbInkName](http://msdn.microsoft.com/library/69e335b8-40b8-c984-84b6-64073a8ed7ab%28Office.15%29.aspx)** constants declared in the Microsoft Publisher type library.

Process colors are assigned different index numbers in the  **Plates** collection than in the **PrintablePlates** collection. Use the **FindPlateByInkName** method to insure that the desired **Plate** or **PrintablePlate** object is accessed.


## Example

The following example returns properties for the plate representing the third spot color defined for the active publication.


```vb
Sub ListPlatePropertiesByInkName() 
Dim pplPlate As Plate 
 
 Set pplPlate = ActiveDocument.Plates.FindPlateByInkName(pbInkNameSpot3) 
 
 With pplPlate 
 Debug.Print "Plate Name: " &; .Name 
 Debug.Print "Index: " &; .Index 
 Debug.Print "Ink Name: " &; .InkName 
 Debug.Print "Color: " &; .Color 
 Debug.Print "Luminance: " &; .Luminance 
 Debug.Print "In Use?: " &; .InUse 
 End With 
End Sub
```


