---
title: Plate.Delete Method (Publisher)
keywords: vbapb10.chm2883600
f1_keywords:
- vbapb10.chm2883600
ms.prod: publisher
api_name:
- Publisher.Plate.Delete
ms.assetid: fadaba7c-6636-f1e2-e360-3fcf8700ab36
ms.date: 06/08/2017
---


# Plate.Delete Method (Publisher)

Deletes the specified plate.


## Syntax

 _expression_. **Delete**( **_PlateReplaceWith_**,  **_ReplaceTint_**)

 _expression_A variable that represents a  **Plate** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|PlateReplaceWith|Optional| **Variant**| **Plate**. The plate with which to replace the deleted plate.|
|ReplaceTint|Optional| **PbReplaceTint**|How to replace tints.|

## Remarks

Returns "Permission Denied" if you attempt to delete the last plate in the  **Plates** collection.

The ReplaceTint parameter can be one of the following  **pbReplaceTint** constants.



| **pbReplaceTintKeepTints**|Maintain the same tint percentage in the ink represented by the replacement plate as in the deleted plate. For example, replace a 100% tint of yellow with a 100% tint of blue.|
| **pbReplaceTintMaintainLuminosity**| Maintain the same lightness value in the ink represented by the replacement plate as in the deleted plate. For example, replace a 100% tint of yellow with an approximately 10% tint of blue.|
| **pbReplaceTintUseDefault**|Use the default. |
If the  **pbReplaceTintMaintainLuminosity** constant is specified, the percentage of replacment ink in each color is calculated based on the luminosity values of the inks represented by the deleted and replacement plates. Publisher performs the following calculation, where _L1_ is the deleted ink luminosity, and _L2_ is the replacement ink luminosity: (100- _L1_)/(100- _L2_).

For example, red ink has a luminosity of 30, and black has a luminosity of 0. Suppose you replaced the red ink plate in a publication with a black ink plate. If  **pbReplaceTintKeepTints** is specified, Publisher performs the following calculation to determine the percentage of black ink for each red color: (100-30)/(100-0). A color that was 100% red would now be 70% black; a color that was 50% red would now be 35% black, and so on.

If the  **pbReplaceTintKeepTints** constant is specified, the percentage of the replacement ink in each color is the same as the deleted color. For example, if red ink is replaced with black ink, 100% tint of red is replaced by 100% tint of black, 50% red with 50% black, and so on.

You cannot specify the  **pbReplaceTintMaintainLuminosity** or **pbReplaceTintUseDefault** constants if the replacement plate represents an ink that has a higher luminosity (that is, is lighter) than the deleted plate. This is because the lighter ink can not be printed at more than 100%, so it will not be able to match the luminosity of the darker ink.


## Example

The following example loops through the active publication's plates collection, determines which plates represent inks not used in the publication, and deletes them. This example assumes that at least one of the plates is in use (the Delete method returns "Permission Denied" if you attempt to delete the last plate in the collection.)


```vb
Sub DeleteUnusedInks() 
 
Dim intCount As Integer 
 
With ActiveDocument.Plates 
 For intCount = .Count To 1 Step -1 
 With .Item(intCount) 
 If .InUse = False Then 
 Debug.Print "Name: " &; .Name 
 .Delete 
 End If 
 End With 
 Next 
End With 
 
End Sub
```


