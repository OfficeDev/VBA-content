---
title: System.CountryRegion Property (Word)
keywords: vbawd10.chm154468455
f1_keywords:
- vbawd10.chm154468455
ms.prod: word
api_name:
- Word.System.CountryRegion
ms.assetid: 51db26e6-9f24-5934-24a4-0ed87bb51f69
ms.date: 06/08/2017
---


# System.CountryRegion Property (Word)

Returns the country/region designation of the system. Read-only  **WdCountry** .


## Syntax

 _expression_ . **CountryRegion**

 _expression_ Required. A variable that represents a **[System](system-object-word.md)** object.


## Example

If the  **CountryRegion** property returns **wdUS** , this example converts the top margin value from points to inches.


```vb
Dim sngMargin As Single 
 
If System.CountryRegion = wdUS Then 
 sngMargin = ActiveDocument.PageSetup.TopMargin 
 MsgBox "Top margin is " &; PointsToInches(sngMargin) 
End If
```


## See also


#### Concepts


[System Object](system-object-word.md)

