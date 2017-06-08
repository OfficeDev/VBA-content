---
title: SLN Function
keywords: vblr6.chm1009289
f1_keywords:
- vblr6.chm1009289
ms.prod: office
ms.assetid: e9587257-b3b8-048f-76ed-609448596a14
ms.date: 06/08/2017
---


# SLN Function



Returns a [Double](vbe-glossary.md) specifying the straight-line depreciation of an asset for a single period.
 **Syntax**
 **SLN( _cost_**, **_salvage_**, **_life_ )**
The  **SLN** function has these[named arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
|**_cost_**|Required.  **Double** specifying initial cost of the asset.|
|**_salvage_**|Required.  **Double** specifying value of the asset at the end of its useful life.|
|**_life_**|Required.  **Double** specifying length of the useful life of the asset.|
 **Remarks**
The depreciation period must be expressed in the same unit as the  **_life_**[argument](vbe-glossary.md). All arguments must be positive numbers.

## Example

This example uses the  **SLN** function to return the straight-line depreciation of an asset for a single period given the asset's initial cost ( `InitCost`), the salvage value at the end of the asset's useful life ( `SalvageVal`), and the total life of the asset in years ( `LifeTime`).


```vb
Dim Fmt, InitCost, SalvageVal, MonthLife, LifeTime, PDepr
Const YEARMONTHS = 12    ' Number of months in a year.
Fmt = "###,##0.00"    ' Define money format.
InitCost = InputBox("What's the initial cost of the asset?")
SalvageVal = InputBox("What's the asset's value at the end of its useful life?")
MonthLife = InputBox("What's the asset's useful life in months?")
Do While MonthLife < YEARMONTHS    ' Ensure period is >= 1 year.
    MsgBox "Asset life must be a year or more."
    MonthLife = InputBox("What's the asset's useful life in months?")
Loop
LifeTime = MonthLife / YEARMONTHS    ' Convert months to years.
If LifeTime <> Int(MonthLife / YEARMONTHS) Then
    LifeTime = Int(LifeTime + 1)    ' Round up to nearest year.
End If
PDepr = SLN(InitCost, SalvageVal, LifeTime)
MsgBox "The depreciation is " &; Format(PDepr, Fmt) &; " per year."

```


