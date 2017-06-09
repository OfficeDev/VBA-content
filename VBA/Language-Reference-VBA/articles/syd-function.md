---
title: SYD Function
keywords: vblr6.chm1009290
f1_keywords:
- vblr6.chm1009290
ms.prod: office
ms.assetid: a5afb589-eaf4-d253-8999-5063bcab680c
ms.date: 06/08/2017
---


# SYD Function



Returns a [Double](vbe-glossary.md) specifying the sum-of-years' digits depreciation of an asset for a specified period.
 **Syntax**
 **SYD( _cost_**, **_salvage_**, **_life_**, **_period_ )**
The  **SYD** function has these[named arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
|**_cost_**|Required.  **Double** specifying initial cost of the asset.|
|**_salvage_**|Required.  **Double** specifying value of the asset at the end of its useful life.|
|**_life_**|Required.  **Double** specifying length of the useful life of the asset.|
|**_period_**|Required.  **Double** specifying period for which asset depreciation is calculated.|
 **Remarks**
The  **_life_** and **_period_**[arguments](vbe-glossary.md) must be expressed in the same units. For example, if **_life_** is given in months, **_period_** must also be given in months. All arguments must be positive numbers.

## Example

This example uses the  **SYD** function to return the depreciation of an asset for a specified period given the asset's initial cost ( `InitCost`), the salvage value at the end of the asset's useful life ( `SalvageVal`), and the total life of the asset in years ( `LifeTime`). The period in years for which the depreciation is calculated is  `PDepr`.


```vb
Dim Fmt, InitCost, SalvageVal, MonthLife, LifeTime, DepYear, PDepr
Const YEARMONTHS = 12    ' Number of months in a year.
Fmt = "###,##0.00"    ' Define money format.
InitCost = InputBox("What's the initial cost of the asset?")
SalvageVal = InputBox("What's the asset's value at the end of its life?")
MonthLife = InputBox("What's the asset's useful life in months?")
Do While MonthLife < YEARMONTHS    ' Ensure period is >= 1 year.
    MsgBox "Asset life must be a year or more."
    MonthLife = InputBox("What's the asset's useful life in months?")
Loop
LifeTime = MonthLife / YEARMONTHS    ' Convert months to years.
If LifeTime <> Int(MonthLife / YEARMONTHS) Then
    LifeTime = Int(LifeTime + 1)    ' Round up to nearest year.
End If 
DepYear = CInt(InputBox("For which year do you want depreciation?"))
Do While DepYear < 1 Or DepYear > LifeTime
    MsgBox "You must enter at least 1 but not more than " &; LifeTime
    DepYear = CInt(InputBox("For what year do you want depreciation?"))
Loop
PDepr = SYD(InitCost, SalvageVal, LifeTime, DepYear)
MsgBox "The depreciation for year " &; DepYear &; " is " &; Format(PDepr, Fmt) &; "."
```


