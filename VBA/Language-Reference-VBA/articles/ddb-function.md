---
title: DDB Function
keywords: vblr6.chm1009279
f1_keywords:
- vblr6.chm1009279
ms.prod: office
ms.assetid: e6ae2093-222c-01cd-86bc-73a3cb79d075
ms.date: 06/08/2017
---


# DDB Function



Returns a [Double](vbe-glossary.md) specifying the depreciation of an asset for a specific time period using the double-declining balance method or some other method you specify.
 **Syntax**
 **DDB( _cost_,** **_salvage_**, **_life_**, **_period_** [, **_factor_** ] **)**
The  **DDB** function has these[named arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
|**_cost_**|Required.  **Double** specifying initial cost of the asset.|
|**_salvage_**|Required.  **Double** specifying value of the asset at the end of its useful life.|
|**_life_**|Required.  **Double** specifying length of useful life of the asset.|
|**_period_**|Required.  **Double** specifying period for which asset depreciation is calculated.|
|**_factor_**|Optional. [Variant](vbe-glossary.md) specifying rate at which the balance declines. If omitted, 2 (double-declining method) is assumed.|
 **Remarks**
The double-declining balance method computes depreciation at an accelerated rate. Depreciation is highest in the first period and decreases in successive periods.
The  **_life_** and **_period_**[arguments](vbe-glossary.md) must be expressed in the same units. For example, if **_life_** is given in months, **_period_** must also be given in months. All arguments must be positive numbers.
The  **DDB** function uses the following formula to calculate depreciation for a given period:
Depreciation /  **_period_** = (( **_cost_** - **_salvage_** ) * **_factor_** ) / **_life_**

## Example

This example uses the  **DDB** function to return the depreciation of an asset for a specified period given the initial cost ( `InitCost`), the salvage value at the end of the asset's useful life ( `SalvageVal`), the total life of the asset in years ( `LifeTime`), and the period in years for which the depreciation is calculated ( `Depr`).


```vb
Dim Fmt, InitCost, SalvageVal, MonthLife, LifeTime, DepYear, Depr
Const YRMOS = 12    ' Number of months in a year.
Fmt = "###,##0.00"
InitCost = InputBox("What's the initial cost of the asset?")
SalvageVal = InputBox("Enter the asset's value at end of its life.")
MonthLife = InputBox("What's the asset's useful life in months?")
Do While MonthLife < YRMOS    ' Ensure period is >= 1 year.
    MsgBox "Asset life must be a year or more."
    MonthLife = InputBox("What's the asset's useful life in months?")
Loop
LifeTime = MonthLife / YRMOS    ' Convert months to years.
If LifeTime <> Int(MonthLife / YRMOS) Then
    LifeTime = Int(LifeTime + 1)    ' Round up to nearest year.
End If 
DepYear = CInt(InputBox("Enter year for depreciation calculation."))
Do While DepYear < 1 Or DepYear > LifeTime
    MsgBox "You must enter at least 1 but not more than " &; LifeTime
    DepYear = InputBox("Enter year for depreciation calculation.")
Loop
Depr = DDB(InitCost, SalvageVal, LifeTime, DepYear)
MsgBox "The depreciation for year " &; DepYear &; " is " &; _
Format(Depr, Fmt) &; "."
```


