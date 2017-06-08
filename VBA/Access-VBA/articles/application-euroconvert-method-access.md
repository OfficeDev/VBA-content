---
title: Application.EuroConvert Method (Access)
keywords: vbaac10.chm12591
f1_keywords:
- vbaac10.chm12591
ms.prod: access
api_name:
- Access.Application.EuroConvert
ms.assetid: 35893059-c6cd-d359-f618-94701a50a049
ms.date: 06/08/2017
---


# Application.EuroConvert Method (Access)

You can use the  **EuroConvert** function to convert a number to euro or from euro to a participating currency. You can also use it to convert a number from one participating currency to another by using the euro as an intermediary (triangulation). The **EuroConvert** function uses fixed conversion rates established by the European Union.


## Syntax

 _expression_. **EuroConvert**( ** _Number_**, ** _SourceCurrency_**, ** _TargetCurrency_**, ** _FullPrecision_**, ** _TriangulationPrecision_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Number_|Required|**Double**|The number you want to convert, or a reference to a field containing the number.|
| _SourceCurrency_|Required|**String**|A string expression, or reference to a field containing the string, corresponding to the International Standards Organization (ISO) acronym for the currency you want to convert. Can be one of the ISO codes listed in the Remarks section.|
| _TargetCurrency_|Required|**String**|A string expression, or reference to a field containing the string, corresponding to the ISO code of the currency to which you want to convert the number. For a list of ISO codes, see the Remarks section.|
| _FullPrecision_|Optional|**Variant**|A  **Boolean** value where **True** (1) ignores the currency-specific rounding rules (called display precision in _sourcecurrency_ argument description) and uses the 6-significant-digit conversion factor with no follow-up rounding. **False** (0) uses the currency-specific rounding rules to display the result. If the parameter is omitted, the default value is **False**.|
| _TriangulationPrecision_|Optional|**Variant**|An  **Integer** value greater than or equal to 3 that specifies the number of significant digits in the calculation precision used for the intermediate euro value when converting between two national currencies.|

### Return Value

Double


## Remarks

The following table contains the ISO codes that can be used with the  _SourceCurrency_ and _TargetCurrency_ arguments.



|**Currency**|**ISO Code**|**Calculation Precision**|**Display Precision**|
|:-----|:-----|:-----|:-----|
|Belgian franc|BEF|0|0|
|Luxembourg franc|LUF|0|0|
|Deutsche mark|DEM|2|2|
|Spanish peseta|ESP|0|0|
|French franc|FRF|2|2|
|Irish punt|IEP|2|2|
|Italian lira|ITL|0|0|
|Netherlands guilder|NLG|2|2|
|Austrian schilling|ATS|2|2|
|Portuguese escudo|PTE|0|0|
|Finnish Markka|FIM|2|2|
|euro|EUR|2|2|
In the preceding table, the calculation precision determines what currency unit to round the result to based on the conversion currency. For example, when converting to Deutsche marks, the calculation precision is 2, and the result is rounded to the nearest pfennig, 100 pfennigs to a mark. The display precision determines how many decimal places appear in the field containing the result.

Later versions of the  **EuroConvert** function may support additional currencies. For information about new participating currencies and updates to the **EuroConvert** function, see the Microsoft Office Euro Currency Web site.



|**Currency**|**ISO Code**|
|:-----|:-----|
|Danish Krone|DKK|
|Drachma|GRD|
|Swedish Krona|SEK|
|Pound Sterling|GBP|
Any trailing zeros are truncated and invalid parameters return #Error.

If the source ISO code is the same as the target ISO code, the original value of the number is active.

This function does not apply a format.

The  **EuroConvert** function uses the current rates established by the European Union. If the rates change, Microsoft will update the function. To get full information about the rules and the rates currently in effect, see the European Commission publications about the euro. For information about obtaining these publications, see the Microsoft Office Euro Currency Web site.


## Example

The first example converts 1.20 Deutsche marks to a euro dollar value (answer = 0.61). The second example converts 1.47 French francs to Deutsche marks (answer = 0.44 DM). They assume conversion rates of 1 euro = 6.55858 French francs and 1.92974 Deutsche marks.


```vb
EuroConvert(1.20,"DEM","EUR") 
EuroConvert(1.47,"FRF","DEM",TRUE,3)
```


## See also


#### Concepts


[Application Object](application-object-access.md)

