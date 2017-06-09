---
title: WorksheetFunction.Convert Method (Excel)
keywords: vbaxl10.chm137344
f1_keywords:
- vbaxl10.chm137344
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Convert
ms.assetid: 3fb95208-6419-da1c-008d-dc00e836183e
ms.date: 06/08/2017
---


# WorksheetFunction.Convert Method (Excel)

Converts a number from one measurement system to another. For example, Convert can translate a table of distances in miles to a table of distances in kilometers.


## Syntax

 _expression_ . **Convert**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|The value in from_units to convert.|
| _Arg2_|Required| **Variant**|The units for number.|
| _Arg3_|Required| **Variant**|The units for the result. Convert accepts the following text values (in quotation marks) for from_unit and to_unit which are listed in the Remarks section below.|

### Return Value

Double


## Remarks



|**Weight and mass**|**From_unit or to_unit**|
|:-----|:-----|
|Gram|"g"|
|Slug|"sg"|
|Pound mass (avoirdupois)|"lbm"|
|U (atomic mass unit)|"u"|
|Ounce mass (avoirdupois)|"ozm"|


|**Distance**|**From_unit or to_unit**|
|:-----|:-----|
|Meter|"m"|
|Statute mile|"mi"|
|Nautical mile|"Nmi"|
|Inch|"in"|
|Foot|"ft"|
|Yard|"yd"|
|Angstrom|"ang"|
|Pica (1/72 in.)|"Pica"|


|**Time**|**From_unit or to_unit**|
|:-----|:-----|
|Year|"yr"|
|Day|"day"|
|Hour|"hr"|
|Minute|"mn"|
|Second|"sec"|


|**Pressure**|**From_unit or to_unit**|
|:-----|:-----|
|Pascal|"Pa" (or "p")|
|Atmosphere|"atm" (or "at")|
|mm of Mercury|"mmHg"|


|**Force**|**From_unit or to_unit**|
|:-----|:-----|
|Newton|"N"|
|Dyne|"dyn" (or "dy")|
|Pound force|"lbf"|


|**Energy**|**From_unit or to_unit**|
|:-----|:-----|
|Joule|"J"|
|Erg|"e"|
|Thermodynamic calorie|"c"|
|IT calorie|"cal"|
|Electron volt|"eV" (or "ev")|
|Horsepower-hour|"HPh" (or "hh")|
|Watt-hour|"Wh" (or "wh")|
|Foot-pound|"flb"|
|BTU|"BTU" (or "btu")|


|**Power**|**From_unit or to_unit**|
|:-----|:-----|
|Horsepower|"HP" (or "h")|
|Watt|"W" (or "w")|


|**Magnetism**|**From_unit or to_unit**|
|:-----|:-----|
|Tesla|"T"|
|Gauss|"ga"|


|**Temperature**|**From_unit or to_unit**|
|:-----|:-----|
|Degree Celsius|"C" (or "cel")|
|Degree Fahrenheit|"F" (or "fah")|
|Kelvin|"K" (or "kel")|


|**Liquid measure**|**From_unit or to_unit**|
|:-----|:-----|
|Teaspoon|"tsp"|
|Tablespoon|"tbs"|
|Fluid ounce|"oz"|
|Cup|"cup"|
|U.S. pint|"pt" (or "us_pt")|
|U.K. pint|"uk_pt"|
|Quart|"qt"|
|Gallon|"gal"|
|Liter|"l" (or "lt")|
The following abbreviated unit prefixes can be prepended to any metric from_unit or to_unit. 



|**Prefix**|**Multiplier**|**Abbreviation**|
|:-----|:-----|:-----|
|exa|1E+18|"E"|
|peta|1E+15|"P"|
|tera|1E+12|"T"|
|giga|1E+09|"G"|
|mega|1E+06|"M"|
|kilo|1E+03|"k"|
|hecto|1E+02|"h"|
|dekao|1E+01|"e"|
|deci|1E-01|"d"|
|centi|1E-02|"c"|
|milli|1E-03|"m"|
|micro|1E-06|"u"|
|nano|1E-09|"n"|
|pico|1E-12|"p"|
|femto|1E-15|"f"|
|atto|1E-18|"a"|

- If the input data types are incorrect, Convert generates an error.
    
- If the unit does not exist, Convert generates an error.
    
- If the unit does not support an abbreviated unit prefix, Convert generates an error.
    
- If the units are in different groups, Convert generates an error.
    
- Unit names and prefixes are case-sensitive.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

