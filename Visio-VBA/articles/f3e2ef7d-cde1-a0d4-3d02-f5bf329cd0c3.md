
# Cell.ResultInt Property (Visio)

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

Gets the value of a cell expressed as an integer. Read-only.


## Syntax

 _expression_. **ResultInt**( **_UnitsNameOrCode_**,  **_fRound_**)

 _expression_A variable that represents a  **Cell** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|UnitsNameOrCode|Required| **Variant**|The units to use when retrieving the cell's value.|
|fRound|Required| **Integer**|Zero (0) to truncate the value; non-zero to round it.|

### Return Value

Long


## Remarks

Getting the  **ResultInt** property is similar to a getting a cell's **Result** property. The difference is that the **ResultInt** property returns an integer for the value of the cell, whereas the **Result** property returns a floating point number.

You can specify UnitsNameOrCode as an integer or a string value. If the string is invalid, an error is generated. For example, the following statements all setUnitsNameOrCode to inches.

lngRet = **Cell.ResultInt**( **visInches**, fRound)

lngRet = **Cell.ResultInt**(65, fRound)

lngRet = **Cell.ResultInt**("in", fRound) where "in" can also be any of the alternate strings representing inches, such as "inch", "in.", or "intCounter".

For a complete list of valid unit strings along with their corresponding Automation constants (integer values), see  [About Units of Measure](b6140312-b8e6-0cf2-9fe0-b14e800216bf.md).

Automation constants for representing units are declared by the Visio type library in member  ** [VisUnitCodes ](fce91c1b-d5c2-6522-2446-0b8f6cacbc84.md)**.

The following constants for fRound are declared in the Visio type library in member **VisRoundFlags**.



|**Constant **|**Value **|**Description **|
|:-----|:-----|:-----|
| **visTruncate**|0 |Truncate the result. |
| **visRound**|1 |Round the result. |
