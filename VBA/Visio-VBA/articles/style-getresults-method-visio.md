---
title: Style.GetResults Method (Visio)
keywords: vis_sdr.chm11416320
f1_keywords:
- vis_sdr.chm11416320
ms.prod: visio
api_name:
- Visio.Style.GetResults
ms.assetid: 43106f2c-6731-b110-f713-7d172909feae
ms.date: 06/08/2017
---


# Style.GetResults Method (Visio)

Gets the results or formulas of many cells.


## Syntax

 _expression_ . **GetResults**( **_SRCStream()_** , **_Flags_** , **_UnitsNamesOrCodes()_** , **_resultArray()_** )

 _expression_ A variable that represents a **Style** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SID_SRCStream()_|Required| **Integer**|An array identifying cells to be queried.|
| _Flags_|Required| **Integer**|Flags that influence the type of entries returned in results.|
| _UnitsNamesOrCodes()_|Required| **Variant**| An array of measurement units that results are to be returned in.|
| _resultArray()_|Required| **Variant**|Out parameter. An array that receives results or formulas of queried cells.|

### Return Value

Nothing


## Remarks

The  **GetResults** method is like the **Result** property for the **Cell** object, except that it can be used to get the results (values) of many cells at once, rather than one cell at a time.

For  **Style** objects, you can use the **GetResults** method to get results of any set of cells.

 _SID_SRCStream()_ is an array of 2-byte integers. For **Style** objects, _SID_SRCStream()_ should be a one-dimensional array of 3 _n_ 2-byte integers for _n_ >= 1. The **GetResults** method interprets _SID_SRCStream()_ as:




```
{sectionIdx, rowIdx, cellIdx }n
```

where  _sectionIdx_ is the section index of the desired cell, _rowIdx_ is its row index and _cellIdx_ is its cell index.

The  _Flags_ argument indicates what data type the returned results should be expressed in. Its value should be one of the following.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visGetFloats**|0|Results returned as doubles (VT_R8).|
| **visGetTruncatedInts**|1|Results returned as truncated long integers (VT_I4).|
| **visGetRoundedInts**|2|Results returned as rounded long integers (VT_I4).|
| **visGetStrings**|3|Results returned as strings (VT_BSTR).|
| **visGetFormulas**|4|Formulas returned as strings (VT_BSTR).|
| **visGetFormulasU**|5|Formulas returned in universal syntax (VT_BSTR).|
The  _UnitsNamesOrCodes()_ argument is an array that controls what measurement units individual results are returned in. Each entry in the array can be a string such as "inches", "inch", "in.", or "i". Strings may be used for all supported Visio units such as centimeters, meters, miles, and so on. You can also indicate desired units with integer constants ( **visCentimeters** , **visInches** , and so on) declared by the Visio type library. Note that the values specified in the _UnitsNamesOrCodes()_ array have no effect if _Flags_ is **visGetFormulas** .

If  _UnitsNamesOrCodes()_ is not null, the application expects it to be a one-dimensional array of 1 <= _u_**Variants** . Each entry can be a string or integer code, or empty (nothing). If the _i_ 'th entry is empty, the _i_ 'th returned result is returned in the units designated by _UnitsNamesOrCodes(j)_, where  _j_ is the index of the most recent prior non-empty entry. Thus if you want all returned values to be in the same units, you need only pass a _UnitsNamesOrCodes()_ array with one entry. If there is no prior non-empty entry, or if no _UnitsNameOrCodes()_ array is supplied, **visNumber** (0x20) is used. This causes internal units (like the **ResultIU** property of a **Cell** object) to be returned.

If the  **GetResults** method succeeds, results returns a one-dimensional array of _n_ variants indexed from zero (0) to _n_ - 1. The type of the returned variants is a function of _Flags_. The  _resultArray()_ parameter is an out parameter that is allocated by the **GetResults** method, which passes ownership back to the caller. The caller should eventually perform **SafeArrayDestroy** on the returned array. Note that **SafeArrayDestroy** has the side effect of clearing the variants referenced by the array's entries, hence deallocating any strings the **GetResults** method returns. (Microsoft Visual Basic and Microsoft Visual Basic for Applications take care of this for you.)


## Example

The following example shows how to use the  **GetResults** method. This example assumes there is an active page that has at least 3 shapes on it. It uses the **GetResults** method to get the width of shape 1, the height of shape 2, and the angle of shape 3.

This example uses the  **GetResults** method of the **Page** object to get 3 cell formulas. The input array has 4 slots for each cell, as it also would for **Master** objects. For **Shape** or **Style** objects, only 3 slots would be needed for each cell (section, row, and cell).




```vb
 
Public Sub GetResults_Example() 
 
 On Error GoTo HandleError 
 
 Dim intCounter As Integer 
 Dim aintSheetSectionRowColumn(1 To (3 * 4)) As Integer 
 
 aintSheetSectionRowColumn(1) = ActivePage.Shapes(1).ID 
 aintSheetSectionRowColumn(2) = visSectionObject 
 aintSheetSectionRowColumn(3) = visRowXFormOut 
 aintSheetSectionRowColumn(4) = visXFormWidth 
 
 aintSheetSectionRowColumn(5) = ActivePage.Shapes(2).ID 
 aintSheetSectionRowColumn(6) = visSectionObject 
 aintSheetSectionRowColumn(7) = visRowXFormOut 
 aintSheetSectionRowColumn(8) = visXFormHeight 
 
 aintSheetSectionRowColumn(9) = ActivePage.Shapes(3).ID 
 aintSheetSectionRowColumn(10) = visSectionObject 
 aintSheetSectionRowColumn(11) = visRowXFormOut 
 aintSheetSectionRowColumn(12) = visXFormAngle 
 
 'Get first two values in inches. The second element in 
 'the units array is left uninitialized (empty) because we 
 'want the second result in the same units as the first 
 'result. The third result is initialized in degrees. Note that 
 'units can be expressed as a string or an integer constant. 
 Dim avarUnits(1 To 3) As Variant 
 avarUnits(1) = "in." 
 avarUnits(3) = visDegrees 
 
 'Return results of cells as an array of floating point numbers. 
 Dim avarResults() As Variant 
 ActivePage.GetResults aintSheetSectionRowColumn, visGetFloats, _ 
 avarUnits, avarResults 
 
 For intCounter = 0 To 3 
 Debug.Print avarResults(intCounter) 
 Next 
 
 Exit Sub 
 
HandleError: 
 MsgBox "Error" 
 Exit Sub 
 
End Sub
```


