---
title: Style.SetResults Method (Visio)
keywords: vis_sdr.chm11416580
f1_keywords:
- vis_sdr.chm11416580
ms.prod: visio
api_name:
- Visio.Style.SetResults
ms.assetid: f03b627b-7b54-0190-96d5-c95eddf44ceb
ms.date: 06/08/2017
---


# Style.SetResults Method (Visio)

Sets the results or formulas of one or more cells.


## Syntax

 _expression_ . **SetResults**( **_SRCStream()_** , **_UnitsNamesOrCodes()_** , **_resultArray()_** , **_Flags_** )

 _expression_ A variable that represents a **Style** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SID_SRCStream()_|Required| **Integer**|An array identifying cells to be modified.|
| _UnitsNamesOrCodes()_|Required| **Variant**|An array of measurement units to be attributed to entries in the results array.|
| _resultArray()_|Required| **Variant**|Results or formulas to be assigned to identified cells.|
| _Flags_|Required| **Integer**|Flags that influence the behavior of  **SetResults** .|

### Return Value

Integer


## Remarks

The  **SetResults** method is like the **Result** method of a **Cell** object, except that it can be used to set the results (values) of many cells at once, rather than one cell at a time.

For  **Style** objects, you can use the **SetResults** method to set results of any set of cells. You tell the **SetResults** method which cells you want to set by passing an array of integers in _SID_SRCStream()_.  _SID_SRCStream()_ is a one-dimensional array of 2-byte integers.

For  **Style** objects, _SID_SRCStream()_ should be a one-dimensional array of 3 _n_ 2-byte integers for _n_ >= 1. The **SetResults** method interprets the stream as:




```
{sectionIdx, rowIdx, cellIdx }n
```

where  _sectionIdx_ is the section index of the desired cell, _rowIdx_ is its row index, and _cellIdx_ is its cell index.

The  _UnitsNamesOrCodes()_ array controls what measurement units individual entries in results are in. Each entry in the array can be a string such as "inches", "inch", "in.", or "i". Strings may be used for all supported Microsoft Office Visio units such as centimeters, meters, miles, and so on. You can also indicate desired units with integer constants ( **visCentimeters** , **visInches** , and so on) declared by the Visio type library in **VisUnitCodes** . For a list of constants used for units of measure, see[About Units of Measure](http://msdn.microsoft.com/library/b6140312-b8e6-0cf2-9fe0-b14e800216bf%28Office.15%29.aspx). Note that the values specified in the  _UnitsNamesOrCodes()_ array have no effect if **visSetFormulas** is set in _Flags_.

If  _UnitsNamesOrCodes()_ is not empty, we expect it to be a one-dimensional array of 1 <= _u_ variants. Each entry can be a string or integer code, or empty (nothing). If the _i_ 'th entry is empty, the _i_ 'th entry in _resultArray()_ is in the units designated by _units(j)_ , where _j_ is the most recent prior entry that is not empty. Thus, if you want all entries in _resultArray()_ to be interpreted in the same units, you need only pass a _UnitsNamesOrCodes()_ array that has one entry. If there is no prior entry that is not empty, or if no _units_ array is supplied, **visNumber** (0x20) will be used. This causes the application to default to internal units (as does the **ResultIU** property of a **Cell** object).

The  _resultArray()_ parameter should be a one-dimensional array of 1 <= _m_ variants. A result can be passed as **Double** , **Integer** , **String** , or a reference to a **String** . Strings are accepted only if **visSetFormulas** is set in _Flags_, in which case strings are interpreted as formulas. If  _resultArray(i)_ is empty, the _i_ 'th cell will be set to the value in _resultArray(j)_, where  _j_ is the index of the most recent prior entry that is not empty. If there is no prior entry that is not empty, the corresponding cell is not altered. If fewer results than cells are specified (if _m < n_ ), the _i_ 'th cell, _i < m_ , will be set to the same value as the _m_ 'th cell. Thus, to set many cells to the same value, you need only pass one copy of the value.

The  _Flags_ parameter should be a bitmask of the following values.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visSetFormulas**|&;H1|Treat strings in results as formulas.|
| **visSetBlastGuards**|&;H2|Override present cell values even if they're guarded.|
| **visSetTestCircular**|&;H4|Test for establishment of circular cell references.|
| **visSetUniversalSyntax**|&;H8|Formulas are in universal syntax|
The value returned by the  **SetResults** method is the number of entries in _SID_SRCStream()_ that were successfully processed. If _i < n_ entries are processed correctly, but an error occurs on the _i_ + 1st entry, the **SetResults** method raises an exception and returns _i_ . Otherwise, _n_ is returned.


## Example

The following example shows how to use the  **SetResults** method. This example assumes there is an active page that has at least 3 shapes on it. It uses the **GetResults** method to get the width of shape 1, the height of shape 2, and the angle of shape 3. It then uses **SetResults** to set the width of shape 1 to the height of shape 2 and the height of shape 2 to the width of shape 1.The angle of shape 3 is left unaltered.

This example uses the  **GetResults** method of the **Page** object to get 3 cell formulas and the **SetResults** method of the same object to set the formulas. The input array has 4 slots for each cell, as it also would for **Master** objects. For **Shape** or **Style** objects, only 3 slots are needed for each cell (section, row, and cell).




```vb
 
Public Sub Set Results_Example() 
 
 On Error GoTo HandleError 
 
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
 
 'Get the first two values in inches. The second element in 
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
 
 'Use SetResults to: 
 
 ' - Set the width of shape 1 to the height of shape 2. 
 
 ' - Set the height of shape 2 to the width of shape 1. 
 
 'NOTE: avarResults() is indexed from 0 to 2. 
 
 Dim varTemp As variant 
 varTemp = avarResults(0) 
 avarResults(0) = avarResults(1) 
 avarResults(1) = varTemp 
 
 'Pass the same array back to SetResults that we 
 'just passed to GetResults, but leave the angle 
 'alone. By setting the sheet ID entry in the third 
 'slot of the aintSheetSectionRowColumn array to 
 'visInvalShapeID, we tell SetResults to ignore that slot. 
 aintSheetSectionRowColumn(9) = visInvalShapeID 
 
 'Set the results of the cells. 
 ActivePage.SetResults aintSheetSectionRowColumn, avarUnits, avarResults, 0 
 
 Exit Sub 
 
HandleError: 
 
 MsgBox "Error" 
 
 Exit Sub 
 
End Sub
```


