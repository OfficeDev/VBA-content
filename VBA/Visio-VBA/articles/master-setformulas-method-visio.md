---
title: Master.SetFormulas Method (Visio)
keywords: vis_sdr.chm10716575
f1_keywords:
- vis_sdr.chm10716575
ms.prod: visio
api_name:
- Visio.Master.SetFormulas
ms.assetid: fb419eb5-6bd3-cfc7-d358-cef9e68dddbf
ms.date: 06/08/2017
---


# Master.SetFormulas Method (Visio)

Sets the formulas of one or more cells.


## Syntax

 _expression_ . **SetFormulas**( **_SID_SRCStream()_** , **_formulaArray()_** , **_Flags_** )

 _expression_ A variable that represents a **Master** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SID_SRCStream()_|Required| **Integer**|A stream identifying cells to be modified.|
| _formulaArray()_|Required| **Variant**|Formulas to be assigned to identified cells.|
| _Flags_|Required| **Integer**|Flags that influence the behavior of  **SetFormulas** .|

### Return Value

Integer


## Remarks

The  **SetFormulas** method behaves like the **Formula** property except that you can use it to set the formulas of many cells at once, rather than one cell at a time.

For  **Master** objects, you can use the **SetFormulas** method to set results of any set of cells in any set of shapes of the page or master. You tell the **SetFormulas** method which cells you want to set by passing an array of integers in _SID_SRCStream()_.  _SID_SRCStream()_ is a one-dimensional array of 2-byte integers.

For  **Master** objects _SID_SRCStream()_ should be a one-dimensional array of 4 _n_ 2-byte integers for _n_ >= 1. The **SetFormulas** method interprets the stream as:




```
{sheetID, sectionIdx, rowIdx, cellIdx }n
```

where  _sheetID_ is the **ID** property of the **Shape** object on the page or master whose cell result is to be modified.

If the  _sheetID_ in an entry is **visInvalShapeID** (-1) or if the bottom byte of _sectionIdx_ is **visSectionInval** (255), the entry is ignored by the **SetResults** method. This is because the same _SID_SRCStream_ array can be used on several calls to **SetFormulas** , **GetFormulas** , and similar methods with the caller only needing to make minor changes to the stream between calls.

The  _formulaArray()_ parameter should be a one-dimensional array of 1 <= _m_ variants. Each **Variant** should be a **String** , a reference to a **String** , or **Empty** . If _formulaArray(i)_ is empty, the _i_ 'th cell will be set to the formula in _formulaArray(j)_, where  _j_ is the index of the most recent prior entry that is not empty. If there is no prior entry that is not empty, the corresponding cell is not altered. If fewer formulas than cells are specified ( _m_ < _n_ ), the _i_ 'th cell, _i_ > _m_ , will be set to the same formula as was chosen to set the _m_ 'th cell to. Thus to set many cells to the same formula, you need pass only one copy of the formula.

The  _Flags_ argument should be a bitmask of the following values.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visSetBlastGuards**|&;H2|Override present cell values even if they're guarded.|
| **visSetTestCircular**|&;H4|Test for establishment of circular cell references.|
| **visSetUniversalSyntax**|&;H8|Formulas are in universal syntax.|
The value returned by the  **SetFormulas** method is the number of entries in _SID_SRCStream()_ that were successfully processed. If _i_ < _n_ entries process correctly, but an error occurs on the _i_ + 1st entry, the **SetFormulas** method raises an exception and returns _i_ . Otherwise, _n_ is returned.


## Example

The following macro shows how to use the  **SetFormulas** method. It assumes that there is an active Microsoft Office Visio page that has at least three shapes on it. It uses the **GetFormulas** method to get the width of shape 1, the height of shape 2, and the angle of shape 3. It then uses **SetFormulas** to set the width of shape 1 to the height of shape 2 and the height of shape 2 to the width of shape 1. The angle of shape 3 is left unaltered.

This example uses the  **GetFormulas** method of the **Page** object to get three cell formulas and the **SetFormulas** method of the same object to set the formulas. The input array has four slots for each cell, as it also would for **Master** objects. For **Shape** or **Style** objects, only three slots are needed for each cell (section, row, and cell).




```vb
 
Public Sub SetFormulas_Example() 
 
 On Error GoTo HandleError 
 
 Dim aintSheetSectionRowColumn(1 To 3 * 4) As Integer 
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
 
 'Return the formulas of the cells. 
 Dim avarFormulaArray() As Variant 
 ActivePage.GetFormulas aintSheetSectionRowColumn, avarFormulaArray 
 
 'Use SetFormulas to: 
 ' - Set the width of shape 1 to height of shape 2. 
 ' - Set height of shape 2 to width of shape 1. 
 ' Note: avarFormulaArray() is indexed from 0 to 2. 
 Dim varTemp As variant 
 varTemp = avarFormulaArray(0) 
 avarFormulaArray(0) = avarFormulaArray(1) 
 avarFormulaArray(1) = varTemp 
 
 'Pass the same array back to SetFormulas that we 
 'just passed to GetFormulas, leaving angle alone. By setting 
 'the sheet ID entry in the third slot of the 
 'aintSheetSectionRowColumn array to visInvalShapeID, 
 'we tell SetFormulas to ignore that slot. 
 aintSheetSectionRowColumn (9) = visInvalShapeID 
 
 'Tell Microsoft Visio to set the formulas of the cells. 
 ActivePage.SetFormulas aintSheetSectionRowColumn, avarFormulaArray, 0 
 
 Exit Sub 
 
HandleError: 
 
 MsgBox "Error" 
 
 Exit Sub 
 
End Sub
```


