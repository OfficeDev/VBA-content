---
title: Range Object (Excel)
keywords: vbaxl10.chm143072
f1_keywords:
- vbaxl10.chm143072
ms.prod: excel
api_name:
- Excel.Range
ms.assetid: b8207778-0dcc-4570-1234-f130532cc8cd
ms.date: 06/08/2017
---


# Range Object (Excel)

Represents a cell, a row, a column, a selection of cells containing one or more contiguous blocks of cells, or a 3-D range.


## Example

Use  **Range** ( _arg_ ), where _arg_ names the range, to return a **Range** object that represents a single cell or a range of cells. The following example places the value of cell A1 in cell A5.


```
Worksheets("Sheet1").Range("A5").Value = _ 
    Worksheets("Sheet1").Range("A1").Value
```

The following example fills the range A1:H8 with random numbers by setting the formula for each cell in the range. When it's used without an object qualifier (an object to the left of the period), the  **Range** property returns a range on the active sheet. If the active sheet isn't a worksheet, the method fails. Use the **[Activate](http://msdn.microsoft.com/library/b198dc36-99d0-42db-6cbb-7f68396fd2f5%28Office.15%29.aspx)** method to activate a worksheet before you use the **Range** property without an explicit object qualifier.




```
Worksheets("Sheet1").Activate 
Range("A1:H8").Formula = "=Rand()"    'Range is on the active sheet
```

The following example clears the contents of the range named  _Criteria_.


 **Note**  If you use a text argument for the range address, you must specify the address in A1-style notation (you cannot use R1C1-style notation).




```
Worksheets(1).Range("Criteria").ClearContents
```

Use  **Cells** ( _row_, _column_ ) where _row_ is the row index and _column_ is the column index, to return a single cell. The following example sets the value of cell A1 to 24.




```
Worksheets(1).Cells(1, 1).Value = 24
```

The following example sets the formula for cell A2.




```
ActiveSheet.Cells(2, 1).Formula = "=Sum(B1:B5)"
```

Although you can also use  `Range("A1")` to return cell A1, there may be times when the **Cells** property is more convenient because you can use a variable for the row or column. The following example creates column and row headings on Sheet1. Be aware that after the worksheet has been activated, the **Cells** property can be used without an explicit sheet declaration (it returns a cell on the active sheet).


 **Note**  Although you could use Visual Basic string functions to alter A1-style references, it is easier (and better programming practice) to use the  `Cells(1, 1)` notation.




```
Sub SetUpTable() 
Worksheets("Sheet1").Activate 
For TheYear = 1 To 5 
    Cells(1, TheYear + 1).Value = 1990 + TheYear 
Next TheYear 
For TheQuarter = 1 To 4 
    Cells(TheQuarter + 1, 1).Value = "Q" &amp; TheQuarter 
Next TheQuarter 
End Sub
```

Use  _expression_. **Cells** ( _row_, _column_ ), where _expression_ is an expression that returns a **Range** object, and _row_ and _column_ are relative to the upper-left corner of the range, to return part of a range. The following example sets the formula for cell C5.




```
Worksheets(1).Range("C5:C10").Cells(1, 1).Formula = "=Rand()"
```

Use  **Range** ( _cell1, cell2_ ), where _cell1_ and _cell2_ are **Range** objects that specify the start and end cells, to return a **Range** object. The following example sets the border line style for cells A1:J10.


 **Note**  Be aware that the period in front of each occurrence of the  **Cells** property. The period is required if the result of the preceding **With** statement is to be applied to the **Cells** property—in this case, to indicate that the cells are on worksheet one (without the period, the **Cells** property would return cells on the active sheet).




```
With Worksheets(1) 
    .Range(.Cells(1, 1), _ 
        .Cells(10, 10)).Borders.LineStyle = xlThick 
End With
```

Use  **Offset** ( _row, column_ ), where _row_ and _column_ are the row and column offsets, to return a range at a specified offset to another range. The following example selects the cell three rows down from and one column to the right of the cell in the upper-left corner of the current selection. You cannot select a cell that is not on the active sheet, so you must first activate the worksheet.




```
Worksheets("Sheet1").Activate 
  'Can't select unless the sheet is active 
Selection.Offset(3, 1).Range("A1").Select
```

Use  **Union** ( _range1, range2_, ...) to return multiple-area ranges—that is, ranges composed of two or more contiguous blocks of cells. The following example creates an object defined as the union of ranges A1:B2 and C3:D4, and then selects the defined range.




```
Dim r1 As Range, r2 As Range, myMultiAreaRange As Range 
Worksheets("sheet1").Activate 
Set r1 = Range("A1:B2") 
Set r2 = Range("C3:D4") 
Set myMultiAreaRange = Union(r1, r2) 
myMultiAreaRange.Select
```

If you work with selections that contain more than one area, the  **[Areas](http://msdn.microsoft.com/library/31fc03b4-25b6-27ae-2350-b34c6c6ba255%28Office.15%29.aspx)** property is useful. It divides a multiple-area selection into individual **Range** objects and then returns the objects as a collection. You can use the **[Count](http://msdn.microsoft.com/library/080cbbe7-056f-b21c-9004-171a6acce664%28Office.15%29.aspx)** property on the returned collection to verify a selection that contains more than one area, as shown in the following example.




```
Sub NoMultiAreaSelection() 
    NumberOfSelectedAreas = Selection.Areas.Count 
    If NumberOfSelectedAreas > 1 Then 
        MsgBox "You cannot carry out this command " &amp; _ 
            "on multi-area selections" 
    End If 
End Sub
```

 **Sample code provided by:** Dennis Wallentin,[VSTO &amp; .NET &amp; Excel](http://xldennis.wordpress.com/)

This example uses the  **AdvancedFilter** method of the **Range** object to create a list of the unique values, and the number of times those unique values occur, in the range of column A.




```
Sub Create_Unique_List_Count()
    'Excel workbook, the source and target worksheets, and the source and target ranges.
    Dim wbBook As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim rnSource As Range
    Dim rnTarget As Range
    Dim rnUnique As Range
    'Variant to hold the unique data
    Dim vaUnique As Variant
    'Number of unique values in the data
    Dim lnCount As Long
    
    'Initialize the Excel objects
    Set wbBook = ThisWorkbook
    With wbBook
        Set wsSource = .Worksheets("Sheet1")
        Set wsTarget = .Worksheets("Sheet2")
    End With
    
    'On the source worksheet, set the range to the data stored in column A
    With wsSource
        Set rnSource = .Range(.Range("A1"), .Range("A100").End(xlDown))
    End With
    
    'On the target worksheet, set the range as column A.
    Set rnTarget = wsTarget.Range("A1")
    
    'Use AdvancedFilter to copy the data from the source to the target,
    'while filtering for duplicate values.
    rnSource.AdvancedFilter Action:=xlFilterCopy, _
                            CopyToRange:=rnTarget, _
                            Unique:=True
                            
    'On the target worksheet, set the unique range on Column A, excluding the first cell
    '(which will contain the "List" header for the column).
    With wsTarget
        Set rnUnique = .Range(.Range("A2"), .Range("A100").End(xlUp))
    End With
    
    'Assign all the values of the Unique range into the Unique variant.
    vaUnique = rnUnique.Value
    
    'Count the number of occurrences of every unique value in the source data,
    'and list it next to its relevant value.
    For lnCount = 1 To UBound(vaUnique)
        rnUnique(lnCount, 1).Offset(0, 1).Value = _
            Application.Evaluate("COUNTIF(" &amp; _
            rnSource.Address(External:=True) &amp; _
            ",""" &amp; rnUnique(lnCount, 1).Text &amp; """)")
    Next lnCount
    
    'Label the column of occurrences with "Occurrences"
    With rnTarget.Offset(0, 1)
        .Value = "Occurrences"
        .Font.Bold = True
    End With

End Sub
```


## Remarks

The following properties and methods for returning a  **Range** object are described in the examples section:


-  **[Range](http://msdn.microsoft.com/library/9a323305-c822-ef9e-1cc8-ec077a976834%28Office.15%29.aspx)** property
    
-  **[Cells](http://msdn.microsoft.com/library/19c14e41-7d8e-b56f-fd60-717df64edee8%28Office.15%29.aspx)** property
    
-  **Range** and **Cells**
    
-  **[Offset](http://msdn.microsoft.com/library/dfbbd1a2-2f73-fd6a-6277-4584823f55a4%28Office.15%29.aspx)** property
    
-  **[Union](http://msdn.microsoft.com/library/7c70a5be-2696-5fc2-bd69-6c6ff4d3291e%28Office.15%29.aspx)** method
    

## Methods



|**Name**|
|:-----|
|[Activate](http://msdn.microsoft.com/library/a0050055-84e7-7611-a961-887fcb063369%28Office.15%29.aspx)|
|[AddComment](http://msdn.microsoft.com/library/89bbacad-4655-bcc1-8010-2ab367cc7b31%28Office.15%29.aspx)|
|[AdvancedFilter](http://msdn.microsoft.com/library/fe1a19fc-ab0f-6149-25d9-6102d5789757%28Office.15%29.aspx)|
|[AllocateChanges](http://msdn.microsoft.com/library/c751c5fb-ce22-64d1-669c-fdb064cf0408%28Office.15%29.aspx)|
|[ApplyNames](http://msdn.microsoft.com/library/3798ecfb-c839-64a9-1088-d7752a3e81ae%28Office.15%29.aspx)|
|[ApplyOutlineStyles](http://msdn.microsoft.com/library/eab9b4ed-5d4c-8205-63f2-fa8e4539da73%28Office.15%29.aspx)|
|[AutoComplete](http://msdn.microsoft.com/library/723a452f-34e1-fcd1-a2d6-4932c5cc0542%28Office.15%29.aspx)|
|[AutoFill](http://msdn.microsoft.com/library/257f6608-9211-86f9-79de-e3c44df8f3fd%28Office.15%29.aspx)|
|[AutoFilter](http://msdn.microsoft.com/library/0f773dbf-63e8-f714-d246-f803a74d366c%28Office.15%29.aspx)|
|[AutoFit](http://msdn.microsoft.com/library/53a35cd3-00e7-f9f5-2cd2-8492d7814a11%28Office.15%29.aspx)|
|[AutoOutline](http://msdn.microsoft.com/library/a2553695-6d45-9b7c-7c45-5255fa3641f0%28Office.15%29.aspx)|
|[BorderAround](http://msdn.microsoft.com/library/3ffeb131-45f7-7799-e04a-11577fedaa16%28Office.15%29.aspx)|
|[Calculate](http://msdn.microsoft.com/library/7c29afda-4980-6992-fc8d-b4caf2f74660%28Office.15%29.aspx)|
|[CalculateRowMajorOrder](http://msdn.microsoft.com/library/8636b550-a3f8-f6cd-baf8-b669354262af%28Office.15%29.aspx)|
|[CheckSpelling](http://msdn.microsoft.com/library/22540515-0b0f-ce3c-0145-e46d477d1b5f%28Office.15%29.aspx)|
|[Clear](http://msdn.microsoft.com/library/56f46ac7-8bb0-2651-8024-312c7cb7356c%28Office.15%29.aspx)|
|[ClearComments](http://msdn.microsoft.com/library/736fd51f-a7cd-02cf-eb45-47e3f3132800%28Office.15%29.aspx)|
|[ClearContents](http://msdn.microsoft.com/library/8c957fdd-e99d-ca0e-7d2c-4cb1db62639a%28Office.15%29.aspx)|
|[ClearFormats](http://msdn.microsoft.com/library/37777379-857a-f4c7-86aa-b109d5f25757%28Office.15%29.aspx)|
|[ClearHyperlinks](http://msdn.microsoft.com/library/1bf0613b-4415-a9cc-88e0-5e71f0ab812b%28Office.15%29.aspx)|
|[ClearNotes](http://msdn.microsoft.com/library/24017be9-d3bf-2e8a-4587-d5b0a03fdcaf%28Office.15%29.aspx)|
|[ClearOutline](http://msdn.microsoft.com/library/80d82c8d-7670-54b5-7aa5-5c39aadcb990%28Office.15%29.aspx)|
|[ColumnDifferences](http://msdn.microsoft.com/library/483995e1-9c8d-c171-4c72-17afd5918d49%28Office.15%29.aspx)|
|[Consolidate](http://msdn.microsoft.com/library/d5fb78a3-c3ec-0d1a-c6ad-b33bc90e431c%28Office.15%29.aspx)|
|[Copy](http://msdn.microsoft.com/library/ac5207ac-6be5-3c7e-2c61-67954a59e9df%28Office.15%29.aspx)|
|[CopyFromRecordset](http://msdn.microsoft.com/library/cec7fded-f4e0-1b1c-5374-8a860828c9cc%28Office.15%29.aspx)|
|[CopyPicture](http://msdn.microsoft.com/library/0b187b51-7a52-0db3-9d55-9c1e5bc5e49b%28Office.15%29.aspx)|
|[CreateNames](http://msdn.microsoft.com/library/00c7c74f-606d-7eee-ac52-f6b21446f5be%28Office.15%29.aspx)|
|[Cut](http://msdn.microsoft.com/library/b9f525c4-c314-450c-f88b-e6c5cdc00112%28Office.15%29.aspx)|
|[DataSeries](http://msdn.microsoft.com/library/cfdb0582-8b6c-029d-2a99-4fa1d4b360ea%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/7d890cc5-5b5b-35f9-2d97-e4fe48f244ee%28Office.15%29.aspx)|
|[DialogBox](http://msdn.microsoft.com/library/d2d4a677-bd6a-910d-ff53-f95585f40925%28Office.15%29.aspx)|
|[Dirty](http://msdn.microsoft.com/library/c3f177ef-19b9-07e7-a42f-978874528207%28Office.15%29.aspx)|
|[DiscardChanges](http://msdn.microsoft.com/library/adeee827-d680-59f3-0966-2c2ca60a59e1%28Office.15%29.aspx)|
|[EditionOptions](http://msdn.microsoft.com/library/5997563b-7f39-6f2d-9265-c72a2d138548%28Office.15%29.aspx)|
|[ExportAsFixedFormat](http://msdn.microsoft.com/library/9786c633-e9bd-3ce3-0246-7bcb3c4b4ce1%28Office.15%29.aspx)|
|[FillDown](http://msdn.microsoft.com/library/bb7c0b2d-8dd9-13e5-b90a-b2708935afa9%28Office.15%29.aspx)|
|[FillLeft](http://msdn.microsoft.com/library/42722b18-8b40-c27b-8bca-ef180cf0f636%28Office.15%29.aspx)|
|[FillRight](http://msdn.microsoft.com/library/b0b9a3a5-5f8c-327e-fb41-dec5c1a2f2b3%28Office.15%29.aspx)|
|[FillUp](http://msdn.microsoft.com/library/52498f52-95f9-5993-7c44-76cd8b696074%28Office.15%29.aspx)|
|[Find](http://msdn.microsoft.com/library/d9585265-8164-cb4d-a9e3-262f6e06b6b8%28Office.15%29.aspx)|
|[FindNext](http://msdn.microsoft.com/library/308c6241-2398-13e6-ba68-17ec713376f6%28Office.15%29.aspx)|
|[FindPrevious](http://msdn.microsoft.com/library/c03f2e17-d28c-8b0d-b8c8-024863523c99%28Office.15%29.aspx)|
|[FlashFill](http://msdn.microsoft.com/library/3ca4a73f-712a-fe69-684d-a959351e5855%28Office.15%29.aspx)|
|[FunctionWizard](http://msdn.microsoft.com/library/a9a0c765-4903-4969-8f09-c8f051213a96%28Office.15%29.aspx)|
|[Group](http://msdn.microsoft.com/library/da736f64-35df-ecaf-88ac-8c61f7d3c0d0%28Office.15%29.aspx)|
|[Insert](http://msdn.microsoft.com/library/e612bbc8-3942-3349-f157-c0459805794a%28Office.15%29.aspx)|
|[InsertIndent](http://msdn.microsoft.com/library/1e004333-a64e-55e4-cf8a-d15e47236f94%28Office.15%29.aspx)|
|[Justify](http://msdn.microsoft.com/library/f8b4d48b-8cbb-977a-fd44-d354661182d2%28Office.15%29.aspx)|
|[ListNames](http://msdn.microsoft.com/library/0523f9b3-d422-76b6-889c-75619cb5b9a6%28Office.15%29.aspx)|
|[Merge](http://msdn.microsoft.com/library/eff315d8-fa8f-e452-2bcd-15be4d97a077%28Office.15%29.aspx)|
|[NavigateArrow](http://msdn.microsoft.com/library/71e2ce3b-3da8-afd5-7fd3-b922c6f8f1c2%28Office.15%29.aspx)|
|[NoteText](http://msdn.microsoft.com/library/cd0e5073-7d04-a52c-f375-f7c59bc8f88a%28Office.15%29.aspx)|
|[Parse](http://msdn.microsoft.com/library/3580aeb7-e868-894a-9dd5-8e37475fb267%28Office.15%29.aspx)|
|[PasteSpecial](http://msdn.microsoft.com/library/d3e991f2-7ef7-2ebc-d4bc-ba4c26be472e%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/42d36dbb-5910-530f-5aea-3793a36dc82b%28Office.15%29.aspx)|
|[PrintPreview](http://msdn.microsoft.com/library/b429a45c-864f-1c48-0775-1cf240f6e7ac%28Office.15%29.aspx)|
|[RemoveDuplicates](http://msdn.microsoft.com/library/0e74bde2-08b3-898d-0b30-53de911bd7e9%28Office.15%29.aspx)|
|[RemoveSubtotal](http://msdn.microsoft.com/library/ec1fd131-551d-009f-1eea-033d805bb34d%28Office.15%29.aspx)|
|[Replace](http://msdn.microsoft.com/library/12647334-f911-69e4-de31-b4df2722eff3%28Office.15%29.aspx)|
|[RowDifferences](http://msdn.microsoft.com/library/89030ca3-9f59-7426-d050-89dcabf00887%28Office.15%29.aspx)|
|[Run](http://msdn.microsoft.com/library/b7a0480a-9f10-8aad-6592-3cbde72720cd%28Office.15%29.aspx)|
|[Select](http://msdn.microsoft.com/library/46c12f85-fae5-15ea-3500-81ff8be49cdb%28Office.15%29.aspx)|
|[SetPhonetic](http://msdn.microsoft.com/library/69a1e491-5505-621a-5ea0-b0600796caa3%28Office.15%29.aspx)|
|[Show](http://msdn.microsoft.com/library/c04cbae7-c424-befd-df73-e92bbe9e2e41%28Office.15%29.aspx)|
|[ShowDependents](http://msdn.microsoft.com/library/f2e062b2-733b-d0e5-b5ed-9587b104bbc7%28Office.15%29.aspx)|
|[ShowErrors](http://msdn.microsoft.com/library/02366ef0-b4dc-a10c-e186-d9392a8b656c%28Office.15%29.aspx)|
|[ShowPrecedents](http://msdn.microsoft.com/library/02b8ca94-d251-a6be-1551-1ba769c3c0fa%28Office.15%29.aspx)|
|[Sort](http://msdn.microsoft.com/library/ede52b2f-9025-fc83-9c16-f09c6b89c5c2%28Office.15%29.aspx)|
|[SortSpecial](http://msdn.microsoft.com/library/706420cb-989a-1b48-b051-ca6e5fe45824%28Office.15%29.aspx)|
|[Speak](http://msdn.microsoft.com/library/12928814-9534-c9f0-e351-7d26f77869e0%28Office.15%29.aspx)|
|[SpecialCells](http://msdn.microsoft.com/library/30c2035c-34e3-3b1a-f243-69a9fed97f3b%28Office.15%29.aspx)|
|[SubscribeTo](http://msdn.microsoft.com/library/6120c474-f4a9-0dce-dae4-a8b39f3d3656%28Office.15%29.aspx)|
|[Subtotal](http://msdn.microsoft.com/library/b4b7b640-5a6c-8c94-d9ab-c9a557190829%28Office.15%29.aspx)|
|[Table](http://msdn.microsoft.com/library/804b0e1d-e92d-387d-1054-90643bfd16ff%28Office.15%29.aspx)|
|[TextToColumns](http://msdn.microsoft.com/library/0b0bf329-ab99-7edc-1b8f-aad03513abde%28Office.15%29.aspx)|
|[Ungroup](http://msdn.microsoft.com/library/ac20c780-1a8e-2709-13c4-a6ca8220fb0a%28Office.15%29.aspx)|
|[UnMerge](http://msdn.microsoft.com/library/dfc49876-29b0-0b61-fe18-3953438f7452%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[AddIndent](http://msdn.microsoft.com/library/47cfb2a4-9050-354f-08f6-e86f0164be02%28Office.15%29.aspx)|
|[Address](http://msdn.microsoft.com/library/aaa2432e-9bb1-4a48-3868-86455bc53938%28Office.15%29.aspx)|
|[AddressLocal](http://msdn.microsoft.com/library/20332d15-dc37-1819-472f-ef208d8b3a5b%28Office.15%29.aspx)|
|[AllowEdit](http://msdn.microsoft.com/library/9f03054c-190f-ce3b-54db-bc6e19b7e1c6%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/10a5b6f8-2ded-be6b-352e-5df9d43c30ed%28Office.15%29.aspx)|
|[Areas](http://msdn.microsoft.com/library/31fc03b4-25b6-27ae-2350-b34c6c6ba255%28Office.15%29.aspx)|
|[Borders](http://msdn.microsoft.com/library/6d313fed-a8f0-94ba-e239-813685cd1d58%28Office.15%29.aspx)|
|[Cells](http://msdn.microsoft.com/library/32a6ecc7-2366-2cec-1feb-0966241a435d%28Office.15%29.aspx)|
|[Characters](http://msdn.microsoft.com/library/5011b6d3-23ab-e2a8-9616-c4c73d3ae60e%28Office.15%29.aspx)|
|[Column](http://msdn.microsoft.com/library/4f540fae-fc9f-30de-5d71-f6496b78930b%28Office.15%29.aspx)|
|[Columns](http://msdn.microsoft.com/library/a1a23288-e911-909d-0bc0-48bdce2ccbac%28Office.15%29.aspx)|
|[ColumnWidth](http://msdn.microsoft.com/library/a6364bb1-2e3d-07d6-20e4-c9fa8f7c5ad3%28Office.15%29.aspx)|
|[Comment](http://msdn.microsoft.com/library/94c07e38-f232-3fba-f08c-878d3848ac55%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/080cbbe7-056f-b21c-9004-171a6acce664%28Office.15%29.aspx)|
|[CountLarge](http://msdn.microsoft.com/library/3a46ef6d-a339-b15e-990d-b11f462fb602%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/d7970f19-b10d-9101-4326-ea2d2460e849%28Office.15%29.aspx)|
|[CurrentArray](http://msdn.microsoft.com/library/147f8834-5aef-900f-75de-df91a6a76005%28Office.15%29.aspx)|
|[CurrentRegion](http://msdn.microsoft.com/library/39277cc5-07ff-8453-7330-b272b365f9dc%28Office.15%29.aspx)|
|[Dependents](http://msdn.microsoft.com/library/47813412-306a-0f99-3ca5-d354b16af468%28Office.15%29.aspx)|
|[DirectDependents](http://msdn.microsoft.com/library/266b054e-6838-ffe7-14e2-e712a149e5be%28Office.15%29.aspx)|
|[DirectPrecedents](http://msdn.microsoft.com/library/d7eebe51-3e4c-e902-e6a5-1617bd21ef4e%28Office.15%29.aspx)|
|[DisplayFormat](http://msdn.microsoft.com/library/c4e044e2-a04e-b655-2973-7e02897ca49d%28Office.15%29.aspx)|
|[End](http://msdn.microsoft.com/library/d46d75c9-b152-e93d-82c3-f59f0e7f69da%28Office.15%29.aspx)|
|[EntireColumn](http://msdn.microsoft.com/library/7be55670-75fd-fb02-dc1a-9d70e3a9d80d%28Office.15%29.aspx)|
|[EntireRow](http://msdn.microsoft.com/library/9e66da51-6cef-4109-ea4e-2acaad42aa1f%28Office.15%29.aspx)|
|[Errors](http://msdn.microsoft.com/library/88dcc606-d412-a9ce-82bc-5fbba8baae87%28Office.15%29.aspx)|
|[Font](http://msdn.microsoft.com/library/d9cb8667-6c71-d311-a6e5-1d30d5718050%28Office.15%29.aspx)|
|[FormatConditions](http://msdn.microsoft.com/library/676ffcc6-f08d-9f91-78af-7b98f8b77dca%28Office.15%29.aspx)|
|[Formula](http://msdn.microsoft.com/library/c5be8952-fc3f-bdb3-d4a6-abf9d94eab1e%28Office.15%29.aspx)|
|[FormulaArray](http://msdn.microsoft.com/library/a0c8bafb-294c-32ff-0591-1a798aebb4b4%28Office.15%29.aspx)|
|[FormulaHidden](http://msdn.microsoft.com/library/b6425c86-7e20-e34e-2d96-eb16075c20b6%28Office.15%29.aspx)|
|[FormulaLocal](http://msdn.microsoft.com/library/c69325d9-d35d-c15a-ae49-7bde2b628428%28Office.15%29.aspx)|
|[FormulaR1C1](http://msdn.microsoft.com/library/76f41bf6-94e2-2e6a-30e4-012a735a3374%28Office.15%29.aspx)|
|[FormulaR1C1Local](http://msdn.microsoft.com/library/be0e3270-43fd-e6c7-1209-11ed3204e563%28Office.15%29.aspx)|
|[HasArray](http://msdn.microsoft.com/library/fac17206-8671-6209-9133-d56da6ea2b9c%28Office.15%29.aspx)|
|[HasFormula](http://msdn.microsoft.com/library/a18bea77-cee9-ae2d-7e97-90a4205e3b1f%28Office.15%29.aspx)|
|[Height](http://msdn.microsoft.com/library/e204a719-d7de-cd18-65b9-c34575bd92e5%28Office.15%29.aspx)|
|[Hidden](http://msdn.microsoft.com/library/7e785c38-a8ae-3810-a88a-0bfb7b74e2d6%28Office.15%29.aspx)|
|[HorizontalAlignment](http://msdn.microsoft.com/library/6689de5b-60de-07db-d2b4-114f0a343ebc%28Office.15%29.aspx)|
|[Hyperlinks](http://msdn.microsoft.com/library/d77f695a-faf2-ce9c-1464-f54b76ee52c9%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/0ff7f261-8829-2858-5097-a638c01e5f3c%28Office.15%29.aspx)|
|[IndentLevel](http://msdn.microsoft.com/library/f4d5af31-904a-27eb-fb2d-e5ae38a7ebb9%28Office.15%29.aspx)|
|[Interior](http://msdn.microsoft.com/library/9599b0f7-9f52-627c-51e6-d8be8aeb9bbf%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/f7d40273-5069-8a9d-14ee-19df225f864c%28Office.15%29.aspx)|
|[Left](http://msdn.microsoft.com/library/634fa7b8-3565-6178-f564-3c5d24c16d26%28Office.15%29.aspx)|
|[ListHeaderRows](http://msdn.microsoft.com/library/d71a9b28-cd5d-677c-9ce1-f8de2b350e5f%28Office.15%29.aspx)|
|[ListObject](http://msdn.microsoft.com/library/bbc404f0-29bd-bb95-2fc8-f826992c4192%28Office.15%29.aspx)|
|[LocationInTable](http://msdn.microsoft.com/library/7a86a0fe-cd46-331e-595b-6be168091d0c%28Office.15%29.aspx)|
|[Locked](http://msdn.microsoft.com/library/93c5f21d-6429-3287-0992-c810b9a429a8%28Office.15%29.aspx)|
|[MDX](http://msdn.microsoft.com/library/6b22b79b-ce44-ce0d-0bb4-e1bf2cd83578%28Office.15%29.aspx)|
|[MergeArea](http://msdn.microsoft.com/library/68586bba-fa9c-e0d4-0eae-a08613551a2c%28Office.15%29.aspx)|
|[MergeCells](http://msdn.microsoft.com/library/42904357-5e55-1eb0-9b06-83b446fc6275%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/39d1a326-e123-443c-29c0-453f7b4a8760%28Office.15%29.aspx)|
|[Next](http://msdn.microsoft.com/library/10712827-9abd-6b8a-49e5-65e3554fcd87%28Office.15%29.aspx)|
|[NumberFormat](http://msdn.microsoft.com/library/351247d2-e4b9-64a0-6dbe-0df535fa701c%28Office.15%29.aspx)|
|[NumberFormatLocal](http://msdn.microsoft.com/library/e34e6f52-9279-7961-adfa-4aa84c44937a%28Office.15%29.aspx)|
|[Offset](http://msdn.microsoft.com/library/dfbbd1a2-2f73-fd6a-6277-4584823f55a4%28Office.15%29.aspx)|
|[Orientation](http://msdn.microsoft.com/library/4f0588b6-2570-fe2f-0cbe-09868b77cff3%28Office.15%29.aspx)|
|[OutlineLevel](http://msdn.microsoft.com/library/bdab08a4-3576-4a65-2556-43ed9e9a576e%28Office.15%29.aspx)|
|[PageBreak](http://msdn.microsoft.com/library/0bec0bba-c2c3-33cd-b39e-55971177c2c8%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/3b4433cc-ce78-b590-31b4-d74f476e104b%28Office.15%29.aspx)|
|[Phonetic](http://msdn.microsoft.com/library/9c6d1d83-b215-d60d-f78f-68e521e25368%28Office.15%29.aspx)|
|[Phonetics](http://msdn.microsoft.com/library/fdc05b76-b574-63ec-045a-42fdcfae8a9e%28Office.15%29.aspx)|
|[PivotCell](http://msdn.microsoft.com/library/976f6393-db3b-d52a-0cbc-88a73bb7c070%28Office.15%29.aspx)|
|[PivotField](http://msdn.microsoft.com/library/56003d2d-60cd-abd2-455e-4a4d3616a615%28Office.15%29.aspx)|
|[PivotItem](http://msdn.microsoft.com/library/02a41786-074b-ae34-5d2c-407006fe526d%28Office.15%29.aspx)|
|[PivotTable](http://msdn.microsoft.com/library/ae3f77dc-5098-d60f-0afc-f4f01dbc33f0%28Office.15%29.aspx)|
|[Precedents](http://msdn.microsoft.com/library/3c00cfb4-1c12-668d-a952-89f9b1ef129f%28Office.15%29.aspx)|
|[PrefixCharacter](http://msdn.microsoft.com/library/1f7d5fbc-136a-5164-4cec-0054f8bcd0b1%28Office.15%29.aspx)|
|[Previous](http://msdn.microsoft.com/library/6ee986eb-9242-63f3-6885-1ad3730f106b%28Office.15%29.aspx)|
|[QueryTable](http://msdn.microsoft.com/library/6370d43c-74b5-1bb9-f849-c70006432504%28Office.15%29.aspx)|
|[Range](http://msdn.microsoft.com/library/7edbda7c-98d9-143d-7b5e-bcfb7f237818%28Office.15%29.aspx)|
|[ReadingOrder](http://msdn.microsoft.com/library/f367af66-21c8-b63f-7a92-3756ee711b18%28Office.15%29.aspx)|
|[Resize](http://msdn.microsoft.com/library/05af0539-8aa3-c83c-1972-dfac618929b9%28Office.15%29.aspx)|
|[Row](http://msdn.microsoft.com/library/3c8d7351-4fc6-748b-c2a8-de3dab4b964e%28Office.15%29.aspx)|
|[RowHeight](http://msdn.microsoft.com/library/103c7209-9a4f-8f9c-7bdc-3013113867a5%28Office.15%29.aspx)|
|[Rows](http://msdn.microsoft.com/library/2b0541f1-119d-8535-8418-ff9482353ec1%28Office.15%29.aspx)|
|[ServerActions](http://msdn.microsoft.com/library/dffb9535-3b82-c134-82b0-b87d8bc258ec%28Office.15%29.aspx)|
|[ShowDetail](http://msdn.microsoft.com/library/1908af55-f61a-2a0f-d828-350e9a680377%28Office.15%29.aspx)|
|[ShrinkToFit](http://msdn.microsoft.com/library/fc9aed64-1000-3419-ceb7-a95c15f8a2d0%28Office.15%29.aspx)|
|[SoundNote](http://msdn.microsoft.com/library/05d40e33-b07f-5079-29da-8843e9f16820%28Office.15%29.aspx)|
|[SparklineGroups](http://msdn.microsoft.com/library/66c6ef19-08a0-91e8-6fef-e827b80d5e62%28Office.15%29.aspx)|
|[Style](http://msdn.microsoft.com/library/78c536c9-7fda-3171-2a93-5c4e57bb8207%28Office.15%29.aspx)|
|[Summary](http://msdn.microsoft.com/library/f9e18651-20b6-1094-2ee5-7cd23559498e%28Office.15%29.aspx)|
|[Text](http://msdn.microsoft.com/library/e38c15b1-5941-0a28-1acf-328bc214a2e0%28Office.15%29.aspx)|
|[Top](http://msdn.microsoft.com/library/0d67ac39-9d35-fc2e-56f1-9bd320a4e8ea%28Office.15%29.aspx)|
|[UseStandardHeight](http://msdn.microsoft.com/library/59e0be39-25ea-c18d-919d-506d4f041f45%28Office.15%29.aspx)|
|[UseStandardWidth](http://msdn.microsoft.com/library/970e3d68-3147-a52f-b831-ae7780c735e0%28Office.15%29.aspx)|
|[Validation](http://msdn.microsoft.com/library/d1cad7e6-bbfa-e280-33e7-048733efc0bc%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/23f28b24-430a-6ea4-4895-0afff8dff218%28Office.15%29.aspx)|
|[Value2](http://msdn.microsoft.com/library/0a5d7e6f-2886-5048-66ad-a5078e3465e7%28Office.15%29.aspx)|
|[VerticalAlignment](http://msdn.microsoft.com/library/b09a2dcb-b51b-b477-6247-fd5b11a67ccf%28Office.15%29.aspx)|
|[Width](http://msdn.microsoft.com/library/75c3aff6-25a0-2f64-2c25-da213b87393b%28Office.15%29.aspx)|
|[Worksheet](http://msdn.microsoft.com/library/af38bdde-d523-a4cd-929e-1f67464b2593%28Office.15%29.aspx)|
|[WrapText](http://msdn.microsoft.com/library/5e61b704-af16-7bad-5eeb-f163e3035513%28Office.15%29.aspx)|
|[XPath](http://msdn.microsoft.com/library/90a353d7-7222-b387-558a-044cb17f09b9%28Office.15%29.aspx)|

## About the Contributor
<a name="AboutContributor"> </a>

Dennis Wallentin is the author of VSTO &amp; .NET &amp; Excel, a blog that focuses on .NET Framework solutions for Excel and Excel Services. Dennis has been developing Excel solutions for over 20 years and is also the coauthor of "Professional Excel Development: The Definitive Guide to Developing Applications Using Microsoft Excel, VBA and .NET (2nd Edition)." 


