
# Range.PasteSpecial Method (Excel)

Pastes a  **[Range](b8207778-0dcc-4570-1234-f130532cc8cd.md)** that has been copied into the specified range.


## Syntax

 _expression_ . **PasteSpecial**( **_Paste_** , **_Operation_** , **_SkipBlanks_** , **_Transpose_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Paste_|Optional| **[XlPasteType](a60202d9-b380-ed88-b7d8-66bf34e032a5.md)**|. The part of the range to be pasted.|
| _Operation_|Optional| **[XlPasteSpecialOperation](b1e01a39-61b8-a3a9-2552-58d79b10afe3.md)**|. The paste operation.|
| _SkipBlanks_|Optional| **Variant**| **True** to have blank cells in the range on the Clipboard not be pasted into the destination range. The default value is **False** .|
| _Transpose_|Optional| **Variant**| **True** to transpose rows and columns when the range is pasted.The default value is **False** .|

### Return Value

Variant


## Example

This example replaces the data in cells D1:D5 on Sheet1 with the sum of the existing contents and cells C1:C5 on Sheet1.


```vb
With Worksheets("Sheet1") 
 .Range("C1:C5").Copy 
 .Range("D1:D5").PasteSpecial _ 
 Operation:=xlPasteSpecialOperationAdd 
End With
```


## See also


#### Concepts


[Range Object](b8207778-0dcc-4570-1234-f130532cc8cd.md)
#### Other resources


[Range Object Members](4336bf81-1e63-7e44-1792-baf366a027a7.md)
