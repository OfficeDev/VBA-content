
# Workbook.ChangeFileAccess Method (Excel)

Changes the access permissions for the workbook. This may require an updated version to be loaded from the disk.


## Syntax

 _expression_ . **ChangeFileAccess**( **_Mode_** , **_WritePassword_** , **_Notify_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Mode_|Required| **[XlFileAccess](7b4a7dc7-11c2-dea9-5e04-dcabe6530ee0.md)**|Specifies the new access mode.|
| _WritePassword_|Optional| **Variant**|Specifies the write-reserved password if the file is write reserved and  _Mode_ is **xlReadWrite** . Ignored if there's no password for the file or if _Mode_ is **xlReadOnly** .|
| _Notify_|Optional| **Variant**| **True** (or omitted) to notify the user if the file cannot be immediately accessed.|

## Remarks

If you have a file open in read-only mode, you don't have exclusive access to the file. If you change a file from read-only to read/write, Microsoft Excel must load a new copy of the file to ensure that no changes were made while you had the file open as read-only.


## Example

This example sets the active workbook to read-only.


```vb
ActiveWorkbook.ChangeFileAccess Mode:=xlReadOnly
```


## See also


#### Concepts


[Workbook Object](8c00aa60-c974-eed3-0812-3c9625eb0d4c.md)
#### Other resources


[Workbook Object Members](dce102a3-25de-3ff4-2ce5-bc56e08baca7.md)
