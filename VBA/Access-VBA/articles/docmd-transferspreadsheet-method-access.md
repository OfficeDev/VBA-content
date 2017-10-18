---
title: DoCmd.TransferSpreadsheet Method (Access)
keywords: vbaac10.chm4189
f1_keywords:
- vbaac10.chm4189
ms.prod: access
api_name:
- Access.DoCmd.TransferSpreadsheet
ms.assetid: 0349d8e0-9363-0eda-4efb-a73c9e643823
ms.date: 06/08/2017
---


# DoCmd.TransferSpreadsheet Method (Access)

The **TransferSpreadsheet** method carries out the TransferSpreadsheet action in Visual Basic.


## Syntax

 _expression_. **TransferSpreadsheet** (**_TransferType_**, **_SpreadsheetType_**, **_TableName_**, **_FileName_**, **_HasFieldNames_**, **_Range_**, **_UseOA_**)

 _expression_ A variable that represents a **DoCmd** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _TransferType_|Optional|[AcDataTransferType](acdatatransfertype-enumeration-access.md)|The type of transfer you want to make. The default value is  **acImport**.|
| _SpreadsheetType_|Optional|[AcSpreadSheetType](acspreadsheettype-enumeration-access.md)|The type of spreadsheet to import from, export to, or link to. |
| _TableName_|Optional|**Variant**|A string expression that is the name of the Office Access table that you want to import spreadsheet data into, export spreadsheet data from, or link spreadsheet data to, or the Access select query whose results you want to export to a spreadsheet.|
| _FileName_|Optional|**Variant**|A string expression that's the file name and path of the spreadsheet that you want to import from, export to, or link to.|
| _HasFieldNames_|Optional|**Variant**|Use **True** (1) to use the first row of the spreadsheet as field names when importing or linking. Use **False** (0) to treat the first row of the spreadsheet as normal data. If you leave this argument blank, the default (**False**) is assumed. When you export Access table or select query data to a spreadsheet, the field names are inserted into the first row of the spreadsheet no matter what you enter for this argument.|
| _Range_|Optional|**Variant**|A string expression that's a valid range of cells or the name of a range in the spreadsheet. This argument applies only to importing. Leave this argument blank to import the entire spreadsheet. When you export to a spreadsheet, you must leave this argument blank. If you enter a range, the export will fail.|
| _UseOA_|Optional|**Variant**|This argument is not supported.|

## Remarks

You can use the **TransferSpreadsheet** method to import or export data between the current Access database or Access project (.adp) and a spreadsheet file. You can also link the data in an Excel spreadsheet to the current Access database. With a linked spreadsheet, you can view and edit the spreadsheet data with Access while still allowing complete access to the data from your Excel spreadsheet program. You can also link to data in a Lotus 1-2-3 spreadsheet file, but this data is read-only in Access.

> [!NOTE]
> You can also use ActiveX Data Objects (ADO) to create a link by using the **ActiveConnection** property for the **Recordset** object.


## Example

The following example imports the data from the specified range of the Lotus spreadsheet Newemps.wk3 into the Access Employees table. It uses the first row of the spreadsheet as field names.

```vb
DoCmd.TransferSpreadsheet acImport, 3, _ 
 "Employees","C:\Lotus\Newemps.wk3", True, "A1:G12"
```


## See also

#### Concepts

[DoCmd Object](docmd-object-access.md)

