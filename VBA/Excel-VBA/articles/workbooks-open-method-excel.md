---
title: Workbooks.Open Method (Excel)
keywords: vbaxl10.chm203082
f1_keywords:
- vbaxl10.chm203082
ms.prod: excel
api_name:
- Excel.Workbooks.Open
ms.assetid: 1d1c3fca-ae1a-0a91-65a2-6f3f0fb308a0
ms.date: 06/08/2017
---


# Workbooks.Open Method (Excel)

Opens a workbook.


## Syntax

 _expression_ . **Open**( **_FileName_** , **_UpdateLinks_** , **_ReadOnly_** , **_Format_** , **_Password_** , **_WriteResPassword_** , **_IgnoreReadOnlyRecommended_** , **_Origin_** , **_Delimiter_** , **_Editable_** , **_Notify_** , **_Converter_** , **_AddToMru_** , **_Local_** , **_CorruptLoad_** )

 _expression_ A variable that represents a **Workbooks** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Optional| **Variant**| **String** . The file name of the workbook to be opened.|
| _UpdateLinks_|Optional| **Variant**|Specifies the way external references (links) in the file, such as the reference to a range in the Budget.xls workbook in the following formula =SUM([Budget.xls]Annual!C10:C25), are updated. If this argument is omitted, the user is prompted to specify how links will be updated. For more information about the values used by this parameter, see the Remarks section. If Microsoft Excel is opening a file in the WKS, WK1, or WK3 format and the UpdateLinks argument is 0, no charts are created; otherwise Microsoft Excel generates charts from the graphs attached to the file.|
| _ReadOnly_|Optional| **Variant**|True to open the workbook in read-only mode.|
| _Format_|Optional| **Variant**|If Microsoft Excel opens a text file, this argument specifies the delimiter character. If this argument is omitted, the current delimiter is used. For more information about the values used by this parameter, see the Remarks section.|
| _Password_|Optional| **Variant**|A string that contains the password required to open a protected workbook. If this argument is omitted and the workbook requires a password, the user is prompted for the password.|
| _WriteResPassword_|Optional| **Variant**|A string that contains the password required to write to a write-reserved workbook. If this argument is omitted and the workbook requires a password, the user will be prompted for the password.|
| _IgnoreReadOnlyRecommended_|Optional| **Variant**| **True** to have Microsoft Excel not display the read-only recommended message (if the workbook was saved with the **Read-Only Recommended** option).|
| _Origin_|Optional| **Variant**|If the file is a text file, this argument indicates where it originated, so that code pages and Carriage Return/Line Feed (CR/LF) can be mapped correctly. Can be one of the following  **XlPlatform** constants: **xlMacintosh** , **xlWindows** , or **xlMSDOS** . If this argument is omitted, the current operating system is used.|
| _Delimiter_|Optional| **Variant**|If the file is a text file and the Format argument is 6, this argument is a string that specifies the character to be used as the delimiter. For example, use  **Chr(9)** for tabs, use "," for commas, use ";" for semicolons, or use a custom character. Only the first character of the string is used.|
| _Editable_|Optional| **Variant**|If the file is a Microsoft Excel 4.0 add-in, this argument is  **True** to open the add-in so that it is a visible window. If this argument is **False** or omitted, the add-in is opened as hidden, and it cannot be unhidden. This option does not apply to add-ins created in Microsoft Excel 5.0 or later. If the file is an Excel template, **True** to open the specified template for editing. **False** to open a new workbook based on the specified template. The default value is **False** .|
| _Notify_|Optional| **Variant**|If the file cannot be opened in read/write mode, this argument is  **True** to add the file to the file notification list. Microsoft Excel will open the file as read-only, poll the file notification list, and then notify the user when the file becomes available. If this argument is **False** or omitted, no notification is requested, and any attempts to open an unavailable file will fail.|
| _Converter_|Optional| **Variant**|The index of the first file converter to try when opening the file. The specified file converter is tried first; if this converter does not recognize the file, all other converters are tried. The converter index consists of the row numbers of the converters returned by the  **[FileConverters](application-fileconverters-property-excel.md)** property.|
| _AddToMru_|Optional| **Variant**| **True** to add this workbook to the list of recently used files. The default value is **False** .|
| _Local_|Optional| **Variant**| **True** saves files against the language of Microsoft Excel (including control panel settings). **False** (default) saves files against the language of Visual Basic for Applications (VBA) (which is typically United States English unless the VBA project where Workbooks.Open is run from is an old internationalized XL5/95 VBA project).|
| _CorruptLoad_|Optional| **[XlCorruptLoad](xlcorruptload-enumeration-excel.md)**|Can be one of the following constants:  **xlNormalLoad** , **xlRepairFile** and **xlExtractData** . The default behavior if no value is specified is **xlNormalLoad** and does not attempt recovery when initiated through the OM.|

### Return Value

A  **[Workbook](workbook-object-excel.md)** object that represents the opened workbook.


## Remarks

By default, macros are enabled when opening files programmatically. Use the  **[AutomationSecurity](application-automationsecurity-property-excel.md)** property to set the macro security mode used when opening files programmatically.

You can specify one of the following values in the UpdateLinks parameter to determine whether external references (links) are updated when the workbook is opened.



|**Value**|**Meaning**|
|:-----|:-----|
|0|External references (links) will not be updated when the workbook is opened.|
|3|External references (links) will be updated when the workbook is opened.|
You can specify one of the following values in the Format parameter to determine the delimiter character for the file.



|**Value**|**Delimiter**|
|:-----|:-----|
|1|Tabs|
|2|Commas|
|3|Spaces|
|4|Semicolons|
|5|Nothing|
|6|Custom character (see the  _Delimiter_ argument)|

## Example

The following code example opens the workbook Analysis.xls and then runs its Auto_Open macro.


```vb
Workbooks.Open "ANALYSIS.XLS" 
ActiveWorkbook.RunAutoMacros xlAutoOpen
```



 **Sample code provided by:** Bill Jelen,[MrExcel.com](http://www.mrexcel.com/)

The following code example imports a sheet from another workbook onto a new sheet in the current workbook. Sheet1 in the current workbook must contain the path name of the workbook to import in cell D3, the file name in cell D4, and the worksheet name in cell D5. The imported worksheet is inserted after Sheet1 in the current workbook.




```vb
Sub ImportWorksheet() 
    ' This macro will import a file into this workbook 
    Sheets("Sheet1").Select 
    PathName = Range("D3").Value 
    Filename = Range("D4").Value 
    TabName = Range("D5").Value 
    ControlFile = ActiveWorkbook.Name 
    Workbooks.Open Filename:=PathName &; Filename 
    ActiveSheet.Name = TabName 
    Sheets(TabName).Copy After:=Workbooks(ControlFile).Sheets(1) 
    Windows(Filename).Activate 
    ActiveWorkbook.Close SaveChanges:=False 
    Windows(ControlFile).Activate 
End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

MVP Bill Jelen is the author of more than two dozen books about Microsoft Excel. He is a regular guest on TechTV with Leo Laporte and is the host of MrExcel.com, which includes more than 300,000 questions and answers about Excel. 


## See also
<a name="AboutContributor"> </a>


#### Concepts


[Workbooks Object](workbooks-object-excel.md)

