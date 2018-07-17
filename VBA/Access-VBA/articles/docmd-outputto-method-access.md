---
title: DoCmd.OutputTo Method (Access)
keywords: vbaac10.chm4165
f1_keywords:
- vbaac10.chm4165
ms.prod: access
api_name:
- Access.DoCmd.OutputTo
ms.assetid: 2a21a7c3-0846-cbec-d5dd-a1648f705557
ms.date: 06/08/2017
---


# DoCmd.OutputTo Method (Access)

The  **OutputTo** method carries out the OutputTo action in Visual Basic.


## Syntax

 _expression_. **OutputTo**( ** _ObjectType_**, ** _ObjectName_**, ** _OutputFormat_**, ** _OutputFile_**, ** _AutoStart_**, ** _TemplateFile_**, ** _Encoding_**, ** _OutputQuality_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ObjectType_|Required|**AcOutputObjectType**|An [AcOutputObjectType](acoutputobjecttype-enumeration-access.md) constant that specifies the type of object to output.|
| _ObjectName_|Optional|**Variant**|A string expression that's the valid name of an object of the type selected by the  _ObjectType_ argument. If you want to output the active object, specify the object's type for the _ObjectType_ argument and leave this argument blank. If you run Visual Basic code containing the **OutputTo** method in a library database, Microsoft Office Access searches for the object with this name, first in the library database, then in the current database.|
| _OutputFormat_|Optional|**Variant**|An  **AcFormat** constant that specifies the output format. If you omit this argument, Access prompts you for the output format.|
| _OutputFile_|Optional|**Variant**|A string expression that's the full name, including the path, of the file you want to output the object to. If you leave this argument blank, Access prompts you for an output file name.|
| _AutoStart_|Optional|**Variant**|Use  **True** (?1) to start the appropriate Microsoft Windows?based application immediately, with the file specified by the _OutputFile_ argument loaded. Use **False** (0) if you don't want to start the application. This argument is ignored for Microsoft Internet Information Server (.htx, .idc) files and Microsoft ActiveX Server (*.asp) files. If you leave this argument blank, the default ( **False** ) is assumed.|
| _TemplateFile_|Optional|**Variant**|A string expression that's the full name, including the path, of the file you want to use as a template for an HTML, HTX, or ASP file.|
| _Encoding_|Optional|**Variant**|The type of character encoding format you want used to output the text or HTML data. You can select MS-DOS, Unicode, or Unicode (UTF-8). The MS-DOS argument setting is available only for text files. If you leave this argument blank, Access outputs the data by using the Windows default encoding for text files and the default system encoding for HTML files.|
| _OutputQuality_|Optional|**AcExportQuality**|An  **[AcExportQuality](acexportquality-enumeration-access.md)** constant that specifies the type of output device to optimize for. The default value is **acExportQualityPrint**.|

## Remarks

You can use the  **OutputTo** method to output the data in the specified Access database object (a datasheet, form, report, module, data access page) to several output formats.

Modules can be output only in MS-DOS Text format, so if you specify  **acOutputModule** for the _ObjectType_ argument, you must specify **acFormatTXT** for the _OutputFormat_ argument. Microsoft Internet Information Server and Microsoft ActiveX Server formats are available only for tables, queries, and forms, so if you specify **acFormatIIS** or **acFormatASP** for the _OutputFormat_ argument, you must specify **acOutputTable**, **acOutputQuery**, or **acOutputForm** for the _ObjectType_ argument.

The Access data is output in the selected format and can be read by any application that uses the same format. For example, you can output an Access report with its formatting to a rich-text format document and then open the document in Microsoft Word.


 **Note**  You can save as a PDF or XPS file from a 2007 Microsoft Office system program only after you install an add-in. For more information, search for "Enable support for other file formats, such as PDF and XPS" on the Office Web site.


## Example

The following code example outputs the Employees table in rich-text format (.rtf) to the Employee.rtf file and immediately opens the file in Microsoft Word for Windows.


```vb
DoCmd.OutputTo acOutputTable, "Employees", _ 
 acFormatRTF, "Employee.rtf", True
```


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

