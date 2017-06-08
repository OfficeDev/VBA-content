---
title: Application.DocumentExport Method (Project)
keywords: vbapj.chm2173
f1_keywords:
- vbapj.chm2173
ms.prod: project-server
api_name:
- Project.Application.DocumentExport
ms.assetid: 891bf868-1256-2688-cdb2-2bccfbf2afc2
ms.date: 06/08/2017
---


# Application.DocumentExport Method (Project)

Exports the active project as a document in PDF or XPS format.


## Syntax

 _expression_. **DocumentExport**( ** _Filename_**, ** _FileType_**, ** _IncludeDocumentProperties_**, ** _IncludeDocumentMarkup_**, ** _ArchiveFormat_**, ** _FromDate_**, ** _ToDate_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filename_|Optional|**String**|Specifies the file name of the exported document. The default value is the name of the active project as a PDF file.|
| _FileType_|Optional|**[PjDocExportType](pjdocexporttype-enumeration-project.md)**|Specifies whether to export the project as a PDF or an XPS document. The default value is  **pjPDF** (0).|
| _IncludeDocumentProperties_|Optional|**Variant**|If  **True** or 1, the last page of the exported document includes some document properties. The default value is **True**.|
| _IncludeDocumentMarkup_|Optional|**Variant**|If  **True** or 1, the last page of the exported document includes a legend of the symbols shown in the view. The default is **True**.|
| _ArchiveFormat_|Optional|**Variant**|If  **True** or 1, exports a PDF document in the ISO 19500-1 compliant (PDF/A) format. The default value is **False**.|
| _FromDate_|Optional|**Variant**|The start date of the range of dates to publish. The default value is the project start date.|
| _ToDate_|Optional|**Variant**|The end date of the range of dates to publish. The default value is the project end date.|

### Return Value

 **Boolean**


## Remarks

Running the  **DocumentExport** method without any parameters brings up the **Browse** dialog box and the name of the active project as a PDF file. If the user cancels the **Browse** or the subsequent **Document Export Options** dialog box, **DocumentExport** returns **False**.

To export a custom format PDF or XPS document, where you can use a pointer to a class in an add-in, see  **[ExportAsFixedFormat](project-exportasfixedformat-method-project.md)**.


## Example

If the active project shows a Network Diagram view, the following example creates an XPS document named TestProject.xps. When you open the file in the  **XPS Viewer** application, the last page includes document properties and a legend that shows the PERT chart symbols.


```
DocumentExport FileName:="TestProject.xps", FileType:=pjXPS
```


