---
title: Application.FilePrint Method (Project)
keywords: vbapj.chm109
f1_keywords:
- vbapj.chm109
ms.prod: project-server
api_name:
- Project.Application.FilePrint
ms.assetid: 47937a14-3c57-a597-0b67-5c095bda8ec7
ms.date: 06/08/2017
---


# Application.FilePrint Method (Project)

Prints the active view.


## Syntax

 _expression_. **FilePrint**( ** _FromPage_**, ** _ToPage_**, ** _PageBreaks_**, ** _Draft_**, ** _Copies_**, ** _FromDate_**, ** _ToDate_**, ** _OnePageWide_**, ** _Preview_**, ** _Color_**, ** _ShowIEPrintDialog_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FromPage_|Optional|**Integer**|A number that specifies the first page to print. The default value is 1.|
| _ToPage_|Optional|**Integer**|A number that specifies the last page to print. The default is the last page in the project.|
| _PageBreaks_|Optional|**Boolean**|**True** if Project uses manual page breaks when printing. The default value is **True**.|
| _Draft_|Optional|**Boolean**|**True** if Project prints the active view in draft mode. The default value is **False**.|
| _Copies_|Optional|**Integer**|A number that specifies the number of copies to print. The default value is 1.|
| _FromDate_|Optional|**Variant**|A number or string that specifies the first date to print. The default is the start date of the project.|
| _ToDate_|Optional|**Variant**|A number or string that specifies the last date to print. The default is the finish date of the project.|
| _OnePageWide_|Optional|**Boolean**|**True** if Project prints only the leftmost columns of the active view. The default value is **False**.|
| _Preview_|Optional|**Boolean**|**True** if Project previews the active view rather than printing it. The default value is **False**.|
| _Color_|Optional|**Boolean**|**True** if Project prints the active view in color. The default value is **False**.|
| _ShowIEPrintDialog_|Optional|**Boolean**|If  **True**, shows the Internet Explorer print dialog while printing.|

### Return Value

 **Boolean**


## Remarks

 **FilePrint** with no arguments acts the same as the **FilePrintPreview** method. It opens the Backstage view and displays the **Print** tab with a print preview.


## Example

The following example prints the active view without using manual page breaks.


```vb
Sub PrintViewWithoutPageBreaks() 
    FilePrint PageBreaks:=False 
End Sub
```

The following command prints the active view to the default printer, and shows the Internet Explorer print dialog.




```vb
Application.FilePrint ShowIEPrintDialog:=True
```


