---
title: Document.PrintOut Method (Visio)
keywords: vis_sdr.chm10551340
f1_keywords:
- vis_sdr.chm10551340
ms.prod: visio
api_name:
- Visio.Document.PrintOut
ms.assetid: c13f7640-7439-0c73-cde5-223b8b4549d3
ms.date: 06/08/2017
---


# Document.PrintOut Method (Visio)

Prints the contents of the active document and provides various printing options.


## Syntax

 _expression_ . **PrintOut**( **_PrintRange_** , **_FromPage_** , **_ToPage_** , **_ScaleCurrentViewToPaper_** , **_PrinterName_** , **_PrintToFile_** , **_OutputFileName_** , **_Copies_** , **_Collate_** , **_ColorAsBlack_** )

 _expression_ A variable that represents a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _PrintRange_|Required| **VisPrintOutRange**|The range of document pages to be printed. See Remarks for possible values.|
| _FromPage_|Optional| **Long**| If PrintRange is **visPrintFromTo** , the first page in the range to be printed. The default is 1, which indicates the first page of the drawing.|
| _ToPage_|Optional| **Long**|If PrintRange is  **visPrintFromTo** , the last page in the range to be printed. The default is -1, which indicates the last page of the drawing.|
| _ScaleCurrentViewToPaper_|Optional| **Boolean**|If PrintRange is  **visPrintCurrentView** , **True** to scale the part of the drawing that fits in the program window to fit on the current default paper size; **False** to print on as many printer pages as necessary. The default is **False** .|
| _PrinterName_|Optional| **String**|The name of the printer to use. The default is the printer currently selected in Visio.|
| _PrintToFile_|Optional| **Boolean**| **True** to send the information for printing to a file on a disk, rather than to the printer; **False** to print to the printer. The default is **False** . If you specify **True** for PrintToFile but do not pass a valid value for OutputFileName, the drawing is sent to the active printer.|
| _OutputFileName_|Optional| **String**|If PrintToFile is  **True** , the name and path of the .prn file to which to print, enclosed in quotation marks.|
| _Copies_|Optional| **Long**|The number of copies to be printed. The default is one copy.|
| _Collate_|Optional| **Boolean**| **True** to collate copies. **False** to not collate. The default is **False** .|
| _ColorAsBlack_|Optional| **Boolean**| **True** to print all colors as black to ensure that all shapes are visible in the printed drawing. This is useful if a monochrome printer translates very light colors in a drawing to white rather than a shade of gray. **False** to print colors normally. The default is **False** .|

### Return Value

Nothing


## Remarks

Calling the  **PrintOut** method is the equivalent of selecting print options in the **Print** dialog box (click the **File** tab, click **Print**, and then click  **Print** again), and then clicking **OK**.

Possible values for  _PrintRange_ are shown in the following table and declared in **VisPrintOutRange** in the Visio type library.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visPrintAll**|0|Prints all foreground pages.|
| **visPrintCurrentPage**|2|Prints the active page.|
| **visPrintCurrentView**|4|Prints the current view area.|
| **visPrintFromTo**|1|Prints pages between the FromPage value and the ToPage value.|
| **visPrintSelection**|3|Prints a selection.|

## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **PrintOut** method to print two copies of the current page to the active printer.


```vb
Public Sub PrintOut_Example() 
 
    'Print two copies of the current page to the default printer 
    ThisDocument.PrintOut visPrintCurrentPage, , , , , , , 2 
 
End Sub
```


