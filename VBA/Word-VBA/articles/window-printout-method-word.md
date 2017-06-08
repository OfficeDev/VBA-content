---
title: Window.PrintOut Method (Word)
keywords: vbawd10.chm157417917
f1_keywords:
- vbawd10.chm157417917
ms.prod: word
api_name:
- Word.Window.PrintOut
ms.assetid: 63ea2dd2-5b3c-1239-16ce-1b4980cde3d3
ms.date: 06/08/2017
---


# Window.PrintOut Method (Word)

Prints all or part of the document displayed in the specified window.


## Syntax

 _expression_ . **PrintOut**( **_Background_** , **_Append_** , **_Range_** , **_OutputFileName_** , **_From_** , **_To_** , **_Item_** , **_Copies_** , **_Pages_** , **_PageType_** , **_PrintToFile_** , **_Collate_** , **_FileName_** , **_ActivePrinterMacGX_** , **_ManualDuplexPrint_** , **_PrintZoomColumn_** , **_PrintZoomRow_** , **_PrintZoomPaperWidth_** , **_PrintZoomPaperHeight_** )

 _expression_ Required. A variable that represents a **[Window](window-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Background_|Optional| **Variant**|Set to  **True** to have the macro continue while Microsoft Word prints the document.|
| _Append_|Optional| **Variant**|Set to  **True** to append the specified document to the file name specified by the OutputFileName argument. **False** to overwrite the contents of OutputFileName.|
| _Range_|Optional| **Variant**|The page range. Can be any  **WdPrintOutRange** constant.|
| _OutputFileName_|Optional| **Variant**|If PrintToFile is  **True** , this argument specifies the path and file name of the output file.|
| _From_|Optional| **Variant**|The starting page number when Range is set to  **wdPrintFromTo** .|
| _To_|Optional| **Variant**|The ending page number when Range is set to  **wdPrintFromTo** .|
| _Item_|Optional| **Variant**|The item to be printed. Can be any  **WdPrintOutItem** constant.|
| _Copies_|Optional| **Variant**|The number of copies to be printed.|
| _Pages_|Optional| **Variant**|The page numbers and page ranges to be printed, separated by commas. For example, "2, 6-10" prints page 2 and pages 6 through 10.|
| _PageType_|Optional| **Variant**|The type of pages to be printed. Can be any  **WdPrintOutPages** constant.|
| _PrintToFile_|Optional| **Variant**| **True** to send printer instructions to a file. Make sure to specify a file name with OutputFileName.|
| _Collate_|Optional| **Variant**|When printing multiple copies of a document,  **True** to print all pages of the document before printing the next copy.|
| _FileName_|Optional| **Variant**|The path and file name of the document to be printed. If this argument is omitted, Word prints the active document. (Available only with the  **Application** object.)|
| _ActivePrinterMacGX_|Optional| **Variant**|This argument is available only in Microsoft Office Macintosh Edition. For additional information about this argument, consult the language reference Help included with Microsoft Office Macintosh Edition.|
| _ManualDuplexPrint_|Optional| **Variant**| **True** to print a two-sided document on a printer without a duplex printing kit. If this argument is **True** , the **PrintBackground** and **PrintReverse** properties are ignored. Use the **PrintOddPagesInAscendingOrder** and **PrintEvenPagesInAscendingOrder** properties to control the output during manual duplex printing. This argument may not be available to you, depending on the language support (U.S. English, for example) that you?ve selected or installed.|
| _PrintZoomColumn_|Optional| **Variant**|The number of pages you want Word to fit horizontally on one page. Can be 1, 2, 3, or 4. Use with the PrintZoomRow argument to print multiple pages on a single sheet.|
| _PrintZoomRow_|Optional| **Variant**|The number of pages you want Word to fit vertically on one page. Can be 1, 2, or 4. Use with the PrintZoomColumn argument to print multiple pages on a single sheet.|
| _PrintZoomPaperWidth_|Optional| **Variant**|The width to which you want Word to scale printed pages, in twips (20 twips = 1 point; 72 points = 1 inch).|
| _PrintZoomPaperHeight_|Optional| **Variant**|The height to which you want Word to scale printed pages, in twips (20 twips = 1 point; 72 points = 1 inch).|

## Example

This example prints the current page of the active document.


```vb
ActiveDocument.PrintOut Range:=wdPrintCurrentPage
```

This example prints all the documents in the current folder. The  **Dir** function is used to return all file names that have the file name extension ".doc".




```vb
adoc = Dir("*.DOC") 
Do While adoc <> "" 
 Application.PrintOut FileName:=adoc 
 adoc = Dir() 
Loop
```

This example prints the first three pages of the document in the active window.




```vb
ActiveDocument.ActiveWindow.PrintOut _ 
 Range:=wdPrintFromTo, From:="1", To:="3"
```

This example prints the comments in the active document.




```vb
If ActiveDocument.Comments.Count >= 1 Then 
 ActiveDocument.PrintOut Item:=wdPrintComments 
End If
```

This example prints the active document, fitting six pages on each sheet.




```vb
ActiveDocument.PrintOut PrintZoomColumn:=3, _ 
 PrintZoomRow:=2
```

This example prints the active document at 75% of actual size.




```vb
ActiveDocument.PrintOut _ 
 PrintZoomPaperWidth:=0.75 * (8.5 * 1440), _ 
 PrintZoomPaperHeight:=0.75 * (11 * 1440)
```


## See also


#### Concepts


[Window Object](window-object-word.md)

