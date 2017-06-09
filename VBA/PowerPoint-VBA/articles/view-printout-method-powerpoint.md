---
title: View.PrintOut Method (PowerPoint)
keywords: vbapp10.chm512012
f1_keywords:
- vbapp10.chm512012
ms.prod: powerpoint
api_name:
- PowerPoint.View.PrintOut
ms.assetid: 244da3c5-ddb2-f79c-b8fc-cad4a293defe
ms.date: 06/08/2017
---


# View.PrintOut Method (PowerPoint)

Prints the specified presentation.


## Syntax

 _expression_. **PrintOut**( **_From_**, **_To_**, **_PrintToFile_**, **_Copies_**, **_Collate_** )

 _expression_ A variable that represents a **View** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _From_|Optional|**Long**|The number of the first page to be printed. If this argument is omitted, printing starts at the beginning of the presentation. Specifying the  **To** and **From** arguments sets the contents of the **[PrintRanges](printranges-object-powerpoint.md)** object and sets the value of the **RangeType** property for the presentation.|
| _To_|Optional|**Long**|The number of the last page to be printed. If this argument is omitted, printing continues to the end of the presentation. Specifying the  **To** and **From** arguments sets the contents of the **[PrintRanges](printranges-object-powerpoint.md)** object and sets the value of the **RangeType** property for the presentation.|
| _PrintToFile_|Optional|**String**|The name of the file to print to. If you specify this argument, the file is printed to a file rather than sent to a printer. If this argument is omitted, the file is sent to a printer.|
| _Copies_|Optional|**Long**|The number of copies to be printed. If this argument is omitted, only one copy is printed. Specifying this argument sets the value of the  **[NumberOfCopies](printoptions-numberofcopies-property-powerpoint.md)** property.|
| _Collate_|Optional|**MsoTriState**|If this argument is omitted, multiple copies are collated. Specifying this argument sets the value of the  **[Collate](printoptions-collate-property-powerpoint.md)** property.|

## Remarks

The  _Collate_ parameter value can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|Prints all copies of one page before printing the first copy of the next page.|
|**msoTrue**|Prints a complete copy of the presentation before the first page of the next copy is printed.|

## See also


#### Concepts


[View Object](view-object-powerpoint.md)

