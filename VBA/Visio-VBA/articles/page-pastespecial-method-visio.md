---
title: Page.PasteSpecial Method (Visio)
keywords: vis_sdr.chm10951020
f1_keywords:
- vis_sdr.chm10951020
ms.prod: visio
api_name:
- Visio.Page.PasteSpecial
ms.assetid: 949a507a-1cc2-0b52-b0dd-3ad40ac9ecdf
ms.date: 06/08/2017
---


# Page.PasteSpecial Method (Visio)

Inserts the contents of the Clipboard, allowing you to control the format of the pasted information and (optionally) establish a link to the source file (for example, a Microsoft Word document).


## Syntax

 _expression_ . **PasteSpecial**( **_Format_** , **_Link_** , **_DisplayAsIcon_** )

 _expression_ A variable that represents a **Page** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Format_|Required| **Long**|The internal Clipboard format.|
| _Link_|Optional| **Variant**| **True** to establish a link to the source of the pasted data; otherwise, **False** (the default). Ignored if the source data is not suitable for, or does not support, linking.|
| _DisplayAsIcon_|Optional| **Variant**| **True** to display the pasted data as an icon; otherwise, **False** (the default).|

### Return Value

Nothing


## Remarks

To simply paste the contents of the Clipboard into an object, use the  **Paste** method.

The value of the  _Format_ argument can be any of the following:




- A value from  **VisPasteSpecialCodes** (see the following table).
    
- Any of the standard Clipboard formats, for example, CF_TEXT. For details, see the Microsoft Platform SDK on MSDN, the Microsoft Developer Network Web site.
    
- Any value returned from a call to the  **RegisterClipboardFormat** function. For details, see the Microsoft Platform SDK on MSDN.
    





 **Note**  Before calling Microsoft Windows functions, you should understand how arguments and data types are handled by the Windows API DLLs. Incorrectly calling Windows functions may result in invalid page faults or other unexpected behaviors. For more information on calling Windows functions, search for "Windows API" on MSDN.

Possible values for  _Format_ declared by the Visio type library in **VisPasteSpecialCodes** are described in the following table.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visPasteBitmap**|2|Paste bitmap.|
| **visPasteDIB**|8|Paste device-independent bitmap.|
| **visPasteEMF**|14|Paste enhanced metafile.|
| **visPasteHyperlink**|65538|Paste hyperlink.|
| **visPasteInk**|65544|Paste Ink data.|
| **visPasteMetafile**|3|Paste metafile.|
| **visPasteOEMText**|7|Paste OEM text.|
| **visPasteOLEObject**|65536|Paste OLE object.|
| **visPasteRichText**|65537|Paste rich text.|
| **visPasteText**|1|Paste ANSI text.|
| **visPasteURL**|65539|Paste Uniform Resource Locator (URL).|
| **visPasteVisioIcon**|65543|Paste Visio icon.|
| **visPasteVisioMastersXML**|65546|Paste Visio masters XML.|
| **visPasteVisioMasters**|65541|Paste Visio masters.|
| **visPasteVisioShapesXML**|65545|Paste Visio shapes XML.|
| **visPasteVisioShapesWithoutDataLinks**|65548|Paste Visio drawing data without internal data links.|
| **visPasteVisioShapes**|65540|Paste Visio shapes.|
| **visPasteVisioText**|65542|Paste Visio text.|

