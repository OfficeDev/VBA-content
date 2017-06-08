---
title: Application.GetOpenFilename Method (Excel)
keywords: vbaxl10.chm133142
f1_keywords:
- vbaxl10.chm133142
ms.prod: excel
api_name:
- Excel.Application.GetOpenFilename
ms.assetid: 83931dc2-59b3-550b-6ce1-880413fd23d6
ms.date: 06/08/2017
---


# Application.GetOpenFilename Method (Excel)

Displays the standard  **Open** dialog box and gets a file name from the user without actually opening any files.


## Syntax

 _expression_ . **GetOpenFilename**( **_FileFilter_** , **_FilterIndex_** , **_Title_** , **_ButtonText_** , **_MultiSelect_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileFilter_|Optional| **Variant**|A string specifying file filtering criteria.|
| _FilterIndex_|Optional| **Variant**|Specifies the index numbers of the default file filtering criteria, from 1 to the number of filters specified in  _FileFilter_. If this argument is omitted or greater than the number of filters present, the first file filter is used.|
| _Title_|Optional| **Variant**|Specifies the title of the dialog box. If this argument is omitted, the title is "Open."|
| _ButtonText_|Optional| **Variant**|Macintosh only.|
| _MultiSelect_|Optional| **Variant**| **True** to allow multiple file names to be selected. **False** to allow only one file name to be selected. The default value is **False** .|

### Return Value

Variant


## Remarks

This string passed in the  _FileFilter_ argument consists of pairs of file filter strings followed by the MS-DOS wildcard file filter specification, with each part and each pair separated by commas. Each separate pair is listed in the **Files of type** drop-down list box. For example, the following string specifies two file filters?text and addin: "Text Files (*.txt),*.txt,Add-In Files (*.xla),*.xla".

To use multiple MS-DOS wildcard expressions for a single file filter type, separate the wildcard expressions with semicolons; for example, "Visual Basic Files (*.bas; *.txt),*.bas;*.txt".

If  _FileFilter_ is omitted, this argument defaults to "All Files (*.*),*.*".

This method returns the selected file name or the name entered by the user. The returned name may include a path specification. If  _MultiSelect_ is **True** , the return value is an array of the selected file names (even if only one filename is selected). Returns **False** if the user cancels the dialog box.

This method may change the current drive or folder.


## Example

This example displays the  **Open** dialog box, with the file filter set to text files. If the user chooses a file name, the code displays that file name in a message box.


```vb
fileToOpen = Application _ 
 .GetOpenFilename("Text Files (*.txt), *.txt") 
If fileToOpen <> False Then 
 MsgBox "Open " &; fileToOpen 
End If
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

