---
title: FileConverter Object (Word)
keywords: vbawd10.chm2457
f1_keywords:
- vbawd10.chm2457
ms.prod: word
api_name:
- Word.FileConverter
ms.assetid: 41af2a9b-75cc-253d-4954-4fb42c88530f
ms.date: 06/08/2017
---


# FileConverter Object (Word)

Represents a file converter that's used to open or save files. The  **FileConverter** object is a member of the **FileConverters** collection. The **[FileConverters](fileconverters-object-word.md)** collection contains all the installed file converters for opening and saving files.


## Remarks

Use  **FileConverters** (Index), where Index is a class name or index number, to return a single **FileConverter** object. The following example displays the extensions associated with the Microsoft Excel worksheet converter.


```
MsgBox FileConverters("MSBiff").Extensions
```

The index number represents the position of the file converter in the  **[FileConverters](fileconverters-object-word.md)** collection. The following example displays the format name of the first file converter.




```
MsgBox FileConverters(1).FormatName
```

You cannot create a new file converter or add one to the  **[FileConverters](fileconverters-object-word.md)** collection. **FileConverter** objects are added during installation of Microsoft Office or by installing supplemental file converters. Use either the **CanSave** or **CanOpen** property to determine whether a **FileConverter** object can be used to open or save document.

File converters for saving documents are listed in the  **Save As** dialog box. File converters for opening documents appear in a dialog box if the **Confirm conversion at Open** check box is selected on the **General** tab in the **Options** dialog box ( **Tools** menu).


## Properties



|**Name**|
|:-----|
|[Application](fileconverter-application-property-word.md)|
|[CanOpen](fileconverter-canopen-property-word.md)|
|[CanSave](fileconverter-cansave-property-word.md)|
|[ClassName](fileconverter-classname-property-word.md)|
|[Creator](fileconverter-creator-property-word.md)|
|[Extensions](fileconverter-extensions-property-word.md)|
|[FormatName](fileconverter-formatname-property-word.md)|
|[Name](fileconverter-name-property-word.md)|
|[OpenFormat](fileconverter-openformat-property-word.md)|
|[Parent](fileconverter-parent-property-word.md)|
|[Path](fileconverter-path-property-word.md)|
|[SaveFormat](fileconverter-saveformat-property-word.md)|

## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)
