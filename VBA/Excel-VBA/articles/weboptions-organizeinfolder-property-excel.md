---
title: WebOptions.OrganizeInFolder Property (Excel)
keywords: vbaxl10.chm662074
f1_keywords:
- vbaxl10.chm662074
ms.prod: excel
api_name:
- Excel.WebOptions.OrganizeInFolder
ms.assetid: 9df9aff2-3a24-3e1f-db3e-7280b50b806b
ms.date: 06/08/2017
---


# WebOptions.OrganizeInFolder Property (Excel)

 **True** if all supporting files, such as background textures and graphics, are organized in a separate folder when you save the specified document as a Web page. **False** if supporting files are saved in the same folder as the Web page. The default value is **True** . Read/write **Boolean** .


## Syntax

 _expression_ . **OrganizeInFolder**

 _expression_ A variable that represents a **WebOptions** object.


## Remarks

The new folder is created in the folder where you have saved the Web page, and is named after the document. If long file names are used, a suffix is added to the folder name. The  **[FolderSuffix](weboptions-foldersuffix-property-excel.md)** property returns the folder suffix for the language support you have selected or installed, or the default folder suffix.

If you save a document that was previously saved with the  **OrganizeInFolder** property set to a different value, Microsoft Excel automatically moves the supporting files into or out of the folder, as appropriate.

If you don't use long file names (that is, if the  **[UseLongFileNames](weboptions-uselongfilenames-property-excel.md)** property is set to **False** ), Microsoft Excel automatically saves any supporting files in a separate folder. The files cannot be saved in the same folder as the Web page.


## Example

This example specifies that all supporting files are saved in the same folder when the document is saved as a Web page.


```vb
Application.DefaultWebOptions.OrganizeInFolder = False
```


## See also


#### Concepts


[WebOptions Object](weboptions-object-excel.md)

