---
title: AddIns.Add Method (Excel)
keywords: vbaxl10.chm187073
f1_keywords:
- vbaxl10.chm187073
ms.prod: excel
api_name:
- Excel.AddIns.Add
ms.assetid: 7e4f100d-6ea1-94e4-83d3-fda63a7815e1
ms.date: 06/08/2017
---


# AddIns.Add Method (Excel)

Adds a new add-in file to the list of add-ins. Returns an  **[AddIn](addin-object-excel.md)** object.


## Syntax

 _expression_ . **Add**( **_FileName_** , **_CopyFile_** )

 _expression_ A variable that represents an **AddIns** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Filename_|Required| **String**|The name of the file that contains the add-in or the ProgID of the automation add-in that you want to add to the list in the add-in manager.|
| _CopyFile_|Optional| **Variant**|Ignored if the add-in file is on a hard disk.  **True** to copy the add-in to your hard disk, if the add-in is on a removable medium (a floppy disk or compact disc). **False** to have the add-in remain on the removable medium. If this argument is omitted, Microsoft Excel displays a dialog box and asks you to choose.|

### Return Value

An  **AddIn** object that represents the new add-in.


## Remarks

This method does not install the new add-in. You must set the  **[Installed](addin-installed-property-excel.md)** property to install the add-in.


## Example

This example inserts the add-in Myaddin.xla from drive A. When you run this example, Microsoft Excel copies the file A:\Myaddin.xla to the Library folder on your hard disk and adds the add-in title to the list in the  **Add-Ins** dialog box.


```vb
Sub UseAddIn() 
 
 Set myAddIn = AddIns.Add(Filename:="A:\MYADDIN.XLA", _ 
 CopyFile:=True) 
 MsgBox myAddIn.Title &; " has been added to the list" 
 
End Sub
```


## See also


#### Concepts


[AddIns Collection](addins-object-excel.md)

