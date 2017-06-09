---
title: AddIns.Add Method (Word)
keywords: vbawd10.chm159318018
f1_keywords:
- vbawd10.chm159318018
ms.prod: word
api_name:
- Word.AddIns.Add
ms.assetid: 09a7ba59-94a6-f6b0-a012-7d5aaa5b5b12
ms.date: 06/08/2017
---


# AddIns.Add Method (Word)

Returns an  **[AddIn](addin-object-word.md)** object that represents an add-in added to the list of available add-ins.


## Syntax

 _expression_ . **Add**( **_FileName_** , **_Install_** )

 _expression_ Required. A variable that represents an **[AddIns](addins-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The path for the template or WLL.|
| _Install_|Optional| **Variant**| **True** to install the add-in. **False** to add the add-in to the list of add-ins but not install it. The default value is **True** .|

## Remarks

Use the  **[Installed](addin-installed-property-word.md)** property of an add-in to see whether it is already installed.


## Example

This example installs a template named MyFax.dot and adds it to the list of add-ins in the  **Templates and Add-ins** dialog box.


```vb
Sub AddTemplate() 
 ' For this example to work correctly, verify that the 
 ' path is correct and the file exists. 
 
 AddIns.Add FileName:="C:\Program Files\Microsoft Office" _ 
 &; "\Templates\Letters &; Faxes\MyFax.dot", Install:=True 
End Sub
```


## See also


#### Concepts


[AddIns Collection Object](addins-object-word.md)

