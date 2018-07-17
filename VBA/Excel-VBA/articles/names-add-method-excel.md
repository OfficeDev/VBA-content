---
title: Names.Add Method (Excel)
keywords: vbaxl10.chm488073
f1_keywords:
- vbaxl10.chm488073
ms.prod: excel
api_name:
- Excel.Names.Add
ms.assetid: 89a888bc-20b1-dd63-ede9-b3ba1d5ffab0
ms.date: 06/08/2017
---


# Names.Add Method (Excel)

Defines a new name for a range of cells.


## Syntax

 _expression_ . **Add**( **_Name_** , **_RefersTo_** , **_Visible_** , **_MacroType_** , **_ShortcutKey_** , **_Category_** , **_NameLocal_** , **_RefersToLocal_** , **_CategoryLocal_** , **_RefersToR1C1_** , **_RefersToR1C1Local_** )

 _expression_ A variable that represents a **Names** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Optional| **Variant**|Specifies the text, in English, to use as the name if the NameLocal parameter is not specified. Names cannot include spaces and cannot be formatted as cell references.|
| _RefersTo_|Optional| **Variant**|Describes what the name refers to, in English, using A1-style notation, if the RefersToLocal, RefersToR1C1, and RefersToR1C1Local parameters are not specified.<table><tr><th>**Note**</th></tr><tr><td>**Nothing** is returned if the reference does not exist.</td></tr></table>|
| _Visible_|Optional| **Variant**| **True** specifies that the name is defined as visible. **False** specifies that the name is defined as hidden. A hidden name does not appear in the **Define Name**,  **Paste Name**, or  **Goto** dialog box. The default value is **True** .|
| _MacroType_|Optional| **Variant**|The macro type, determined by one of the following values:<ul><li><p>1 - User-defined function (<b>Function</b>  procedure)</p></li><li><p>2 - Macro (<b>Sub</b>  procedure)</p></li><li><p>3 or omitted - None (the name does not  refer to a user-defined function or macro)</p></li></ul>|
| _ShortcutKey_|Optional| **Variant**|Specifies the macro shortcut key. Must be a single letter, such as "z" or "Z". Applies only for command macros.|
| _Category_|Optional| **Variant**|The category of the macro or function if the MacroType argument equals 1 or 2. The category is used in the Function Wizard. Existing categories can be referred to either by number, starting at 1, or by name, in English. Microsoft Office Excel 2007 creates a new category if the specified category does not exist.|
| _NameLocal_|Optional| **Variant**|Specifies the localized text to use as the name if the Name parameter is not specified. Names cannot include spaces and cannot be formatted as cell references.|
| _RefersToLocal_|Optional| **Variant**|Describes what the name refers to, in localized text using A1-style notation, if the RefersTo, RefersToR1C1, and RefersToR1C1Local parameters are not specified.|
| _CategoryLocal_|Optional| **Variant**|Specifies the localized text that identifies the category of a custom function if the Category parameter is not specified.|
| _RefersToR1C1_|Optional| **Variant**|Describes what the name refers to, in English using R1C1-style notation, if the RefersTo, RefersToLocal, and RefersToR1C1Local parameters are not specified.|
| _RefersToR1C1Local_|Optional| **Variant**|Describes what the name refers to, in localized text using R1C1-style notation, if the RefersTo, RefersToLocal, and RefersToR1C1 parameters are not specified.|

### Return Value

A  **[Name](name-object-excel.md)** object that represents the new name.


## Example

This example defines a new name for the range A1:D3 on Sheet1 in the active workbook. 

|**Note**|
|:-----|  
|**Nothing** is returned if Sheet1 does not exist.|
  

```vb
Sub MakeRange() 
 
    ActiveWorkbook.Names.Add _ 
        Name:="tempRange", _ 
        RefersTo:="=Sheet1!$A$1:$D$3" 
 
End Sub
```


## See also


#### Concepts


[Names Object](names-object-excel.md)

