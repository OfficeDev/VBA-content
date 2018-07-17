---
title: Application.MacroOptions Method (Excel)
keywords: vbaxl10.chm133324
f1_keywords:
- vbaxl10.chm133324
ms.prod: excel
api_name:
- Excel.Application.MacroOptions
ms.assetid: c81abbc5-0865-9e86-f188-652c88ac6baa
ms.date: 06/08/2017
---


# Application.MacroOptions Method (Excel)

Corresponds to options in the  **Macro Options** dialog box. You can also use this method to display a user defined function (UDF) in a built-in or new category within the **Insert Function** dialog box.


## Syntax

 _expression_ . **MacroOptions**( **_Macro_** , **_Description_** , **_HasMenu_** , **_MenuText_** , **_HasShortcutKey_** , **_ShortcutKey_** , **_Category_** , **_StatusBar_** , **_HelpContextID_** , **_HelpFile_** , **_ArgumentDescriptions_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Macro_|Optional| **Variant**|The macro name or the name of a user defined function (UDF).|
| _Description_|Optional| **Variant**|The macro description.|
| _HasMenu_|Optional| **Variant**|This argument is ignored.|
| _MenuText_|Optional| **Variant**|This argument is ignored.|
| _HasShortcutKey_|Optional| **Variant**| **True** to assign a shortcut key to the macro ( _ShortcutKey_ must also be specified). If this argument is **False** , no shortcut key is assigned to the macro. If the macro already has a shortcut key, setting this argument to **False** removes the shortcut key. The default value is **False** .|
| _ShortcutKey_|Optional| **Variant**|Required if  _HasShortcutKey_ is **True** ; ignored otherwise. The shortcut key.|
| _Category_|Optional| **Variant**|An integer that specifies an existing macro function category (Financial, Date &; Time, or User Defined, for example). See the Remarks section to determine the integers that are mapped to the built-in categories. You can also specify a string for a custom category. If you provide a string it will be treated as the category name that is displayed in the  **Insert Function** dialog box. If the category name has never been used, a new category is defined with that name. If you use a category name that is the same as a built-in name (see list in Remarks section), Microsoft Excel will map the user defined function to that built-in category.|
| _StatusBar_|Optional| **Variant**|The status bar text for the macro.|
| _HelpContextID_|Optional| **Variant**|An integer that specifies the context ID for the Help topic assigned to the macro.|
| _HelpFile_|Optional| **Variant**|The name of the Help file that contains the Help topic defined by  _HelpContextId_.|
| _ArgumentDescriptions_|Optional| **Array**|A one-dimensional array that contains the descriptions for the arguments to a UDF that are displayed in the  **Function Arguments** dialog box.|

## Remarks

The following table lists which integers are mapped to the built-in categories that can be used in the  **_Category_** parameter.



| **Integer**| **Category**|
|1| **Financial**|
|2| **Date &; Time**|
|3| **Math &; Trig**|
|4| **Statistical**|
|5| **Lookup &; Reference**|
|6| **Database**|
|7| **Text**|
|8| **Logical**|
|9| **Information**|
|10| **Commands**|
|11| **Customizing**|
|12| **Macro Control**|
|13| **DDE/External**|
|14| **User Defined**|
|15|First custom category|
|16|Second custom category|
|17|Third custom category|
|18|Fourth custom category|
|19|Fifth custom category|
|20|Sixth custom category|
|21|Seventh custom category|
|22|Eighth custom category|
|23|Ninth custom category|
|24|Tenth custom category|
|25|Eleventh custom category|
|26|Twelfth custom category|
|27|Thirteenth custom category|
|28|Fourteenth custom category|
|29|Fifteenth custom category|
|30|Sixteenth custom category|
|31|Seventeenth custom category|
|32|Eighteenth custom category|

## Example

This example adds a user-defined macro called "TestMacro" to a custom category named "My Custom Category". After you run this example, you should see "My Custom Category" which contains the "TestMacro" user-defined function in the  **Or select a category** drop-down list in the **Insert Function** dialog box.


```vb
Function TestMacro() 
    MsgBox ActiveWorkbook.Name 
End Function 
 
Sub AddUDFToCustomCategory() 
    Application.MacroOptions Macro:="TestMacro", Category:="My Custom Category" 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

