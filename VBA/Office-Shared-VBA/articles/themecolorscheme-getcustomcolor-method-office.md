---
title: ThemeColorScheme.GetCustomColor Method (Office)
ms.prod: office
api_name:
- Office.ThemeColorScheme.GetCustomColor
ms.assetid: 67ac156e-19ab-245e-b6f8-03514f802acb
ms.date: 06/08/2017
---


# ThemeColorScheme.GetCustomColor Method (Office)

Gets a value that represents a color in the color scheme of a Microsoft Office theme. 


## Syntax

 _expression_. **GetCustomColor**( **_Name_** )

 _expression_ An expression that returns a **ThemeColorScheme** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the custom color.|

### Return Value

MsoRGBType


## Remarks

If the named custom color doesn't exist, an error is generated.


## Example

The following example creates a variable representing the color scheme in an Office theme and then creates another variable containing a custom color. This custom color can then be combined with other colors to define the theme.


```
Dim tTheme As OfficeTheme 
Dim tcsThemeColorScheme As ThemeColorScheme 
Dim csCustomColor As MsoRGBType 
Set tcsThemeColorScheme = tTheme.ThemeColorScheme 
csCustomColor = tcsThemeColorScheme.GetCustomColor("CheerfulColor") 

```


## See also


#### Concepts


[ThemeColorScheme Object](themecolorscheme-object-office.md)
#### Other resources


[ThemeColorScheme Object Members](themecolorscheme-members-office.md)

