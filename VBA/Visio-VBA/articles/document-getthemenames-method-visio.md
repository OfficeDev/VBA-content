---
title: Document.GetThemeNames Method (Visio)
keywords: vis_sdr.chm10560075
f1_keywords:
- vis_sdr.chm10560075
ms.prod: visio
api_name:
- Visio.Document.GetThemeNames
ms.assetid: 63477332-5db2-40ff-6918-7ab20a9f0fd0
ms.date: 06/08/2017
---


# Document.GetThemeNames Method (Visio)

Returns a locale-specific array of names of themes contained in the document.


## Syntax

 _expression_ . **GetThemeNames**( **_eType_** , **_NameArray()_** )

 _expression_ An expression that returns a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _eType_|Required| **VisThemeTypes**|The type of the theme, an enumerated value from the  **VisThemeTypes** enumeration. See Remarks for possible values.|
| _NameArray()_|Required| **String**|Out parameter. An array of locale-specific theme names returned by the method.|

### Return Value

Nothing


## Remarks

For the  _eType_ parameter, pass a value from the **VisThemeTypes** enumeration, which is declared in the Visio type library.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visThemeTypeColor**|1|Color themes.|
| **visThemeTypeEffect**|2|Effect themes.|
For the  _NameArray()_ out parameter, pass an empty, dimensionless array of type **String** . Visio returns the array filled with locale-specific names of themes contained in the document.

To get locale-independent themes in the document, use the  **[Document.GetThemeNamesU](document-getthemenamesu-method-visio.md)** method.




 **Note**   Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, themes, and layers. When a user names a shape, for example, the user is specifying a local name.Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions of Visio, universal names were not visible in the user interface.) As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. 


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **GetThemeNames** method to get the list of locale-specific theme color and theme effect names in the active document. It prints the list in the **Immediate** window.


```vb
Public Sub GetThemeNames_Example() 
 
    Dim astrNames() As String 
    Dim strThemeName As String 
    Dim intArrayCounter As Integer 
     
    ActiveDocument.GetThemeNames visThemeTypeColor, astrNames 
     
    For intArrayCounter = LBound(astrNames) To UBound(astrNames) 
        strThemeName = astrNames(intArrayCounter) 
        Debug.Print strThemeName 
    Next 
     
    Debug.Print "-------------------------------------------" 
     
    ActiveDocument.GetThemeNames visThemeTypeEffect, astrNames 
     
    For intArrayCounter = LBound(astrNames) To UBound(astrNames) 
        strThemeName = astrNames(intArrayCounter) 
        Debug.Print strThemeName 
    Next 
     
End Sub
```


