---
title: Document.GetThemeNamesU Method (Visio)
keywords: vis_sdr.chm10560080
f1_keywords:
- vis_sdr.chm10560080
ms.prod: visio
api_name:
- Visio.Document.GetThemeNamesU
ms.assetid: 7a7280ae-10c9-9bc7-c121-29791e4df557
ms.date: 06/08/2017
---


# Document.GetThemeNamesU Method (Visio)

Returns a locale-independent array of names of themes contained in the document.


## Syntax

 _expression_ . **GetThemeNamesU**( **_eType_** , **_NameArray()_** )

 _expression_ An expression that returns a **Document** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _eType_|Required| **VisThemeTypes**|The type of the theme, an enumerated value from the  **VisThemeTypes** enumeration. See Remarks for possible values.|
| _NameArray()_|Required| **String**|Out parameter. An array of locale-independent theme names returned by the method.|

### Return Value

Nothing


## Remarks

For the  _eType_ parameter, pass a value from the **VisThemeTypes** enumeration, which is declared in the Visio type library.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visThemeTypeColor**|1|Color themes.|
| **visThemeTypeEffect**|2|Effect themes.|
For the  _NameArray()_ out parameter, pass an empty, dimensionless array of type **String** . Visio returns the array filled with locale-independent names of themes contained in the document.

To get the names of locale-specific themes in the document, use the  **[Document.GetThemeNames](document-getthemenames-method-visio.md)** method.




 **Note**   Beginning with Microsoft Visio 2000, you can use both local and universal names to refer to Visio shapes, masters, documents, pages, rows, add-ons, cells, hyperlinks, styles, fonts, master shortcuts, UI objects, themes, and layers. When a user names a shape, for example, the user is specifying a local name.Beginning with Microsoft Office Visio 2003, the ShapeSheet spreadsheet displays only universal names in cell formulas and values. (In prior versions of Visio, universal names were not visible in the user interface.) As a developer, you can use universal names in a program when you don't want to change a name each time a solution is localized. 


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **GetThemeNamesU** method to get the list of locale-independent theme color and theme effect names in the active document. It prints the list in the **Immediate** window.


```vb
Public Sub GetThemeNamesU_Example() 
 
    Dim astrNames() As String 
    Dim strThemeName As String 
    Dim intArrayCounter As Integer 
     
    ActiveDocument.GetThemeNamesU visThemeTypeColor, astrNames 
     
    For intArrayCounter = LBound(astrNames) To UBound(astrNames) 
        strThemeName = astrNames(intArrayCounter) 
        Debug.Print strThemeName 
    Next 
     
    Debug.Print "-------------------------------------------" 
     
    ActiveDocument.GetThemeNamesU visThemeTypeEffect, astrNames 
     
    For intArrayCounter = LBound(astrNames) To UBound(astrNames) 
        strThemeName = astrNames(intArrayCounter) 
        Debug.Print strThemeName 
    Next 
     
End Sub
```


