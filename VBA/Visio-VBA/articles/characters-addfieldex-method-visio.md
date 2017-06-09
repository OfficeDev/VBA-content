---
title: Characters.AddFieldEx Method (Visio)
keywords: vis_sdr.chm10251445
f1_keywords:
- vis_sdr.chm10251445
ms.prod: visio
api_name:
- Visio.Characters.AddFieldEx
ms.assetid: 14f56159-ed60-e1cf-1c04-b789672b51ec
ms.date: 06/08/2017
---


# Characters.AddFieldEx Method (Visio)

Replaces the text represented by a  **Characters** object with a new field of the category, code, format, language ID, and calendar ID you specify.


## Syntax

 _expression_ . **AddFieldEx**( **_Category_** , **_Code_** , **_Format_** , **_LangID_** , **_CalendarID_** )

 _expression_ A variable that represents a **Characters** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Category_|Required| **VisFieldCategories**| The category for the new field.|
| _Code_|Required| **VisFieldCodes**|The code for the new field.|
| _Format_|Required| **VisFieldFormats**|The format for the new field.|
| _LangID_|Optional| **Long**|The language to use for the new field. |
| _CalendarID_|Optional| **Long**|The calendar to use for the new field.|

### Return Value

Nothing


## Remarks

Constant values for  _Category, Code_, and  _Format_ are declared by the Visio type library in **[VisFieldCategories](visfieldcategories-enumeration-visio.md)** , **[VisFieldCodes](visfieldcodes-enumeration-visio.md)** , and **[VisFieldFormats](visfieldformats-enumeration-visio.md)** respectively.

The  _LangID_ argument should be one of the standard IDs used by Microsoft Windows to encode different language versions. For example, the language ID is &;H0409 for the U.S. version of Microsoft Visio. To see a list of possible language IDs, search for "VERSIONINFO" in the Microsoft Platform SDK on MSDN.

The  _CalendarID_ argument should be one of the following values, which are declared in **VisCellVals** in the Visio type library. The default value is **visCalWestern** , which sets the calendar to the Western calendar.



|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| **visCalWestern**|0|Western|
| **visCalArabicHijri**|1|Arabic Hijiri|
| **visCalHebrewLunar**|2|Hebrew Lunar|
| **visCalChineseTaiwan**|3|Taiwan Calendar|
| **visCalJapaneseEmperor**|4|Japanese Emperor Reign|
| **visCalThaiBuddhism**|5|Thai Buddhist|
| **visCalKoreanDanki**|6|Korean Danki|
| **visCalSakaEra**|7|Saka Era|
| **visCalTranslitEnglish**|8| English transliterated|
| **visCalTranslitFrench**|9| French transliterated|
Using the  **AddFieldEx** method is similar to clicking **Field** on the **Insert** tab and inserting any of the following categories of fields into the text:


- Date/Time
    
- Document Info
    
- Geometry
    
- Object Info
    
- Page Info
    


To add a custom formula field, use the  **AddCustomField** or **AddCustomFieldU** method. When you do not pass values (or pass default values) for the optional _LangID_ and _CalendarID_ arguments, **AddFieldEx** acts exactly like **AddField** .


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **AddFieldEx** property to format a date field in a shape. It draws a rectangle on the drawing page and then inserts a field that displays the current date in Greek, using the Western calendar.


```vb
Public Sub AddFieldEx_Example() 
 
 Dim vsoCharacters As Visio.Characters 
 Dim vsoShape As Visio.Shape 
 
 ActiveWindow.DeselectAll 
 
 Set vsoShape = Application.ActivePage.DrawRectangle(3, 5, 5, 3) 
 vsoShape.Text = "Date: " 
 
 Set vsoCharacters = vsoShape.Characters 
 
 'Set Begin property equal to End property to 
 'append new text to existing text. 
 vsoCharacters.Begin = vsoCharacters.End 
 
 'Add a field for the current date, in Greek, 
 'using the Western calendar and the long date format. 
 vsoCharacters.AddFieldEx visFCatDateTime, visFCodeCurrentDate, visFmtMsoDateLong, 1032, visCalWestern 
 
End Sub
```


