---
title: GetAllSettings Function
keywords: vblr6.chm1020903
f1_keywords:
- vblr6.chm1020903
ms.prod: office
ms.assetid: f87675b2-d14e-593d-94ab-259ab8da094d
ms.date: 06/08/2017
---


# GetAllSettings Function



Returns a list of key settings and their respective values (originally created with  **SaveSetting** ) from an application's entry in the Windows[registry](vbe-glossary.md) or (on the Macintosh) information in the application's initialization file.
 **Syntax**
 **GetAllSettings( _appname,_** **_section_ )**
The  **GetAllSettings** function syntax has these[named arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
|**_appname_**|Required. [String expression](vbe-glossary.md) containing the name of the application or[project](vbe-glossary.md) whose key settings are requested. On the Macintosh, this is the filename of the initialization file in the Preferences folder in the System folder.|
|**_section_**|Required. String e **xpression** containing the name of the section whose key settings are requested. **GetAllSettings** returns a[Variant](vbe-glossary.md) whose contents is a two-dimensional[array](vbe-glossary.md) of strings containing all the key settings in the specified section and their corresponding values.|
 **Remarks**
 **GetAllSettings** returns an uninitialized **Variant** if either **_appname_** or **_section_** does not exist.

## Example

This example first uses the  **SaveSetting** statement to make entries in the Windows registry for the application specified as **_appname_**, then uses the **GetAllSettings** function to display the settings. Note that application names and **_section_** names can't be retrieved with **GetAllSettings**. Finally, the **DeleteSetting** statement removes the application's entries.


```vb
' Variant to hold 2-dimensional array returned by GetAllSettings
' Integer to hold counter.
Dim MySettings As Variant, intSettings As Integer
' Place some settings in the registry.
SaveSetting appname := "MyApp", section := "Startup", _
key := "Top", setting := 75
SaveSetting "MyApp","Startup", "Left", 50
' Retrieve the settings.
MySettings = GetAllSettings(appname := "MyApp", section := "Startup")
    For intSettings = LBound(MySettings, 1) To UBound(MySettings, 1)
        Debug.Print MySettings(intSettings, 0), MySettings(intSettings, 1)
    Next intSettings
DeleteSetting "MyApp", "Startup"


```


