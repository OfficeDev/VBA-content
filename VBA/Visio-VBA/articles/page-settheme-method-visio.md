---
title: Page.SetTheme Method (Visio)
ms.prod: visio
ms.assetid: 5a186f58-9a7a-bd8a-826b-85da75a4d59f
ms.date: 06/08/2017
---


# Page.SetTheme Method (Visio)

Sets the theme for the specified page.


## Syntax

 _expression_ . **SetTheme**_(varThemeIndex,_ _varColorScheme,_ _varEffectScheme,_ _varConnectorScheme,_ _varFontScheme)_

 _expression_ A variable that represents a **Page** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
|||||
| _varThemeIndex_|Required|VARIANT|The theme to apply.|
| _varColorScheme_|Optional|VARIANT|The color scheme theme component to apply.|
| _varEffectScheme_|Optional|VARIANT|The effect scheme theme component to apply.|
| _varConnectorScheme_|Optional|VARIANT|The connector scheme theme component to apply.|
| _varFontScheme_|Optional|VARIANT|The font scheme theme component to apply.|

### Return value

 **VOID**


## Remarks

Possible themes correspond to those displayed in the  **Themes** and the **Colors**,  **Effects**, and  **Connectors** galleries on the **Design** tab of the ribbon. You can specify values for just the first, required parameter, or for any combination of the first parameter and one or more of the other four parameters. If you pass a value for the only first parameter, _varThemeIndex_ and you pass nothing for the other four optional parameters, Visio sets all five parameters to the theme value you specified for the first parameter. For example, if you pass ?Linear? for the first parameter, Visio sets the color scheme, effect scheme, connector scheme, and font scheme to ?Linear? as well. If you pass ?Linear? for the first parameter and ?Gemstone? for the second parameter, Visio sets the effect scheme, connector scheme, and font scheme to ?Linear,? but sets the color scheme to ?Gemstone,? and so on.


## See also


#### Concepts


[Page Object](page-object-visio.md)

