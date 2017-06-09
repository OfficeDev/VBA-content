---
title: Options.DefaultTray Property (Word)
keywords: vbawd10.chm162988071
f1_keywords:
- vbawd10.chm162988071
ms.prod: word
api_name:
- Word.Options.DefaultTray
ms.assetid: e96df21c-2cfc-3f14-5283-107187d193af
ms.date: 06/08/2017
---


# Options.DefaultTray Property (Word)

Returns or sets the default tray your printer uses to print documents. Read/write  **String** .


## Syntax

 _expression_ . **DefaultTray**

 _expression_ A variable that represents a **[Options](options-object-word.md)** object.


## Remarks

When setting this property, you must specify a string found in the  **Default** tray box on the **Print** tab in the **Options** dialog box. You can use the **[DefaultTrayID](options-defaulttrayid-property-word.md)** property and specify a **WdPaperTray** constant to set this same option.


## Example

This example sets Word up to use the lower print tray.


```
Options.DefaultTray = "Lower tray"
```

This example returns the string found in the  **Default tray** box on the **Print** tab in the **Options** dialog box.




```
Msgbox Options.DefaultTray
```


## See also


#### Concepts


[Options Object](options-object-word.md)

