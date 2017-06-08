---
title: Global.CustomizationContext Property (Word)
keywords: vbawd10.chm163119172
f1_keywords:
- vbawd10.chm163119172
ms.prod: word
api_name:
- Word.Global.CustomizationContext
ms.assetid: e541c2ee-4a4e-5fc0-fd1a-5c9a99d8f7e9
ms.date: 06/08/2017
---


# Global.CustomizationContext Property (Word)

Returns or sets a  **Template** or **[Document](document-object-word.md)** object that represents the template or document in which changes to menu bars, toolbars, and key bindings are stored. Read/write. .


## Syntax

 _expression_ . **CustomizationContext**

 _expression_ A variable that represents a **[Global](global-object-word.md)** object.


## Remarks

Corresponds to the value of the  **Save in** box on the **Commands** tab in the **Customize** dialog box ( **Tools** menu).


## Example

This example adds the ALT+CTRL+W key combination to the FileClose command. The keyboard customization is saved in the Normal template.


```
CustomizationContext = NormalTemplate 
KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, _ 
 wdKeyAlt, wdKeyW), _ 
 KeyCategory:=wdKeyCategoryCommand, Command:="FileClose"
```

This example adds the File Versions button to the Standard toolbar. The command bar customization is saved in the template attached to the active document.




```
CustomizationContext = ActiveDocument.AttachedTemplate 
Application.CommandBars("Standard").Controls.Add _ 
 Type:=msoControlButton, _ 
 ID:=2522, Before:=8
```


## See also


#### Concepts


[Global Object](global-object-word.md)

