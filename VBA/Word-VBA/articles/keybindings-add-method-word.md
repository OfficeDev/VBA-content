---
title: KeyBindings.Add Method (Word)
keywords: vbawd10.chm160825445
f1_keywords:
- vbawd10.chm160825445
ms.prod: word
api_name:
- Word.KeyBindings.Add
ms.assetid: b73a8af4-6e8f-7613-a8a5-b0c9f7c995ae
ms.date: 06/08/2017
---


# KeyBindings.Add Method (Word)

Returns a  **KeyBinding** object that represents a new shortcut key for a macro, built-in command, font, AutoText entry, style, or symbol.


## Syntax

 _expression_ . **Add**( **_KeyCategory_** , **_Command_** , **_KeyCode_** , **_KeyCode2_** , **_CommandParameter_** )

 _expression_ Required. A variable that represents a **[KeyBindings](keybindings-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _KeyCategory_|Required| **WdKeyCategory**|The category of the key assignment.|
| _Command_|Required| **String**|The command that the specified key combination executes.|
| _KeyCode_|Required| **Long**|A key you specify by using one of the  **WdKey** constants.|
| _KeyCode2_|Optional| **Variant**|A second key you specify by using one of the  **WdKey** constants.|
| _CommandParameter_|Optional| **Variant**|Additional text, if any, required for the command specified by Command. For details, see the Remarks section below.|

### Return Value

KeyBinding


## Remarks

You can use the  **BuildKeyCode** method to create the KeyCode or KeyCode2 argument.

In the following table, the left column contains commands that require a command value, and the right column describes what you must do to specify CommandParameter for each of these commands. (The equivalent action in the  **Customize Keyboard** dialog box ( **Tools** menu) to specifying CommandParameter is selecting an item in the list box that appears when you select one of the following commands in the **Commands** box.)



|**If Command is set to**|**CommandParameter must be**|
|:-----|:-----|
| **Borders** , **Color** , or **Shading**|A number ? specified as text ? corresponding to the position of the setting selected in the list box that contains values, where 0 (zero) is the first item, 1 is the second item, and so on|
| **Columns**|A number between 1 and 45 ? specified as text ? corresponding to the number of columns you want to apply|
| **Condensed**|A text measurement between 0.1 point and 12.75 points specified in 0.05-point increments (72 points = 1 inch)|
| **Expanded**|A text measurement between 0.1 point and 12.75 points specified in 0.05-point increments (72 points = 1 inch)|
| **FileOpenFile**|The path and file name of the file to be opened. If the path isn't specified, the current folder is used.|
| **Font Size**|A positive text measurement, specified in 0.5-point increments (72 points = 1 inch)|
| **Lowered, Raised**|A text measurement between 1 point and 64 points, specified in 0.5-point increments (72 points = 1 inch)|
| **Symbol**|A string created by concatenating a  **Chr()** instruction and the name of a symbol font (for example, `Chr(167) &; "Symbol"`)|

## Example

This example adds the CTRL+ALT+W key combination to the  **FileClose** command. The keyboard customization is saved in the Normal template.


```
CustomizationContext = NormalTemplate 
KeyBindings.Add _ 
    KeyCategory:=wdKeyCategoryCommand, _ 
    Command:="FileClose", _ 
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyW)
```

This example adds the ALT+F4 key combination to the Arial font and then displays the number of items in the  **KeyBindings** collection. The example then clears the ALT+F4 key combination (returned it to its default setting) and redisplays the number of items in the **KeyBindings** collection.




```
CustomizationContext = ActiveDocument.AttachedTemplate 
Set myKey = KeyBindings.Add(KeyCategory:=wdKeyCategoryFont, _ 
    Command:="Arial", KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyF4)) 
MsgBox KeyBindings.Count &; " keys in KeyBindings collection" 
myKey.Clear 
MsgBox KeyBindings.Count &; " keys in KeyBindings collection"
```

This example adds the CTRL+ALT+S key combination to the  **Font** command with 8 points specified for the font size.




```
CustomizationContext = NormalTemplate 
KeyBindings.Add KeyCategory:=wdKeyCategoryCommand, _ 
    Command:="FontSize", _ 
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyS), _ 
    CommandParameter:="8"
```

This example adds the CTRL+ALT+H key combination to the Heading 1 style in the active document.




```
CustomizationContext = ActiveDocument 
KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _ 
    Command:="Heading 1", _ 
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyH)
```

This example adds the CTRL+ALT+O key combination to the AutoText entry named "Hello."




```
CustomizationContext = ActiveDocument 
KeyBindings.Add KeyCategory:=wdKeyCategoryAutoText, _ 
    Command:="Hello", _ 
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyO)
```


## See also


#### Concepts


[KeyBindings Collection Object](keybindings-object-word.md)

