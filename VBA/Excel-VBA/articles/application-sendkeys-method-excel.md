---
title: Application.SendKeys Method (Excel)
keywords: vbaxl10.chm183108
f1_keywords:
- vbaxl10.chm183108
ms.prod: excel
api_name:
- Excel.Application.SendKeys
ms.assetid: 585666b9-adbc-5d04-c480-58e55ea7fb9d
ms.date: 06/08/2017
---


# Application.SendKeys Method (Excel)

Sends keystrokes to the active application.


## Syntax

 _expression_ . **SendKeys**( **_Keys_** , **_Wait_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Keys_|Required| **Variant**|The key or key combination you want to send to the application, as text.|
| _Wait_|Optional| **Variant**| **True** to have Microsoft Excel wait for the keys to be processed before returning control to the macro. **False** (or omitted) to continue running the macro without waiting for the keys to be processed.|

## Remarks

This method places keystrokes in a key buffer. In some cases, you must call this method before you call the method that will use the keystrokes. For example, to send a password to a dialog box, you must call the  **SendKeys** method before you display the dialog box.

The  _Keys_ argument can specify any single key or any key combined with ALT, CTRL, or SHIFT (or any combination of those keys). Each key is represented by one or more characters, such as `"a"` for the character a, or `"{ENTER}"` for the ENTER key.

To specify characters that aren't displayed when you press the corresponding key (for example, ENTER or TAB), use the codes listed in the following table. Each code in the table represents one key on the keyboard.



|**Key**|**Code**|
|:-----|:-----|
|BACKSPACE| `{BACKSPACE}` or `{BS}`|
|BREAK| `{BREAK}`|
|CAPS LOCK| `{CAPSLOCK}`|
|CLEAR| `{CLEAR}`|
|DELETE or DEL| `{DELETE}` or `{DEL}`|
|DOWN ARROW| `{DOWN}`|
|END| `{END}`|
|ENTER (numeric keypad)| `{ENTER}`|
|ENTER| `~` (tilde)|
|ESC| `{ESCAPE}` or `{ESC}`|
|HELP| `{HELP}`|
|HOME| `{HOME}`|
|INS| `{INSERT}`|
|LEFT ARROW| `{LEFT}`|
|NUM LOCK| `{NUMLOCK}`|
|PAGE DOWN| `{PGDN}`|
|PAGE UP| `{PGUP}`|
|RETURN| `{RETURN}`|
|RIGHT ARROW| `{RIGHT}`|
|SCROLL LOCK| `{SCROLLLOCK}`|
|TAB| `{TAB}`|
|UP ARROW| `{UP}`|
|F1 through F15| `{F1}` through `{F15}`|
You can also specify keys combined with SHIFT and/or CTRL and/or ALT. To specify a key combined with another key or keys, use the following table.



|**To combine a key with**|**Precede the key code with**|
|:-----|:-----|
|SHIFT| `+` (plus sign)|
|CTRL| `^` (caret)|
|ALT| `%` (percent sign)|

## Example

This example uses the  **SendKeys** method to quit Microsoft Excel.


```vb
Application.SendKeys("%fx")
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

