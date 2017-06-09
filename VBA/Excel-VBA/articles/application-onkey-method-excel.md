---
title: Application.OnKey Method (Excel)
keywords: vbaxl10.chm133180
f1_keywords:
- vbaxl10.chm133180
ms.prod: excel
api_name:
- Excel.Application.OnKey
ms.assetid: 43662d8b-19e2-2b4a-4c3a-c64be4007643
ms.date: 06/08/2017
---


# Application.OnKey Method (Excel)

Runs a specified procedure when a particular key or key combination is pressed.


## Syntax

 _expression_ . **OnKey**( **_Key_** , **_Procedure_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Key_|Required| **String**|A string indicating the key to be pressed.|
| _Procedure_|Optional| **Variant**|A string indicating the name of the procedure to be run. If  _Procedure_ is "" (empty text), nothing happens when _Key_ is pressed. This form of **OnKey** changes the normal result of keystrokes in Microsoft Excel. If _Procedure_ is omitted, _Key_ reverts to its normal result in Microsoft Excel, and any special key assignments made with previous **OnKey** methods are cleared.|

## Remarks

The  _Key_ argument can specify any single key combined with ALT, CTRL, or SHIFT, or any combination of these keys. Each key is represented by one or more characters, such as `"a"` for the character a, or `"{ENTER}"` for the ENTER key.

To specify characters that aren't displayed when you press the corresponding key (ENTER or TAB, for example), use the codes listed in the following table. Each code in the table represents one key on the keyboard.



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
|ESC|{ `ESCAPE}` or `{ESC}`|
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



|**To combine keys with**|**Precede the key code by**|
|:-----|:-----|
|SHIFT| `+` (plus sign)|
|CTRL| `^` (caret)|
|ALT| `%` (percent sign)|
To assign a procedure to one of the special characters (+, ^, %, and so on), enclose the character in braces. For details, see the example.


## Example

This example assigns "InsertProc" to the key sequence CTRL+PLUS SIGN and assigns "SpecialPrintProc" to the key sequence SHIFT+CTRL+RIGHT ARROW.


```vb
Application.OnKey "^{+}", "InsertProc" 
Application.OnKey "+^{RIGHT}", "SpecialPrintProc"
```

This example returns SHIFT+CTRL+RIGHT ARROW to its normal meaning.




```vb
Application.OnKey "+^{RIGHT}"
```

This example disables the SHIFT+CTRL+RIGHT ARROW key sequence.




```vb
Application.OnKey "+^{RIGHT}", ""
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

