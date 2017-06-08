---
title: Application.DDEExecute Method (Excel)
keywords: vbaxl10.chm132089
f1_keywords:
- vbaxl10.chm132089
ms.prod: excel
api_name:
- Excel.Application.DDEExecute
ms.assetid: 18cd97e6-4dff-2386-84bf-25e8c85b5277
ms.date: 06/08/2017
---


# Application.DDEExecute Method (Excel)

Runs a command or performs some other action or actions in another application by way of the specified DDE channel.


## Syntax

 _expression_ . **DDEExecute**( **_Channel_** , **_String_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Channel_|Required| **Long**|The channel number returned by the  **[DDEInitiate](application-ddeinitiate-method-excel.md)** method.|
| _String_|Required| **String**|The message defined in the receiving application.|

## Remarks

The  **DDEExecute** method is designed to send commands to another application. You can also use it to send keystrokes to another application, although the **[SendKeys](application-sendkeys-method-excel.md)** method is the preferred way to send keystrokes. The _String_ argument can specify any single key combined with ALT, CTRL, or SHIFT, or any combination of those keys. Each key is represented by one or more characters, such as `"a"` for the character a, or `"{ENTER}"` for the ENTER key.

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
You can also specify keys combined with SHIFT and/or CTRL and/or ALT. To specify a key combined with one or more of the keys just mentioned, use the following table.



|**To combine a key with**|**Precede the key code with**|
|:-----|:-----|
|SHIFT| `+` (plus sign)|
|CTRL| `^` (caret)|
|ALT| `%` (percent sign)|

## Example

This example opens a channel to Word, opens the Word document Formletr.doc, and then sends the FilePrint command to WordBasic.


```
channelNumber = Application.DDEInitiate( _ 
 app:="WinWord", _ 
 topic:="C:\WINWORD\FORMLETR.DOC") 
Application.DDEExecute channelNumber, "[FILEPRINT]" 
Application.DDETerminate channelNumber
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

