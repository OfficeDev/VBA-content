---
title: SendKeys Statement
keywords: vblr6.chm1009015
f1_keywords:
- vblr6.chm1009015
ms.prod: office
ms.assetid: 8da3e83d-333a-444f-a660-917350fe2bc6
ms.date: 06/08/2017
---


# SendKeys Statement

Sends one or more keystrokes to the active window as if typed at the keyboard.

 **Syntax**

 **SendKeys** **_string_** [, **_wait_** ]

The  **SendKeys** statement syntax has these[named arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
|**_string_**|Required. [String expression](vbe-glossary.md) specifying the keystrokes to send.|
|**_Wait_**|Optional. [Boolean](vbe-glossary.md) value specifying the wait mode. If **False** (default), control is returned to the[procedure](vbe-glossary.md) immediately after the keys are sent. If **True**, keystrokes must be processed before control is returned to the procedure.|
 **Remarks**
Each key is represented by one or more characters. To specify a single keyboard character, use the character itself. For example, to represent the letter A, use  `"A"` for **_string_**. To represent more than one character, append each additional character to the one preceding it. To represent the letters A, B, and C, use `"ABC"` for **_string_**.
The plus sign ( **+** ), caret ( **^** ), percent sign ( **%** ), tilde ( **~** ), and parentheses **( )** have special meanings to **SendKeys**. To specify one of these characters, enclose it within braces ( `{}`). For example, to specify the plus sign, use  `{+}`. Brackets ([ ]) have no special meaning to  **SendKeys**, but you must enclose them in braces. In other applications, brackets do have a special meaning that may be significant when[dynamic data exchange](vbe-glossary.md) (DDE) occurs. To specify brace characters, use `{{}` and `{}}`.
To specify characters that aren't displayed when you press a key, such as ENTER or TAB, and keys that represent actions rather than characters, use the codes shown below:


|**Key**|**Code**|
|:-----|:-----|
|BACKSPACE| `{BACKSPACE}, {BS}, or{BKSP}`|
|BREAK| `{BREAK}`|
|CAPS LOCK| `{CAPSLOCK}`|
|DEL or DELETE| `{DELETE} or{DEL}`|
|DOWN ARROW| `{DOWN}`|
|END| `{END}`|
|ENTER| `{ENTER} or ~`|
|ESC| `{ESC}`|
|HELP| `{HELP}`|
|HOME| `{HOME}`|
|INS or INSERT| `{INSERT} or {INS}`|
|LEFT ARROW| `{LEFT}`|
|NUM LOCK| `{NUMLOCK}`|
|PAGE DOWN| `{PGDN}`|
|PAGE UP| `{PGUP}`|
|PRINT SCREEN| `{PRTSC}`|
|RIGHT ARROW| `{RIGHT}`|
|SCROLL LOCK| `{SCROLLLOCK}`|
|TAB| `{TAB}`|
|UP ARROW| `{UP}`|
|F1| `{F1}`|
|F2| `{F2}`|
|F3| `{F3}`|
|F4| `{F4}`|
|F5| `{F5}`|
|F6| `{F6}`|
|F7| `{F7}`|
|F8| `{F8}`|
|F9| `{F9}`|
|F10| `{F10}`|
|F11| `{F11}`|
|F12| `{F12}`|
|F13| `{F13}`|
|F14| `{F14}`|
|F15| `{F15}`|
|F16| `{F16}`|
To specify keys combined with any combination of the SHIFT, CTRL, and ALT keys, precede the key code with one or more of the following codes:


|**Key**|**Code**|
|:-----|:-----|
|SHIFT| `+`|
|CTRL| `^`|
|ALT| `%`|
To specify that any combination of SHIFT, CTRL, and ALT should be held down while several other keys are pressed, enclose the code for those keys in parentheses. For example, to specify to hold down SHIFT while E and C are pressed, use " `+(EC)`".
To specify repeating keys, use the form  `{key number}`. You must put a space between  `key` and `number`. For example,  `{LEFT 42}` means press the LEFT ARROW key 42 times; `{h 10}` means press H 10 times.

 **Note**  You can't use  **SendKeys** to send keystrokes to an application that is not designed to run in Microsoft Windows or Macintosh. **Sendkeys** also can't send the PRINT SCREEN key `{PRTSC}` to any application.


## Example

This example uses the  **Shell** function to run the Calculator application included with Microsoft Windows. It uses the **SendKeys** statement to send keystrokes to add some numbers, and then quit the Calculator. (To see the example, paste it into a procedure, then run the procedure. Because **AppActivate** changes the focus to the Calculator application, you can't single step through the code.). On the Macintosh, use a Macintosh application that accepts keyboard input instead of the Windows Calculator.


```vb
Dim ReturnValue, I 
ReturnValue = Shell("CALC.EXE", 1)    ' Run Calculator. 
AppActivate ReturnValue     ' Activate the Calculator. 
For I = 1 To 100    ' Set up counting loop. 
    SendKeys I &; "{+}", True    ' Send keystrokes to Calculator 
Next I    ' to add each value of I. 
SendKeys "=", True    ' Get grand total. 
SendKeys "%{F4}", True    ' Send ALT+F4 to close Calculator. 

```


