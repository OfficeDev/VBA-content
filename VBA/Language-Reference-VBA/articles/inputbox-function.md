---
title: InputBox Function
keywords: vblr6.chm1008945
f1_keywords:
- vblr6.chm1008945
ms.prod: office
ms.assetid: 701fb7bb-3663-93ae-df74-a2fd401215f6
ms.date: 06/08/2017
---


# InputBox Function



Displays a prompt in a dialog box, waits for the user to input text or click a button, and returns a [String](vbe-glossary.md) containing the contents of the text box.
 **Syntax**
 **InputBox( _prompt_** [, **_title_** ] [, **_default_** ] [, **_xpos_** ] [, **_ypos_** ] [, **_helpfile_**, **_context_** ] **)**
The  **InputBox** function syntax has these[named arguments](vbe-glossary.md):


| <strong>Part</strong>              | <strong>Description</strong>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   |
|:-----------------------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <strong><em>prompt</em></strong>   | Required. [String expression](vbe-glossary.md) displayed as the message in the dialog box. The maximum length of <strong><em>prompt</em></strong> is approximately 1024 characters, depending on the width of the characters used. If <strong><em>prompt</em></strong> consists of more than one line, you can separate the lines using a carriage return character ( <strong>Chr(</strong> 13 <strong>)</strong> ), a linefeed character ( <strong>Chr(</strong> 10 <strong>)</strong> ), or carriage return-linefeed character combination ( <strong>Chr(</strong> 13 <strong>)</strong> &; <strong>Chr(</strong> 10 <strong>)</strong> ) between each line. |
| <strong><em>title</em></strong>    | Optional. String expression displayed in the title bar of the dialog box. If you omit  <strong><em>title</em></strong>, the application name is placed in the title bar.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| <strong><em>default</em></strong>  | Optional. String expression displayed in the text box as the default response if no other input is provided. If you omit  <strong><em>default</em></strong>, the text box is displayed empty.                                                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| <strong><em>xpos</em></strong>     | Optional. [Numeric expression](vbe-glossary.md) that specifies, in twips, the horizontal distance of the left edge of the dialog box from the left edge of the screen. If <strong><em>xpos</em></strong> is omitted, the dialog box is horizontally centered.                                                                                                                                                                                                                                                                                                                                                                                                  |
| <strong><em>ypos</em></strong>     | Optional. Numeric expression that specifies, in twips, the vertical distance of the upper edge of the dialog box from the top of the screen. If  <strong><em>ypos</em></strong> is omitted, the dialog box is vertically positioned approximately one-third of the way down the screen.                                                                                                                                                                                                                                                                                                                                                                        |
| <strong><em>helpfile</em></strong> | Optional. String expression that identifies the Help file to use to provide context-sensitive Help for the dialog box. If  <strong><em>helpfile</em></strong> is provided, <strong><em>context</em></strong> must also be provided.                                                                                                                                                                                                                                                                                                                                                                                                                            |
| <strong><em>context</em></strong>  | Optional. Numeric expression that is the Help context number assigned to the appropriate Help topic by the Help author. If  <strong><em>context</em></strong> is provided, <strong><em>helpfile</em></strong> must also be provided.                                                                                                                                                                                                                                                                                                                                                                                                                           |

 **Remarks**
When both  **_helpfile_** and **_context_** are provided, the user can press F1 (Windows) or HELP (Macintosh) to view the Help topic corresponding to the **_context_**. Some[host applications](vbe-glossary.md), for example, Microsoft Excel, also automatically add a  **Help** button to the dialog box. If the user clicks **OK** or presses ENTER , the **InputBox** function returns whatever is in the text box. If the user clicks **Cancel**, the function returns a zero-length string ("").

 **Note**  To specify more than the first named argument, you must use  **InputBox** in an[expression](vbe-glossary.md). To omit some positional [arguments](vbe-glossary.md), you must include the corresponding comma delimiter.


## InputBox Function Example

This example shows various ways to use the  **InputBox** function to prompt the user to enter a value. If the x and y positions are omitted, the dialog box is automatically centered for the respective axes. The variable `MyValue` contains the value entered by the user if the user clicks **OK** or presses the ENTER key . If the user clicks **Cancel**, a zero-length string is returned.


```vb
Dim Message, Title, Default, MyValue
Message = "Enter a value between 1 and 3"    ' Set prompt.
Title = "InputBox Demo"    ' Set title.
Default = "1"    ' Set default.
' Display message, title, and default value.
MyValue = InputBox(Message, Title, Default)

' Use Helpfile and context. The Help button is added automatically.
MyValue = InputBox(Message, Title, , , , "DEMO.HLP", 10)

' Display dialog box at position 100, 100.
MyValue = InputBox(Message, Title, Default, 100, 100)
```


