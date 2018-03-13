---
title: MsgBox Function
keywords: vblr6.chm1008978
f1_keywords:
- vblr6.chm1008978
ms.prod: office
ms.assetid: 715595a7-4286-a0cb-dec9-2d2e79bda102
ms.date: 06/08/2017
---


# MsgBox Function



Displays a message in a dialog box, waits for the user to click a button, and returns an  **Integer** indicating which button the user clicked.
 **Syntax**
 **MsgBox( _prompt_** [, **_buttons_** ] [, **_title_** ] [, **_helpfile_**, **_context_** ] **)**
The  **MsgBox** function syntax has these[named arguments](vbe-glossary.md):


| <strong>Part</strong>              | <strong>Description</strong>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     |
|:-----------------------------------|:-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <strong><em>prompt</em></strong>   | Required. [String expression](vbe-glossary.md) displayed as the message in the dialog box. The maximum length of <strong><em>prompt</em></strong> is approximately 1024 characters, depending on the width of the characters used. If <strong><em>prompt</em></strong> consists of more than one line, you can separate the lines using a carriage return character ( <strong>Chr(</strong> 13 <strong>)</strong> ), a linefeed character ( <strong>Chr(</strong> 10 <strong>)</strong> ), or carriage return - linefeed character combination ( <strong>Chr(</strong> 13 <strong>)</strong> &; <strong>Chr(</strong> 10 <strong>)</strong> ) between each line. |
| <strong><em>buttons</em></strong>  | Optional. [Numeric expression](vbe-glossary.md) that is the sum of values specifying the number and type of buttons to display, the icon style to use, the identity of the default button, and the modality of the message box. If omitted, the default value for <strong><em>buttons</em></strong> is 0.                                                                                                                                                                                                                                                                                                                                                        |
| <strong><em>title</em></strong>    | Optional. String expression displayed in the title bar of the dialog box. If you omit  <strong><em>title</em></strong>, the application name is placed in the title bar.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| <strong><em>helpfile</em></strong> | Optional. String expression that identifies the Help file to use to provide context-sensitive Help for the dialog box. If  <strong><em>helpfile</em></strong> is provided, <strong><em>context</em></strong> must also be provided.                                                                                                                                                                                                                                                                                                                                                                                                                              |
| <strong><em>context</em></strong>  | Optional. Numeric expression that is the Help context number assigned to the appropriate Help topic by the Help author. If  <strong><em>context</em></strong> is provided, <strong><em>helpfile</em></strong> must also be provided.                                                                                                                                                                                                                                                                                                                                                                                                                             |

 **Settings**
The  **_buttons_**[argument](vbe-glossary.md) settings are:


| <strong>Constant</strong>              | <strong>Value</strong> | <strong>Description</strong>                                                                                   |
|:---------------------------------------|:-----------------------|:---------------------------------------------------------------------------------------------------------------|
| <strong>vbOKOnly</strong>              | 0                      | Display  <strong>OK</strong> button only.                                                                      |
| <strong>vbOKCancel</strong>            | 1                      | Display  <strong>OK</strong> and <strong>Cancel</strong> buttons.                                              |
| <strong>vbAbortRetryIgnore</strong>    | 2                      | Display  <strong>Abort</strong>, <strong>Retry</strong>, and <strong>Ignore</strong> buttons.                  |
| <strong>vbYesNoCancel</strong>         | 3                      | Display  <strong>Yes</strong>, <strong>No</strong>, and <strong>Cancel</strong> buttons.                       |
| <strong>vbYesNo</strong>               | 4                      | Display  <strong>Yes</strong> and <strong>No</strong> buttons.                                                 |
| <strong>vbRetryCancel</strong>         | 5                      | Display  <strong>Retry</strong> and <strong>Cancel</strong> buttons.                                           |
| <strong>vbCritical</strong>            | 16                     | Display  <strong>Critical Message</strong> icon.                                                               |
| <strong>vbQuestion</strong>            | 32                     | Display  <strong>Warning Query</strong> icon.                                                                  |
| <strong>vbExclamation</strong>         | 48                     | Display  <strong>Warning Message</strong> icon.                                                                |
| <strong>vbInformation</strong>         | 64                     | Display  <strong>Information Message</strong> icon.                                                            |
| <strong>vbDefaultButton1</strong>      | 0                      | First button is default.                                                                                       |
| <strong>vbDefaultButton2</strong>      | 256                    | Second button is default.                                                                                      |
| <strong>vbDefaultButton3</strong>      | 512                    | Third button is default.                                                                                       |
| <strong>vbDefaultButton4</strong>      | 768                    | Fourth button is default.                                                                                      |
| <strong>vbApplicationModal</strong>    | 0                      | Application modal; the user must respond to the message box before continuing work in the current application. |
| <strong>vbSystemModal</strong>         | 4096                   | System modal; all applications are suspended until the user responds to the message box.                       |
| <strong>vbMsgBoxHelpButton</strong>    | 16384                  | Adds Help button to the message box.                                                                           |
| <strong>VbMsgBoxSetForeground</strong> | 65536                  | Specifies the message box window as the foreground window.                                                     |
| <strong>vbMsgBoxRight</strong>         | 524288                 | Text is right aligned.                                                                                         |
| <strong>vbMsgBoxRtlReading</strong>    | 1048576                | Specifies text should appear as right-to-left reading on Hebrew and Arabic systems.                            |

The first group of values (0-5) describes the number and type of buttons displayed in the dialog box; the second group (16, 32, 48, 64) describes the icon style; the third group (0, 256, 512) determines which button is the default; and the fourth group (0, 4096) determines the modality of the message box. When adding numbers to create a final value for the  **_buttons_** argument, use only one number from each group.

 **Note**  These [constants](vbe-glossary.md) are specified by Visual Basic for Applications. As a result, the names can be used anywhere in your code in place of the actual values.

 **Return Values**


| <strong>Constant</strong> | <strong>Value</strong> | <strong>Description</strong> |
|:--------------------------|:-----------------------|:-----------------------------|
| <strong>vbOK</strong>     | 1                      | <strong>OK</strong>          |
| <strong>vbCancel</strong> | 2                      | <strong>Cancel</strong>      |
| <strong>vbAbort</strong>  | 3                      | <strong>Abort</strong>       |
| <strong>vbRetry</strong>  | 4                      | <strong>Retry</strong>       |
| <strong>vbIgnore</strong> | 5                      | <strong>Ignore</strong>      |
| <strong>vbYes</strong>    | 6                      | <strong>Yes</strong>         |
| <strong>vbNo</strong>     | 7                      | <strong>No</strong>          |

 **Remarks**
When both  **_helpfile_** and **_context_** are provided, the user can press F1 (Windows) or HELP (Macintosh) to view the Help topic corresponding to the **context**. Some[host applications](vbe-glossary.md), for example, Microsoft Excel, also automatically add a  **Help** button to the dialog box.
If the dialog box displays a  **Cancel** button, pressing the ESC key has the same effect as clicking **Cancel**. If the dialog box contains a **Help** button, context-sensitive Help is provided for the dialog box. However, no value is returned until one of the other buttons is clicked.

 **Note**  To specify more than the first named argument, you must use  **MsgBox** in an[expression](vbe-glossary.md). To omit some positional [arguments](vbe-glossary.md), you must include the corresponding comma delimiter.


## Example

This example uses the  **MsgBox** function to display a critical-error message in a dialog box with Yes and No buttons. The No button is specified as the default response. The value returned by the **MsgBox** function depends on the button chosen by the user. This example assumes that `DEMO.HLP` is a Help file that contains a topic with a Help context number equal to is a Help file that contains a topic with a Help context number equal to `1000`.


```vb
Dim Msg, Style, Title, Help, Ctxt, Response, MyString
Msg = "Do you want to continue ?"    ' Define message.
Style = vbYesNo + vbCritical + vbDefaultButton2    ' Define buttons.
Title = "MsgBox Demonstration"    ' Define title.
Help = "DEMO.HLP"    ' Define Help file.
Ctxt = 1000    ' Define topic
        ' context. 
        ' Display message.
Response = MsgBox(Msg, Style, Title, Help, Ctxt)
If Response = vbYes Then    ' User chose Yes.
    MyString = "Yes"    ' Perform some action.
Else    ' User chose No.
    MyString = "No"    ' Perform some action.
End If
```


