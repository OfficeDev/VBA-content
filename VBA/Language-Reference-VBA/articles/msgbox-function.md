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


|**Part**|**Description**|
|:-----|:-----|
|**_prompt_**|Required. [String expression](vbe-glossary.md) displayed as the message in the dialog box. The maximum length of **_prompt_** is approximately 1024 characters, depending on the width of the characters used. If **_prompt_** consists of more than one line, you can separate the lines using a carriage return character ( **Chr(** 13 **)** ), a linefeed character ( **Chr(** 10 **)** ), or carriage return - linefeed character combination ( **Chr(** 13 **)** &; **Chr(** 10 **)** ) between each line.|
|**_buttons_**|Optional. [Numeric expression](vbe-glossary.md) that is the sum of values specifying the number and type of buttons to display, the icon style to use, the identity of the default button, and the modality of the message box. If omitted, the default value for **_buttons_** is 0.|
|**_title_**|Optional. String expression displayed in the title bar of the dialog box. If you omit  **_title_**, the application name is placed in the title bar.|
|**_helpfile_**|Optional. String expression that identifies the Help file to use to provide context-sensitive Help for the dialog box. If  **_helpfile_** is provided, **_context_** must also be provided.|
|**_context_**|Optional. Numeric expression that is the Help context number assigned to the appropriate Help topic by the Help author. If  **_context_** is provided, **_helpfile_** must also be provided.|
 **Settings**
The  **_buttons_**[argument](vbe-glossary.md) settings are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbOKOnly**|0|Display  **OK** button only.|
|**vbOKCancel**|1|Display  **OK** and **Cancel** buttons.|
|**vbAbortRetryIgnore**|2|Display  **Abort**, **Retry**, and **Ignore** buttons.|
|**vbYesNoCancel**|3|Display  **Yes**, **No**, and **Cancel** buttons.|
|**vbYesNo**|4|Display  **Yes** and **No** buttons.|
|**vbRetryCancel**|5|Display  **Retry** and **Cancel** buttons.|
|**vbCritical**|16|Display  **Critical Message** icon.|
|**vbQuestion**|32|Display  **Warning Query** icon.|
|**vbExclamation**|48|Display  **Warning Message** icon.|
|**vbInformation**|64|Display  **Information Message** icon.|
|**vbDefaultButton1**|0|First button is default.|
|**vbDefaultButton2**|256|Second button is default.|
|**vbDefaultButton3**|512|Third button is default.|
|**vbDefaultButton4**|768|Fourth button is default.|
|**vbApplicationModal**|0|Application modal; the user must respond to the message box before continuing work in the current application.|
|**vbSystemModal**|4096|System modal; all applications are suspended until the user responds to the message box.|
|**vbMsgBoxHelpButton**|16384|Adds Help button to the message box.|
|**VbMsgBoxSetForeground**|65536|Specifies the message box window as the foreground window.|
|**vbMsgBoxRight**|524288|Text is right aligned.|
|**vbMsgBoxRtlReading**|1048576|Specifies text should appear as right-to-left reading on Hebrew and Arabic systems.|
The first group of values (0-5) describes the number and type of buttons displayed in the dialog box; the second group (16, 32, 48, 64) describes the icon style; the third group (0, 256, 512) determines which button is the default; and the fourth group (0, 4096) determines the modality of the message box. When adding numbers to create a final value for the  **_buttons_** argument, use only one number from each group.

 **Note**  These [constants](vbe-glossary.md) are specified by Visual Basic for Applications. As a result, the names can be used anywhere in your code in place of the actual values.

 **Return Values**


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
|**vbOK**|1|**OK**|
|**vbCancel**|2|**Cancel**|
|**vbAbort**|3|**Abort**|
|**vbRetry**|4|**Retry**|
|**vbIgnore**|5|**Ignore**|
|**vbYes**|6|**Yes**|
|**vbNo**|7|**No**|
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


