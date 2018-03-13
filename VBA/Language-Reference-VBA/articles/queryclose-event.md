---
title: QueryClose Event
keywords: vblr6.chm1107501
f1_keywords:
- vblr6.chm1107501
ms.prod: office
api_name:
- Office.QueryClose
ms.assetid: 8a12c265-bbb8-ed72-8bde-7b9c3bdf86bd
ms.date: 06/08/2017
---


# QueryClose Event



Occurs before a  **UserForm** closes.
 **Syntax**
 **Private Sub UserForm_QueryClose(**_cancel_**As Integer**, _closemode_**As Integer)**
The  **QueryClose** event syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                                                                                                                                                |
|:----------------------|:------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>cancel</em>       | An integer. Setting this [argument](vbe-glossary.md) to any value other than 0 stops the QueryClose event in all loaded user forms and prevents the <strong>UserForm</strong> and application from closing. |
| <em>closemode</em>    | A value or [constant](vbe-glossary.md) indicating the cause of the QueryClose event.                                                                                                                        |

 **Return Values**
The  _closemode_ argument returns the following values:


| <strong>Constant</strong>          | <strong>Value</strong> | <strong>Description</strong>                                                                                                     |
|:-----------------------------------|:-----------------------|:---------------------------------------------------------------------------------------------------------------------------------|
| <strong>vbFormControlMenu</strong> | 0                      | The user has chosen the  <strong>Close</strong> command from the <strong>Control</strong> menu on the <strong>UserForm</strong>. |
| <strong>vbFormCode</strong>        | 1                      | The  <strong>Unload</strong> statement is invoked from code.                                                                     |
| <strong>vbAppWindows</strong>      | 2                      | The current Windows operating environment session is ending.                                                                     |
| <strong>vbAppTaskManager</strong>  | 3                      | The Windows  <strong>Task Manager</strong> is closing the application.                                                           |

These constants are listed in the Visual Basic for Applications [object library](vbe-glossary.md) in the[Object Browser](vbe-glossary.md). Note that  **vbFormMDIForm** is also specified in the **Object Browser**, but is not yet supported.
 **Remarks**
This event is typically used to make sure there are no unfinished tasks in the user forms included in an application before that application closes. For example, if a user hasn't saved new data in any  **UserForm**, the application can prompt the user to save the data.
When an application closes, you can use the  **QueryClose** event procedure to set the **Cancel** property to **True**, stopping the closing process.

## Example

The following code forces the user to click the  **UserForm's** client area to close it. If the user tries to use the **Close** box in the title bar, the _Cancel_ parameter is set to a nonzero value, preventing termination. However, if the user has clicked the client area, _CloseMode_ has the value 1 and `Unload Me` is executed.


```vb
Private Sub UserForm_Activate()
    UserForm1.Caption = "You must Click me to kill me!"
End Sub

Private Sub UserForm_Click()
  Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'Prevent user from closing with the Close box in the title bar.
    If CloseMode <> 1 Then Cancel = 1
    UserForm1.Caption = "The Close box won't work! Click me!"
End Sub
```


