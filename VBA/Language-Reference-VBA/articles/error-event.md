---
title: Error Event
keywords: fm20.chm2000100
f1_keywords:
- fm20.chm2000100
ms.prod: office
api_name:
- Office.Error
ms.assetid: 12901147-8a12-b94b-0aa9-6cd9fe43b2e8
ms.date: 06/08/2017
---


# Error Event



Occurs when a control detects an error and cannot return the error information to a calling program.
 <strong>Syntax</strong>
For MultiPage 
<strong>Private Sub</strong><em>object</em> <em><strong>Error(</strong>_index</em><strong>As Long</strong>, <strong>ByVal</strong><em>Number</em><strong>As Integer</strong>, <strong>ByVal</strong><em>Description</em><strong>As MSForms.ReturnString</strong>, <strong>ByVal</strong><em>SCode</em><strong>As SCode</strong>, <strong>ByVal</strong><em>Source</em><strong>As String</strong>, <strong>ByVal</strong><em>HelpFile</em><strong>As String</strong>, <strong>ByVal</strong><em>HelpContext</em><strong>As Long</strong>, <strong>ByVal</strong><em>CancelDisplay</em><strong>As MSForms.ReturnBoolean)</strong>
For other controls 
<strong>Private Sub</strong><em>object</em> <em><strong>Error( ByVal</strong>_Number</em><strong>As Integer</strong>, <strong>ByVal</strong><em>Description</em><strong>As MSForms.ReturnString</strong>, <strong>ByVal</strong><em>SCode</em><strong>As SCode</strong>, <strong>ByVal</strong><em>Source</em><strong>As String</strong>, <strong>ByVal</strong><em>HelpFile</em><strong>As String</strong>, <strong>ByVal</strong><em>HelpContext</em><strong>As Long</strong>, <strong>ByVal</strong><em>CancelDisplay</em><strong>As MSForms.ReturnBoolean)</strong>
The  
<strong>Error</strong> event syntax has these parts:


| <strong>Part</strong>  | <strong>Description</strong>                                                                                                                                       |
|:-----------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>object</em>        | Required. A valid object name.                                                                                                                                     |
| <em>index</em>         | Required. The index of the page in a  <strong>MultiPage</strong> associated with this event.                                                                       |
| <em>Number</em>        | Required. Specifies a unique value that the control uses to identify the error.                                                                                    |
| <em>Description</em>   | Required. A textual description of the error.                                                                                                                      |
| <em>SCode</em>         | Required. Specifies the [OLE status code](glossary-vba.md) for the error. The low-order 16 bits specify a value that is identical to the <em>Number</em> argument. |
| <em>Source</em>        | Required. The string that identifies the control which initiated the event.                                                                                        |
| <em>HelpFile</em>      | Required. Specifies a fully qualified path name for the Help file that describes the error.                                                                        |
| <em>HelpContext</em>   | Required. Specifies the [context ID](glossary-vba.md) of the Help file topic that contains a description of the error.                                             |
| <em>CancelDisplay</em> | Required. Specifies whether to display the error string in a message box.                                                                                          |

 **Remarks**
The code written for the Error event determines how the control responds to the error condition.
The ability to handle error conditions varies from one application to another. The Error event is initiated when an error occurs that the application is not equipped to handle.

