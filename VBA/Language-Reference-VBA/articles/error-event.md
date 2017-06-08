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
 **Syntax**
For MultiPage **Private Sub**_object_ _**Error(**_index_**As Long**, **ByVal**_Number_**As Integer**, **ByVal**_Description_**As MSForms.ReturnString**, **ByVal**_SCode_**As SCode**, **ByVal**_Source_**As String**, **ByVal**_HelpFile_**As String**, **ByVal**_HelpContext_**As Long**, **ByVal**_CancelDisplay_**As MSForms.ReturnBoolean)**
For other controls **Private Sub**_object_ _**Error( ByVal**_Number_**As Integer**, **ByVal**_Description_**As MSForms.ReturnString**, **ByVal**_SCode_**As SCode**, **ByVal**_Source_**As String**, **ByVal**_HelpFile_**As String**, **ByVal**_HelpContext_**As Long**, **ByVal**_CancelDisplay_**As MSForms.ReturnBoolean)**
The  **Error** event syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object name.|
| _index_|Required. The index of the page in a  **MultiPage** associated with this event.|
| _Number_|Required. Specifies a unique value that the control uses to identify the error.|
| _Description_|Required. A textual description of the error.|
| _SCode_|Required. Specifies the [OLE status code](glossary-vba.md) for the error. The low-order 16 bits specify a value that is identical to the _Number_ argument.|
| _Source_|Required. The string that identifies the control which initiated the event.|
| _HelpFile_|Required. Specifies a fully qualified path name for the Help file that describes the error.|
| _HelpContext_|Required. Specifies the [context ID](glossary-vba.md) of the Help file topic that contains a description of the error.|
| _CancelDisplay_|Required. Specifies whether to display the error string in a message box.|
 **Remarks**
The code written for the Error event determines how the control responds to the error condition.
The ability to handle error conditions varies from one application to another. The Error event is initiated when an error occurs that the application is not equipped to handle.

