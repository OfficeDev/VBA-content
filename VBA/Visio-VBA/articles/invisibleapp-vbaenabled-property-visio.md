---
title: InvisibleApp.VBAEnabled Property (Visio)
keywords: vis_sdr.chm17552085
f1_keywords:
- vis_sdr.chm17552085
ms.prod: visio
api_name:
- Visio.InvisibleApp.VBAEnabled
ms.assetid: d95a7957-50f2-9eaa-a493-bfa55d5c6cb9
ms.date: 06/08/2017
---


# InvisibleApp.VBAEnabled Property (Visio)

Specifies whether Microsoft Visual Basic for Applications (VBA) is enabled in the application. Read-only.


## Syntax

 _expression_ . **VBAEnabled**

 _expression_ A variable that represents an **InvisibleApp** object.


### Return Value

Boolean


## Remarks

If a document that contains a VBA project is opened with VBA enabled, and then VBA becomes disabled while the document is open:




- Microsoft Visio no longer executes macros in that document, but the macro names still appear in the  **Macros** dialog box (press Alt+F8).
    
- Visio continues firing events to the project.
    


If a document that contains a VBA project is opened with VBA disabled, and then VBA becomes enabled while the document is open:




- Visio does not fire events to the project, even though VBA has become enabled.
    
- Macros remain disabled.
    


The  **VBAEnabled** property is set to **True** if the **Trust access to the VBA project object model** check box is selected under **Developer Macro Settings** on the **Macro Settings** page of the **Trust Center** (click the **BackstageButton** tab, click **Options**, click  **Trust Center**, and then click  **Trust Center Settings**). If it is not selected, the property reports  **False** .


## Example

You may have a document that requires VBA to be enabled to run properly, for example, code in a document's  **DocumentOpened** event handler. The following code can be run from an add-on to verify whether VBA is enabled in the application before a document that depends on VBA is opened.

Before running this procedure, supply a valid document file name for the variable  _filename_ .




```vb
Public Sub VBAEnabled_Example() 
 
 Dim vsoDocument As Visio.Document 
 Dim blsStatus As Boolean 
 
 blsStatus = Application.VBAEnabled 
 If Not blsStatus Then 
 
 MsgBox "For this process to continue, VBA must be enabled." &; _ 
 " Please enable VBA and start over." 
 
 Else 
 
 Set vsoDocument = Documents.Open("filename ") 
 
 End if 
 
End Sub
```


