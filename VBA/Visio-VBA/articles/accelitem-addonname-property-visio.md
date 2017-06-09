---
title: AccelItem.AddOnName Property (Visio)
keywords: vis_sdr.chm14513050
f1_keywords:
- vis_sdr.chm14513050
ms.prod: visio
api_name:
- Visio.AccelItem.AddOnName
ms.assetid: 5775be57-b4aa-6cdf-89d0-84a646fb8e55
ms.date: 06/08/2017
---


# AccelItem.AddOnName Property (Visio)

Gets or sets the name of an add-on or procedure that is run when its associated accelerator key is pressed. Read/write.


## Syntax

 _expression_ . **AddOnName**

 _expression_ A variable that represents an **AccelItem** object.


### Return Value

String


## Remarks

Assuming that the name of the add-on in the  **Addons** collection is _string_ , if the project of the currently active document (or another project if it is referenced) does not have a procedure named _string_ , or if the arguments passed in _string_ do not match those specified in the procedure, Microsoft Visio runs the add-on named _string_ . If no add-on named _string_ can be found, Visio does nothing and reports no error. (You can use the **TraceFlags** property to monitor the procedures and add-ons that Visio attempts to run.)

If  _string_ is an add-on, use the **AddOnArgs** property to specify arguments to send to the add-on when it is run.

If  _string_ is a procedure, specify arguments using _procname(arguments)_ or _procname arguments_ .

When calling a procedure in a standard module it is recommended that you prefix the string with the module name that contains the procedure (for example,  _moduleName.procName_ ) because more than one module can have a procedure with the same name.

To call a procedure in a project other than the project of the active document, use the syntax  _projName.modName.procName_ (you must have explicitly set a reference to _projName_ in your Visual Basic project).

If the  **AddOnName** property is set, Visio ignores the object's **CmdNum** property.


 **Note**  Beginning with Visio 2002, the  **AddOnName** property cannot execute a string that contains arbitrary VBA code. To call code that in previous versions of Visio you would have passed to the **AddOnName** property, move the code to a procedure in a document's VBA project that is called from the **AddOnName** property.


