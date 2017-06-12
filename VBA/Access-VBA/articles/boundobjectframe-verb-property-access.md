---
title: BoundObjectFrame.Verb Property (Access)
keywords: vbaac10.chm10918
f1_keywords:
- vbaac10.chm10918
ms.prod: access
api_name:
- Access.BoundObjectFrame.Verb
ms.assetid: edbca2b1-fe7a-f0d0-1baf-fedbccb6dfb7
ms.date: 06/08/2017
---


# BoundObjectFrame.Verb Property (Access)

You can use the  **Verb** property to specify the operation to perform when an OLE object is activated, which is permitted when the control's **Action** property is set to **acOLEActivate**. Read/write **Long**.


## Syntax

 _expression_. **Verb**

 _expression_ A variable that represents a **BoundObjectFrame** object.


## Remarks

You can set the  **Verb** property by specifying an **Integer** data type value indicating the position of a verb in the list of verbs returned by the **ObjectVerbs** property. You can set the **Verb** property to 1 to specify the first verb in the list, you can set it to 2 to specify the second verb in the list, and so on.

If you don't use the  **ObjectVerbs** property to identify a specific verb, you can set the **Verb** property to one of the following values to indicate the operation to perform. These values specify the standard verbs supported by all objects.



|**Constant**|**Description**|
|:-----|:-----|
|**acOLEVerbPrimary**|Performs the default operation for the object.|
|**acOLEVerbShow**|Activates the object for editing.|
|**acOLEVerbOpen**|Opens the object in a separate application window.|
|**acOLEVerbHide**|For embedded objects, hides the application that was used to create the object.|
With some applications' objects, you can use these additional values. 



|**Constant**|**Description**|
|:-----|:-----|
|**acOLEVerbInPlaceUIActivate**|Activates the object for editing within the control. The menus and toolbars of the OLE server become available in the OLE container.|
|**acOLEVerbInPlaceActivate**|Activates the object within the control. The menus and toolbars of the OLE server aren't available in the OLE container.|
Each object supports its own set of verbs. For example, many objects support the verbs Edit and Play. You can use the  **ObjectVerbs** and **ObjectVerbsCount** properties to find out which verbs are supported by an object.

Microsoft Access automatically uses an object's default verb if the user double-clicks an object for which the  **AutoActivate** property is set to Double-Click.


## Example

The following example activates the control "OLEUnbound0" in the form "frmOperations" by opening up the OLE object in its own application window for editing. In this case, "OLEUnbound0" contains a new bitmap image, which is linked to the Microsoft Paint program.


```vb
With Forms.Item("frmOperations").Controls.Item("OLEUnbound0") 
 .Action = acOLEActivate 
 .Verb = acOLEVerbOpen 
End With
```


## See also


#### Concepts


[BoundObjectFrame Object](boundobjectframe-object-access.md)

