---
title: PropertyPage.Apply Method (Outlook)
keywords: vbaol11.chm383
f1_keywords:
- vbaol11.chm383
ms.prod: outlook
api_name:
- Outlook.PropertyPage.Apply
ms.assetid: fdb35048-2471-4402-8137-c75994680b3c
ms.date: 06/08/2017
---


# PropertyPage.Apply Method (Outlook)

Applies the changes that have been made in a custom property page.


## Syntax

 _expression_ . **Apply**

 _expression_ A variable that represents a **PropertyPage** object.


### Return Value

An HRESULT value that represents the response of the event.


## Remarks

Because the [PropertyPage](propertypage-object-outlook.md) is an abstract object that is implemented in your application (rather than by Microsoft Outlook itself), the implementation of the **Apply** method resembles an event procedure in your program code. That is, you write the code that implements the method in much the same way you would write an event procedure. In other words, Outlook calls the **Apply** method to notify your program that the user has taken an action in the dialog box displaying the custom property page that requires your program to apply the property values changed by the user.


## Example

This Microsoft Visual Basic for Applications (VBA) example sets two global variables to reflect the values in controls on a form and then sets a global variable representing the  **[Dirty](propertypage-dirty-property-outlook.md)** property to **False** .


```vb
Private Sub PropertyPage_Apply() 
 
 globWorkGroup = Form1.Text1.Text 
 
 globUserType = Form1.Combo1.Text 
 
 globDirty = False 
 
End Sub
```


## See also


#### Concepts


[PropertyPage Object](propertypage-object-outlook.md)

