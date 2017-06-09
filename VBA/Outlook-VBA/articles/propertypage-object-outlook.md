---
title: PropertyPage Object (Outlook)
keywords: vbaol11.chm380
f1_keywords:
- vbaol11.chm380
ms.prod: outlook
api_name:
- Outlook.PropertyPage
ms.assetid: 22e561d5-603e-2cf3-e142-6173dd0d4c25
ms.date: 06/08/2017
---


# PropertyPage Object (Outlook)

Represents a custom property page in the  **Options** dialog box or in the folder **Properties** dialog box.


## Remarks

Outlook uses this object to allow a custom property page to interact with the  **Apply** button in the dialog box.

The  **PropertyPage** object is an abstract object. That is, the **PropertyPage** object in the Microsoft Outlook Object Library contains no implementation code. Instead, it is provided as a template to help you implement the object in Microsoft Visual Basic for Applications (VBA). This provides a predefined set of interfaces that Outlook can use to determine whether your custom property page has changed and to notify your program that the user has clicked the **Apply** or **OK** button. (If your custom property page does not rely on the **Apply** button, then you do not need to implement the **PropertyPage** object.)

A custom property page is an ActiveX control that is displayed by Outlook in the  **Options** dialog box or in the folder **Properties** dialog box when the user clicks on the custom property page's tab. To implement the **PropertyPage** object, the module that contains the implementation code must contain the following statement.




```
Implements Outlook.PropertyPage
```

The module must also contain procedures that implement the properties and methods of the  **PropertyPage** object. For example, to implement the **Dirty** property, a procedure similar to the following appears in the module.




```
Private Property Get PropertyPage_Dirty() As Boolean 
 
 PropertyPage_Dirty = gblDirty 
 
End Property
```

To implement a method of the  **PropertyPage** object, the module must contain a statement similar to the following.




```
Private Sub PropertyPage_Apply() 
 
 ' Code to set properties according to the user's 
 
 ' selections goes here. 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Apply](propertypage-apply-method-outlook.md)|
|[GetPageInfo](propertypage-getpageinfo-method-outlook.md)|

## Properties



|**Name**|
|:-----|
|[Dirty](propertypage-dirty-property-outlook.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
