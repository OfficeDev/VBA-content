---
title: Using Fields with Controls
keywords: olmain11.chm1045294
f1_keywords:
- olmain11.chm1045294
ms.prod: outlook
ms.assetid: 83618967-a027-13f7-4963-8656093074e4
ms.date: 06/08/2017
---


# Using Fields with Controls

When you drag a field from the  **Field Chooser**, the field automatically binds to the appropriate control. Unless you have a special requirement to use a standard control from the  **Control Toolbox**, you should use the  **Field Chooser** to provide access to fields on your forms.

When you put a control from the  **Control Toolbox** on a form, you must bind the control to a form if you want to save a value to or from a control. In most cases, you would bind controls like the check box, combo box, date, list box, option button, and text box to fields. Other controls, such as the **[Image](image-object-outlook-forms-script.md)** and label controls, that contain static information with which the user does not interact are generally not bound to a field.

To bind a control to a field, right-click the control, and then click  **Properties** on the shortcut menu. Click the **Value** tab. Click **Choose Field**, and then click a field or click  **New** to create a custom field. Outlook fields are based on MAPI properties. In this way, the values of fields are stored with the item when you save or send the item. The controls from the **Control Toolbox** are only the visual containers for a field on a form. You can set the appearance of the control using its properties, but you cannot save a value. Controls only exist when the specific form appears that contains the controls. Fields can be used on any form. If you change a field value in one place, this value changes everywhere the field is used.

For example, to change the value of a custom field called Fax, you use the following code:



```
Item.UserProperties.Find("Fax").Value = "555-1234"

```

Note that since this is a field, you do not need to specify the page or the control the field is bound to. In the following code example, a control called txtFax is made invisible. When you work with a control, you must specify the page and the control name.



```vb
Item.GetInspector.ModifiedFormPages("General").Controls("txtFax").Visible = False
```

You can bind a control to a field at run time by using the internal property named  **ItemProperty**. The following example binds a  **TextBox** to a field named Business Address.



```
Item.GetInspector.SetControlItemProperty("Textbox1", "Business Address")
```


 **Note**  If you create a control by dragging a plain text field to a form, you cannot bind the control to a field of a different type. For example, you cannot drag a Subject field to a form and then bind it to a field containing an Email type (such as the To field).


