---
title: Add Method (Microsoft Forms)
keywords: fm20.chm5224953
f1_keywords:
- fm20.chm5224953
ms.prod: office
ms.assetid: 1030fff9-736c-9ece-5600-2e4f3b4281b8
ms.date: 06/08/2017
---


# Add Method (Microsoft Forms)



Adds or inserts a  **Tab** or **Page** in a **TabStrip** or **MultiPage**, or adds a control by its programmatic identifier ( _ProgID_ ) to a page or form.
 **Syntax**
For MultiPage, TabStrip **Set**_Object_ = _object_. **Add(** [ _Name_ [, _Caption_ [, _index_ ]]] **)**
For other controls **Set**_Control_ = _object_. **Add(**_ProgID_ [, _Name_ [, _Visible_ ]] **)**
The  **Add** method syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                                                                                                                                                                                                                                                    |
|:----------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>object</em>       | Required. A valid object name.                                                                                                                                                                                                                                                                                  |
| <em>Name</em>         | Optional. Specifies the name of the object being added. If a name is not specified, the system generates a default name based on the rules of the application where the form is used.                                                                                                                           |
| <em>Caption</em>      | Optional. Specifies the caption to appear on a tab or a control. If a caption is not specified, the system generates a default caption based on the rules of the application where the form is used.                                                                                                            |
| <em>index</em>        | Optional. Identifies the position of a page or tab within a  <strong>Pages</strong> or <strong>Tabs</strong> collection. If an index is not specified, the system appends the page or tab to the end of the <strong>Pages</strong> or <strong>Tabs</strong> collection and assigns the appropriate index value. |
| <em>ProgID</em>       | Required. Programmatic identifier. A text string with no spaces that identifies an object class. The standard syntax for a  <em>ProgID</em> is <Vendor>.<Component>.<Version>. A <em>ProgID</em> is mapped to a class identifier (CLSID).                                                                       |
| <em>Visible</em>      | Optional.  <strong>True</strong> if the object is visible (default). <strong>False</strong> if the object is hidden.                                                                                                                                                                                            |

 **Settings**
 _ProgID_ values for individual controls are:


|                                |                       |
|:-------------------------------|:----------------------|
| <strong>CheckBox</strong>      | Forms.CheckBox.1      |
| <strong>ComboBox</strong>      | Forms.ComboBox.1      |
| <strong>CommandButton</strong> | Forms.CommandButton.1 |
| <strong>Frame</strong>         | Forms.Frame.1         |
| <strong>Image</strong>         | Forms.Image.1         |
| <strong>Label</strong>         | Forms.Label.1         |
| <strong>ListBox</strong>       | Forms.ListBox.1       |
| <strong>MultiPage</strong>     | Forms.MultiPage.1     |
| <strong>OptionButton</strong>  | Forms.OptionButton.1  |
| <strong>ScrollBar</strong>     | Forms.ScrollBar.1     |
| <strong>SpinButton</strong>    | Forms.SpinButton.1    |
| <strong>TabStrip</strong>      | Forms.TabStrip.1      |
| <strong>TextBox</strong>       | Forms.TextBox.1       |
| <strong>ToggleButton</strong>  | Forms.ToggleButton.1  |

 **Remarks**
For a  **MultiPage** control, the **Add** method returns a **Page** object. For a **TabStrip**, it returns a **Tab** object. The index value for the first **Page** or **Tab** of a[collection](vbe-glossary.md) is 0, the value for the second **Page** or **Tab** is 1, and so on.
For the  **Controls** collection of an object, the **Add** method returns a control corresponding to the specified _ProgID_. The AddControl event occurs after the control is added.
You can add a control to a user form's  **Controls** collection at design time, but you must use the **Designer** property of the Microsoft Visual Basic for Applications Extensibility Library to do so. The **Designer** property returns the **UserForm** object.
The following syntax will return the  **Text** property of the specified control:



```
userform1.thebox.text
```

If you add a control at run time, you must use the exclamation syntax to reference properties of that control. For example, to return the  **Text** property of a control added at run time, use the following syntax:



```
userform1!thebox.text
```


 **Note**  You can change a control's  **Name** property at[run time](vbe-glossary.md) only if you added that control at run time with the **Add** method.


