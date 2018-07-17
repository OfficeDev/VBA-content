---
title: Control Object (Access)
keywords: vbaac10.chm10174
f1_keywords:
- vbaac10.chm10174
ms.prod: access
api_name:
- Access.Control
ms.assetid: ce2362e5-4390-590e-06c0-6f27e8d988cd
ms.date: 06/08/2017
---


# Control Object (Access)

The  **Control** object represents a control on a form, report, or section, within another control, or attached to another control.


## Remarks

All controls on a form or report belong to the  **Controls** collection for that **Form** or **Report** object. Controls within a particular section belong to the **Controls** collection for that section. Controls within a tab control or option group control belong to the **Controls** collection for that control. A label control that is attached to another control belongs to the **Controls** collection for that control.

When you refer to an individual  **Control** object in the **Controls** collection, you can refer to the **Controls** collection either implicitly or explicitly.




```
' Implicitly refer to NewData control in Controls 
' collection. 
Me!NewData
```




```
' Use if control name contains space. 
Me![New Data]
```




```
' Performance slightly slower. 
Me("NewData")
```




```
' Refer to a control by its index in the controls 
' collection. 
Me(0)
```




```
' Refer to a NewData control by using the subform 
' Controls collection. 
Me.ctlSubForm.Controls!NewData
```




```
' Explicitly refer to the NewData control in the 
' Controls collection. 
Me.Controls!NewData
```




```
Me.Controls("NewData")
```




```
Me.Controls(0)
```


 **Note**  You can use the  **Me** keyword to represent a **Form** or **Report** object within code only if you're referring to the form or report from code within the class module. If you're referring to a form or report from a standard module or a different form's or report's module, you must use the full reference to the form or report.

Each  **Control** object is denoted by a particular intrinsic constant. For example, the intrinsic constant **acTextBox** is associated with a text box control, and **acCommandButton** is associated with a command button. The constants for the various Microsoft Access controls are set forth in the control's **ControlType** property.

To determine the type of an existing control, you can use the  **ControlType** property. However, you don't need to know the specific type of a control in order to use it in code. You can simply represent it with a variable of data type **Control**.

If you do know the data type of the control to which you are referring, and the control is a built-in Microsoft Access control, you should represent it with a variable of a specific type. For example, if you know that a particular control is a text box, declare a variable of type  **TextBox** to represent it, as shown in the following code.




```
Dim txt As TextBox 
Set txt = Forms!Employees!LastName 

```


 **Note**  If a control is an ActiveX control, then you must declare a variable of type  **Control** to represent it; you cannot use a specific type. If you're not certain what type of control a variable will point to, declare the variable as type **Control**.

The option group control can contain other controls within its  **Controls** collection, including option button, check box, toggle button, and label controls.

The tab control contains a  **[Pages](http://msdn.microsoft.com/library/e77c8d31-1cb7-d647-6faa-2eb234ce0708%28Office.15%29.aspx)** collection, which is a special type of **Controls** collection. The **Pages** collection contains **[Page](http://msdn.microsoft.com/library/6351b0ea-bd07-5ee6-ea20-0d410e09d939%28Office.15%29.aspx)** objects, which are controls. Each **Page** object in turn contains a **Controls** collection, which contains all of the controls on that page.

Other  **Control** objects have a **Controls** collection that can contain an attached label. These controls include the text box, option group, option button, toggle button, check box, combo box, list box, command button, bound object frame, and unbound object frame controls.


## Methods



|**Name**|
|:-----|
|[Dropdown](http://msdn.microsoft.com/library/45957d42-3e81-f7eb-9579-e5e75c833f59%28Office.15%29.aspx)|
|[Move](http://msdn.microsoft.com/library/fd52e497-642f-72a9-af64-971d8c888e71%28Office.15%29.aspx)|
|[Requery](http://msdn.microsoft.com/library/95f68520-7bbc-6627-0702-477b839f98c5%28Office.15%29.aspx)|
|[SetFocus](http://msdn.microsoft.com/library/21e2a6d1-7dd9-92ae-a6a6-72ed67dbc61d%28Office.15%29.aspx)|
|[SizeToFit](http://msdn.microsoft.com/library/421b93c4-b648-a7d7-5e0c-845672d4dab8%28Office.15%29.aspx)|
|[Undo](http://msdn.microsoft.com/library/d2c2d6ee-7086-db63-c471-03530cf7f2ab%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/b46574ca-6159-c34a-befd-7d390bdc39f9%28Office.15%29.aspx)|
|[BottomPadding](http://msdn.microsoft.com/library/df4e0be6-aec9-3e04-c273-3fa0d5d8c026%28Office.15%29.aspx)|
|[Column](http://msdn.microsoft.com/library/7e4594a5-113e-9fe0-fb96-04b1ee7e798d%28Office.15%29.aspx)|
|[Controls](http://msdn.microsoft.com/library/81b01d02-c346-8750-cc8a-4623f24219f6%28Office.15%29.aspx)|
|[Form](http://msdn.microsoft.com/library/86612c78-65f8-dc56-77da-d031502822f7%28Office.15%29.aspx)|
|[GridlineColor](http://msdn.microsoft.com/library/0966f9d9-70a0-cdd9-fc89-7bf9239e63b6%28Office.15%29.aspx)|
|[GridlineStyleBottom](http://msdn.microsoft.com/library/d53fb6e0-3613-095f-a52d-747819fc5601%28Office.15%29.aspx)|
|[GridlineStyleLeft](http://msdn.microsoft.com/library/fe8829f8-bad9-2b34-f613-22b65b3325d4%28Office.15%29.aspx)|
|[GridlineStyleRight](http://msdn.microsoft.com/library/1d0bf3f0-97d4-d88f-047f-270985520e45%28Office.15%29.aspx)|
|[GridlineStyleTop](http://msdn.microsoft.com/library/d2a5a630-d6ff-75ae-5921-9c2953d8e9c6%28Office.15%29.aspx)|
|[GridlineWidthBottom](http://msdn.microsoft.com/library/f51c8d07-a9ce-ce99-622b-7f35290812fb%28Office.15%29.aspx)|
|[GridlineWidthLeft](http://msdn.microsoft.com/library/e148bfb1-a668-f2e3-ef0b-f243e943bef3%28Office.15%29.aspx)|
|[GridlineWidthRight](http://msdn.microsoft.com/library/12df6aff-9e00-35ff-47ca-40be9622ee8a%28Office.15%29.aspx)|
|[GridlineWidthTop](http://msdn.microsoft.com/library/5d3d0d5a-3c72-26fc-66d2-1b7af9768b36%28Office.15%29.aspx)|
|[HorizontalAnchor](http://msdn.microsoft.com/library/e81daacc-3c0b-608c-aea1-e01bc162b6b3%28Office.15%29.aspx)|
|[Hyperlink](http://msdn.microsoft.com/library/f27ac8cc-f5ba-cbc5-4153-7b24ce199679%28Office.15%29.aspx)|
|[ItemData](http://msdn.microsoft.com/library/5eb23c40-566e-33bb-9b73-0ecc701ea5e5%28Office.15%29.aspx)|
|[ItemsSelected](http://msdn.microsoft.com/library/348bc66f-4274-df2e-fdec-d36f678fd7de%28Office.15%29.aspx)|
|[Layout](http://msdn.microsoft.com/library/c290a3e7-bba1-0d67-1e82-a53a4b7b2588%28Office.15%29.aspx)|
|[LayoutID](http://msdn.microsoft.com/library/1cf53242-e9e8-dc87-907a-788036844f4c%28Office.15%29.aspx)|
|[LeftPadding](http://msdn.microsoft.com/library/42354a61-958a-7c9a-6091-a1884c77ef8a%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/b1e31997-1b99-0476-eda8-afef8975420b%28Office.15%29.aspx)|
|[Object](http://msdn.microsoft.com/library/bfa11d67-ef96-128f-ef0d-efc555b51b5d%28Office.15%29.aspx)|
|[ObjectVerbs](http://msdn.microsoft.com/library/e94a1718-0cd7-6d4a-b319-03b180233824%28Office.15%29.aspx)|
|[OldValue](http://msdn.microsoft.com/library/eb805182-2e02-f134-2515-12b3ca564154%28Office.15%29.aspx)|
|[Pages](http://msdn.microsoft.com/library/fd4ea2c0-ea8c-51a0-a012-8ba5848d3516%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/e85b37ce-72cd-2326-4f64-6613dde9d2b9%28Office.15%29.aspx)|
|[Properties](http://msdn.microsoft.com/library/418b9ddb-b7d3-813c-7806-9ae9904175d7%28Office.15%29.aspx)|
|[Report](http://msdn.microsoft.com/library/1c1f4703-bda7-de97-eb13-830238a5170a%28Office.15%29.aspx)|
|[RightPadding](http://msdn.microsoft.com/library/3c9d50a2-42e7-f292-a3bc-42bed689fcef%28Office.15%29.aspx)|
|[Selected](http://msdn.microsoft.com/library/80477eda-78aa-6cdd-062f-dd9caac816c6%28Office.15%29.aspx)|
|[SmartTags](http://msdn.microsoft.com/library/2f8b1435-31d4-4388-614c-4f26544eed7c%28Office.15%29.aspx)|
|[TopPadding](http://msdn.microsoft.com/library/52197046-2042-fc96-f72d-d81413546e9e%28Office.15%29.aspx)|
|[VerticalAnchor](http://msdn.microsoft.com/library/ce4da8b9-aaad-85db-fd3a-490fbd87c380%28Office.15%29.aspx)|

## See also


#### Other resources


[Control Object Members](http://msdn.microsoft.com/library/c6f2ed0f-f8e1-d56e-22a5-a365b64b7754%28Office.15%29.aspx)
[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
