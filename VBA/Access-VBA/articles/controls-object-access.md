---
title: Controls Object (Access)
keywords: vbaac10.chm10176
f1_keywords:
- vbaac10.chm10176
ms.prod: access
api_name:
- Access.Controls
ms.assetid: 26771888-86e8-28c3-6668-f793474cbb5b
ms.date: 06/08/2017
---


# Controls Object (Access)

The  **Controls** collection contains all of the controls on a form, report, or subform, within another control, or attached to another control. The **Controls** collection is a member of a **[Form](http://msdn.microsoft.com/library/72ef9219-142b-b690-b696-3eba9a5d4522%28Office.15%29.aspx)**, **[Report](report-object-access.md)**, and **[SubForm](subform-object-access.md)** objects.


## Remarks

You can enumerate individual controls, count them, and set their properties in the  **Controls** collection. For example, you can enumerate the **Controls** collection of a particular form and set the **Height** property of each control to a specified value.

It is faster to refer to the  **Controls** collection implicitly, as in the following examples, which refer to a control called NewData on a form named OrderForm. Of the following syntax examples, `Me!NewData` is the fastest way to refer to the control.




```
Me!NewData               ' Or Forms!OrderForm!NewData.
```




```
Me![New Data]            ' Use if control name contains space.
```




```
Me("NewData")            ' Performance is slightly slower.
```

You can also refer to an individual control by referring explicitly to the  **Controls** collection.




```
Me.Controls!NewData      ' Or Forms!OrderForm.Controls!NewData.
```




```
Me.Controls![New Data]
```




```
Me.Controls("NewData")
```

Additionally, you can refer to a control by its index in the collection. The  **Controls** collection is indexed beginning with zero.




```
Me(0)                    ' Refer to first item in collection.
```




```
Me.Controls(0)
```


 **Note**  You can use the  **Me** keyword to represent a form or report within code only if you're referring to the form or report from code within the form module or report module. If you're referring to a form or report from a standard module or a different form's or report's module, you must use the full reference to the form or report.

To work with the controls on a section of a form or report, use the  **Section** property to return a reference to a **Section** object. Then refer to the **Controls** collection of the **Section** object.

Two types of  **Control** objects, the tab control and option group control, have **Controls** collections that can contain multiple controls. The **Controls** collection belonging to the option group control contains any option button, check box, toggle button, or label controls in the option group.

The tab control contains a  **[Pages](http://msdn.microsoft.com/library/e77c8d31-1cb7-d647-6faa-2eb234ce0708%28Office.15%29.aspx)** collection, which is a special type of **Controls** collection. The **Pages** collection contains **[Page](http://msdn.microsoft.com/library/6351b0ea-bd07-5ee6-ea20-0d410e09d939%28Office.15%29.aspx)** objects. **Page** objects are also controls. The **[ControlType](http://msdn.microsoft.com/library/dec0d7dd-f0e1-a8d7-f026-9ff128481d2a%28Office.15%29.aspx)** property constant for a **Page** control is **acPage**. A **Page** object, in turn, has its own **Controls** collection, which contains all the controls on an individual page.

Other  **Control** objects have a **Controls** collection that can contain an attached label. These controls include the text box, option group, option button, toggle button, check box, combo box, list box, command button, bound object frame, and unbound object frame controls.


## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/c8650732-ffee-830b-9d9d-571a09af3a4c%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/531c1674-4782-aa8f-64f5-0493a29886e3%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/aac9c15e-0a29-c324-299c-b692883c25ed%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/89ec2e2d-ebab-c6db-9810-75f83c712c4d%28Office.15%29.aspx)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
[Controls Object Members](http://msdn.microsoft.com/library/e82f868f-9c18-7845-d476-f6399c441e97%28Office.15%29.aspx)
