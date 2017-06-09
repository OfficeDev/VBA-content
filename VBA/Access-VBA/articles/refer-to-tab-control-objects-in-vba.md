---
title: Refer to Tab Control Objects in VBA
ms.prod: access
ms.assetid: cf090068-7f0b-7ea6-1565-8a05860f9378
ms.date: 06/08/2017
---


# Refer to Tab Control Objects in VBA

You can use a tab control to present several pages of information about a single form. A tab control is useful when your form contains information that can be sorted into two or more categories. 

In most ways, a tab control works like other controls on a form and can be referred to as a member of a form's  **[Controls](controls-object-access.md)** collection. For example, to refer to a tab control named TabControl1 on a form named Form1, you can use the following expression:



```vb
Form1.Controls!TabControl1 

```

However, because the  **Controls** collection is the default collection of the **[Form](form-object-access.md)** object, you do not have to explicitly refer to the **Controls** collection. That is, you can omit the reference to the **Controls** collection from the expression, like this:



```vb
Form1!TabControl1 

```


## Referring to the Pages Collection

A tab control contains one or more pages. Each page in a tab control is referenced as a member of the tab control's  **[Pages](tabcontrol-pages-property-access.md)** collection. Each page in the **Pages** collection can be referred to by either its **[PageIndex](page-pageindex-property-access.md)** property setting (which reflects the page's position in the collection starting with 0), or by the page's **[Name](page-name-property-access.md)** property setting. There is no default collection for the **[TabControl](tabcontrol-object-access.md)** object, so when referring to items in the **Pages** collection by their index value, or to properties of the **Pages** collection, you must explicitly refer to the **Pages** collection.

For example, to change the value of the  **[Caption](page-caption-property-access.md)** property for the first page of a tab control named TabControl1 by referring to its index value in the **Pages** collection, you can use the following statement:




```vb
TabControl1.Pages(0).Caption = "First Page" 

```

Because each page is a member of the form's  **Controls** collection, you can refer to a page solely by its **Name** property without referring to the tab control's name or its **Pages** collection. For example, to change the value of the **Caption** property of a page with its **Name** property set to Page1, use the following statement:




```vb
Page1.Caption = "First Page" 

```


 **Note**  If a user or code changes a page's  **PageIndex** property, the reference to the page's index and the page's position in the page order change. In this case, if you want to maintain an absolute reference to a page, refer to the page's **Name** property.

The  **Pages** collection has one property, **[Count](pages-count-property-access.md)**, that returns the number of pages in a tab control. Note that this property is not a property of the tab control itself, but of its **Pages** collection, so you must explicitly refer to the collection. For example, to determine the number of pages in TabControl1, use the following statement:




```vb
TabControl1.Pages.Count 

```


## Referring to and Changing the Current Page

A tab control's default property is  **[Value](tabcontrol-value-property-access.md)**, which returns an integer that identifies the current page: 0 for the first page, 1 for the second page, and so on. The **Value** property is available only in VBA code or in expressions. By reading the **Value** property at run time, you can determine which page is currently on top. For example, the following statement returns the value for the current page of TabControl1:


```vb
TabControl1.Value 

```


 **Note**  Because the  **Value** property is the default property for a tab control, you do not have to refer to it explicitly. For this reason, you could omit `.Value` from the preceding example.

Setting a tab control's  **Value** property at run time changes the focus to the specified page, making it the current page. For example, the following statement moves the focus to the third page of TabControl1:




```vb
TabControl1 = 2 

```

This is useful if you set a tab control's  **[Style](tabcontrol-style-property-access.md)** property to None (which displays no tabs) and want to use command buttons on the form to determine which page has the focus. To use a command button to display a page, add an event procedure to the button's **[OnClick](commandbutton-onclick-property-access.md)** event that sets the tab control's **Value** property to the integer that identifies the appropriate page.

By using the  **Value** property with the **Pages** collection, you can set properties at run time for the page that is on top. For example, you can hide the current page and all of its controls by setting the page's **[Visible](page-visible-property-access.md)** property to **False**. The following statement hides the current page of TabControl1:




```vb
TabControl1.Pages(TabControl1).Visible = False 

```

Each page in a tab control also has a  **PageIndex** property that specifies the position of a page within the **Pages** collection using the same numbering sequence as the tab control's **Value** property: 0 for the first page, 1 for the second page, and so on. Setting the value of a page's **PageIndex** property changes the order in which pages appear in the tab control. For example, if you wanted to make a page named Page1 the second page, you'd use the following statement:




```vb
Page1.PageIndex = 1 

```

The  **PageIndex** property is more typically set at design time in a page's property sheet. You can also set the page order by right-clicking the border of a tab control and then clicking **Page Order** on the shortcut menu.


## Referring to Controls on a Tab Control Page

The controls you place on a tab control page are part of the same collection as all controls on the form. For this reason, each control on a tab control page must have a name that is unique with respect to all other controls on the same form. You can refer to controls on a tab control page by using the same syntax for controls on a form without a tab control. 


```vb
Forms!Employees!HomePhone 

```

Because each control on a form has its own  **Controls** collection, you can also refer to the controls on a tab control as members of its **Controls** collection. For example, the following code enumerates (lists) all the controls on the tab control of the Employees form.




```vb
Sub ListTabControlControls() 
 
   Dim tabCtl As TabControl 
   Dim ctlCurrent As Control 
 
On Error GoTo ErrorHandler 
 
   ' Return reference to tab control on Employees form. 
   Set tabCtl = Forms!Employees!TabCtl0 
 
   ' List all controls on the tab control in the Debug window. 
   For Each ctlCurrent In tabCtl 
      Debug.Print ctlCurrent.Name 
   Next ctlCurrent 
 
   Set tabCtl = Nothing 
   Set ctlCurrent = Nothing 
 
   Exit Sub 
 
ErrorHandler: 
   MsgBox "Error #: " &; Err.Number &; vbCrLf &; vbCrLf &; Err.Description 
End Sub
```

Additionally, each page on a tab control has its own  **Controls** collection. By using a page's **Controls** collection, you can refer to controls on each page. The following code enumerates the controls for each page of the tab control on the Employees form.




```vb
Sub ListPageControls() 
 
   Dim tabCtl As TabControl 
   Dim pagCurrent As Page 
   Dim ctlCurrent As Control 
   Dim intPageNum As Integer 
 
On Error GoTo ErrorHandler 
 
   ' Return reference to tab control on Employees form. 
   Set tabCtl = Forms!Employees!TabCtl0 
 
   ' List all controls for each page on the tab control in the 
   ' Debug window. 
   For Each pagCurrent In tabCtl.Pages 
      intPageNum = intPageNum + 1 
      Debug.Print "Page " &; intPageNum &; " Controls:" 
      For Each ctlCurrent In pagCurrent.Controls 
         Debug.Print ctlCurrent.Name 
      Next ctlCurrent 
      Debug.Print 
   Next pagCurrent 
 
   Set tabCtl = Nothing 
   Set ctlCurrent = Nothing 
   Set pagCurrent = Nothing 
 
   Exit Sub 
 
ErrorHandler: 
   MsgBox "Error #: " &; Err.Number &; vbCrLf &; vbCrLf &; Err.Description 
End Sub
```


