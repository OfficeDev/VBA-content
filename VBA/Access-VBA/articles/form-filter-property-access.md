---
title: Form.Filter Property (Access)
keywords: vbaac10.chm13346
f1_keywords:
- vbaac10.chm13346
ms.prod: access
api_name:
- Access.Form.Filter
ms.assetid: 5eb49f82-8519-981c-a663-9862736ac95f
ms.date: 06/08/2017
---


# Form.Filter Property (Access)

You can use the  **Filter** property to specify a subset of records to be displayed when a filter is applied to a form, reportquery, or table. Read/write **String**.


## Syntax

 _expression_. **Filter**

 _expression_ A variable that represents a **Form** object.


## Remarks

If you want to specify a server filter within a Microsoft Access project (.adp) for data located on a server, use the  **ServerFilter** property.

The  **Filter** property is a string expression consisting of a WHERE clause without the WHERE keyword. For example, the following Visual Basic code defines and applies a filter to show only customers from the USA:




```vb
Me.Filter = "Country = 'USA'" 
Me.FilterOn = True
```


 **Note**  Setting the  **Filter** property has no effect on the ADO **Filter** property.

You can use the  **Filter** property to save a filter and apply it at a later time. Filters are saved with the objects in which they are created. They are automatically loaded when the object is opened, but they aren't automatically applied.

When a new object is created, it inherits the  **RecordSource**, **Filter**, **OrderBy**, and **OrderByOn** properties of the table or query it was created from.

To apply a saved filter to a form, query, or table, you can click  **Apply Filter** on the toolbar, click **Apply Filter/Sort** on the **Records** menu, or use a macro or Visual Basic to set the **FilterOn** property to **True**. For reports, you can apply a filter by setting the **FilterOn** property to Yes in the report's property sheet.

The  **Apply Filter** button indicates the state of the **Filter** and **FilterOn** properties. The button remains disabled until there is a filter to apply. If an existing filter is currently applied, the **Apply Filter** button appears pressed in.

To apply a filter automatically when a form is opened, specify in the  **OnOpen** event property setting of the form either a macro that uses the ApplyFilter action or an event procedure that uses the **ApplyFilter** method of the **DoCmd** object.

You can remove a filter by clicking the pressed-in  **Apply Filter** button, clicking **Remove Filter/Sort** on the **Records** menu, or using Visual Basic to set the **FilterOn** property to **False**.

When the  **Filter** property is set in form Design view, Microsoft Access does not attempt to validate the SQL expression. If the SQL expression is invalid, an error occurs when the filter is applied.

 **Link provided by:**
![Community Member Icon](images/8b9774c4-6c97-470e-b3a2-56d8f786444c.png) The[UtterAccess](http://www.utteraccess.com) community


- [Forms: Populate Controls/Text Boxes Based on Combobox Selection](http://www.utteraccess.com/wiki/index.php/Forms:_Populate_Controls/Text_Boxes_Based_on_Combobox_Selection)
    

## About the Contributors
<a name="AboutContributors"> </a>

UtterAccess is the premier Microsoft Access wiki and help forum. Click here to join. 


## See also
<a name="AboutContributors"> </a>


#### Concepts


[Form Object](form-object-access.md)

