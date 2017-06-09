---
title: SubForm.LinkMasterFields Property (Access)
keywords: vbaac10.chm11928
f1_keywords:
- vbaac10.chm11928
ms.prod: access
api_name:
- Access.SubForm.LinkMasterFields
ms.assetid: b5be0557-a75c-dacc-e842-b9196edf37ce
ms.date: 06/08/2017
---


# SubForm.LinkMasterFields Property (Access)

You can use the  **LinkMasterFields** property (along with the **LinkChildFields** property) together to specify how Microsoft Access links records in a form or report to records in a subform, subreport, or embedded object, such as a chart. If these properties are set, Microsoft Access automatically updates the related record in the subform when you change to a new record in a main form. Read/write **String**.


## Syntax

 _expression_. **LinkMasterFields**

 _expression_ A variable that represents a **SubForm** object.


## Remarks

You can set the  **LinkChildFields** and **LinkMasterFields** properties for the subform, subreport, or embedded object as follows:


- The  **LinkChildFields** property. Enter the name of one or more linking fields in the subform, subreport, or embedded object.
    
- The  **LinkMasterFields** property. Enter the name of one or more linking fields or controls in the main form or report.
    
You can use the Subform/Subreport Field Linker to set these properties by clicking the  **Build** button to the right of the property box in the property sheet.

The properties can only be set in Design view or during the  **Open** event of a form or report.

The fields or controls you use to set these properties don't need to have the same names, but they must contain the same kind of data and have the same or a compatible data type and field size. For example, an AutoNumber field is compatible with a Number field if the  **FieldSize** property for the Number field is set to **Long Integer**.

You can use the name of a control (including the name of a calculated control) to set the  **LinkMasterFields** property, but you can't use the name of a control to set the **LinkChildFields** property. If you want to use a calculated value as the link for a subform, subreport, or embedded object, define a calculated field in the child object's underlying query and set the **LinkChildFields** property to the field.

When you specify more than one field or control name for these property settings, you must enter the same number of fields or controls for each property setting and separate the names with a semicolon (;).

When you create a subform or subreport by dragging a form or report from the Database window onto another form or report or by using the Form Wizard, Microsoft Access automatically sets the  **LinkChildFields** and **LinkMasterFields** properties under the following conditions:


- Both the main form or report and the child object are based on tables, and a relationship between those tables has been defined with the  **Relationships** command. Microsoft Access uses the fields that relate the two tables as the linking fields.
    
- The main form or report is based on a table with a primary key, and the subform or subreport is based on a table or query that contains a field with the same name and the same or a compatible data type as the primary key. Microsoft Access uses the primary key from the main object's underlying table and the identically named field from the child object's underlying table or query as the linking fields.
    

 **Note**  The linking fields don't have to be included in the main object or in the child object. As long as they are contained in the objects' underlying tables or queries, you can use the fields to link the objects. When you use a wizard, Microsoft Access automatically includes the linking fields.


## See also


#### Concepts


[SubForm Object](subform-object-access.md)

