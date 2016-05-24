
# DisplayControl Property

 **Last modified:** December 30, 2015

 _ **Applies to:** Access 2013 | Access 2016_

You can use the  **DisplayControl** property in table Design view to specify the default control you want to use for displaying a field.


## Setting

You can set the  **DisplayControl** property in the table's property sheet in table Design view by clicking the **Lookup** tab in the **Field Properties** section.

This property contains a drop-down list of the available controls for the selected field. For fields with a Text or Number data type, this property can be set to Text Box, List Box, or Combo Box. For fields with a Yes/No data type, this property can be set to Check Box, Text Box, or Combo Box.


## Remarks

When you select a control for this property, any additional properties needed to configure the control are also displayed on the  **Lookup** tab.

Setting this property and any related control type properties will affect the field display in both Datasheet view and Form view. The field is displayed by using the control and control property settings set in table Design view. If a field had its  **DisplayControl** property set in table Design view and you drag it from the field list in form Design view, Microsoft Access copies the appropriate properties to the control's property sheet.

