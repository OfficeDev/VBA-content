
# CustomControl Object (Access)

 **Last modified:** July 28, 2015

When setting the properties of an ActiveX control, you may need or prefer to use the control's custom properties dialog box. This custom properties dialog box provides an alternative to the list of properties in the Microsoft Access property sheet for setting ActiveX control properties in Design view.

## Remarks


 **Note**  This information only applies to ActiveX controls in a Microsoft Access database environment.

 **Two Ways to Set Properties**

The reason for the custom properties dialog box is that not all applications that use ActiveX controls provide a property sheet like the one in Microsoft Access. The custom properties dialog box provides an interface for setting key control properties regardless of the interface supplied by the hosting application.

For some ActiveX control properties, you can choose either of these two locations to set the property:


- The Microsoft Access property sheet.
    
- The ActiveX control's custom properties dialog box.
    
In some cases, the custom properties dialog box is the only way to set a property in Design view. This is usually the situation when the interface needed to set a property doesn't work inside the Microsoft Access property sheet. For example, the  **GridFont** property for the Calendar control has a number of arguments; you can't set more than one argument per property in the Microsoft Access property sheet.

 **Finding the Custom Properties Dialog Box**

Not all ActiveX controls provide a custom properties dialog box. To see whether a control provides this custom properties dialog box, look for the  **Custom** property in the Microsoft Access property sheet for this control. If the list of properties contains the name **Custom**, then the control provides the custom properties dialog box.

After you click the  **Custom** property box in the Microsoft Access property sheet, click the **Build** button to the right of the property box to display the control's custom properties dialog box, often presented as a tabbed dialog box. Choose the tab that contains the interface for setting the properties that you want to set.

 **Using the Custom Properties Dialog Box**

After you make changes on one tab, you can often apply those changes immediately by clicking the  **Apply** button (if provided). You can click other tabs to set other properties as needed. To approve all changes made in the custom properties dialog box, click the **OK** button. To return to the Microsoft Access property sheet without changing any property settings, click the **Cancel** button.

You can also view the custom properties dialog box by clicking the  **Properties** subcommand of the ActiveX control **Object** command (for example, **Calendar Control Object**) on the  **Edit** menu, or by clicking this same subcommand on the shortcut menu for the ActiveX control. In addition, some properties in the Microsoft Access property sheet for the ActiveX control, like the **GridFontColor** property of the Calendar control, have a **Build** button to the right of the property box. When you click the **Build** button, the custom properties dialog box is displayed, with the appropriate tab selected (for example, **Colors**).


## See also


#### Concepts


 [Access Object Model Reference](2de134a4-6c5c-d2a3-8377-f4dd973ba650.md)
#### Other resources


 [CustomControl Object Members](3093550b-7994-fb58-044c-90e8da535f9d.md)
