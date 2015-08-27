
# Shape Members (Visio)
Represents anything you can select in a drawing window: a basic shape, a group, a guide, or an object from another application embedded or linked in Microsoft Visio.

 **Last modified:** July 28, 2015

 _**Applies to:** Visio 2013 Preview_

 **In this article**
 [Events](#sectionSection0)
 [Methods](#sectionSection1)
 [Properties](#sectionSection2)



## Events
<a name="sectionSection0"> </a>



|**Name**|**Description**|
|:-----|:-----|
| [BeforeSelectionDelete](3979ee0b-155d-7c16-8141-b2131270b6c6.md)|Occurs before selected objects are deleted.|
| [BeforeShapeDelete](6cbfc832-cdf6-1289-feb4-1b1fcbb3574f.md)|Occurs before a shape is deleted.|
| [BeforeShapeTextEdit](f64b57b6-c92c-dd17-9698-211d9ca2fe83.md)|Occurs before a shape is opened for text editing in the user interface.|
| [CellChanged](d3324bb1-f944-e644-79ef-55022b31fbd2.md)|Occurs after the value changes in a cell in a document.|
| [ConvertToGroupCanceled](f5b312cf-97ab-15c8-3d1c-07edd2023a40.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelConvertToGroup** event.|
| [FormulaChanged](cf141b03-5eaf-bf42-a64f-049f8dec2a14.md)|Occurs after a formula changes in a cell in the object that receives the event.|
| [GroupCanceled](89ce290b-a164-4581-b83d-64d205765aeb.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelGroup** event.|
| [QueryCancelConvertToGroup](18fffdd9-2d6a-90d5-ac34-ce6f3a5e8df6.md)|Occurs before the application converts a selection of shapes to a group in response to a user action in the interface. If any event handler returns  **True**, the operation is canceled.|
| [QueryCancelGroup](a2283176-3584-317e-3645-9e6f3dece076.md)|Occurs before the application groups a selection of shapes in response to a user action in the interface. If any event handler returns  **True**, the operation is canceled.|
| [QueryCancelSelectionDelete](d050cf74-b427-32ef-fe11-77246bb9cf55.md)|Occurs before the application deletes a selection of shapes in response to a user action in the interface. If any event handler returns  **True**, the operation is canceled.|
| [QueryCancelUngroup](de7ffc8b-ad3d-2738-4470-be9d34c43b69.md)|Occurs before the application ungroups a selection of shapes in response to a user action in the interface. If any event handler returns  **True**, the operation is canceled.|
| [SelectionAdded](ca63a476-a7d0-bd27-6c41-5e36b4ef56ed.md)|Occurs after one or more shapes are added to a document.|
| [SelectionDeleteCanceled](10811705-9619-d4d8-80f5-f1fa08eed52f.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelSelectionDelete** event.|
| [ShapeAdded](89e562f4-f3b0-54bd-cbac-515eecb70c97.md)|Occurs after one or more shapes are added to a document.|
| [ShapeChanged](3c31acbc-99c9-f047-7aaa-01eddf4242ea.md)|Occurs after a property of a shape that is not stored in a cell is changed in a document.|
| [ShapeDataGraphicChanged](6c4a9bab-cad0-5f37-a1f8-ca040526e1b5.md)|Occurs after a data graphic is applied to or deleted from a shape.|
| [ShapeExitedTextEdit](ba707fd6-2a5a-65f6-6db4-ed3b5250a103.md)|Occurs after a shape is no longer open for interactive text editing.|
| [ShapeLinkAdded](5cd7431f-18da-184c-7976-06f174cd5f73.md)|Occurs after a shape is linked to a data row.|
| [ShapeLinkDeleted](9233b720-f228-0403-d705-15f5eb39e3b4.md)|Occurs after the link between a shape and a data row is deleted.|
| [ShapeParentChanged](b26b9740-a3bf-1100-0f7b-f34cb03be53c.md)|Occurs after shapes are grouped or a group is ungrouped.|
| [TextChanged](e6516896-de9e-e90f-679b-541c15ab26db.md)|Occurs after the text of a shape is changed in a document.|
| [UngroupCanceled](aca15d4f-c623-471b-80b2-80f6afd2d5c7.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelUngroup** event.|

## Methods
<a name="sectionSection1"> </a>



|**Name**|**Description**|
|:-----|:-----|
| [AddGuide](1155354e-3855-4def-bafb-0d70c933a57a.md)|Adds a guide to a group shape.|
| [AddHyperlink](fbf77a65-88a1-e710-60a2-efde9e7df968.md)|Adds a  **Hyperlink** object to a Microsoft Visio shape.|
| [AddNamedRow](c18380b1-418d-454f-3c90-fa4624291628.md)|Adds a row that has the specified name to the specified ShapeSheet section.|
| [AddRow](8b8dcf65-9b42-b3bf-0da3-61d3fbd02996.md)|Adds a row to a ShapeSheet section at a specified position.|
| [AddRows](8b267f98-e077-0854-a1aa-a0ce8719a2c5.md)|Adds the specified number of rows to a ShapeSheet section at a specified position.|
| [AddSection](64396db4-8361-ece9-b029-24d62ba0a290.md)|Adds a new section to a ShapeSheet spreadsheet.|
| [AddToContainers](ddd3f532-cbbf-3c63-0e02-49f4ea8ca90c.md)|Adds the shape to all underlying containers that allow it as a member.|
| [AutoConnect](36b634be-9943-1aec-f8e0-70467b82eed1.md)|Automatically draws a connection in the specified direction between the shape and another shape on the drawing page.|
| [BoundingBox](68053d27-b7da-9ae7-7557-5d49c8d3e1d6.md)|Returns a rectangle that tightly encloses a shape.|
| [BreakLinkToData](1f4ed559-061e-f016-739c-e760e634dba8.md)|Breaks the link between the shape and the data row to which it is linked in the specified data recordset.|
| [BringForward](88e5c746-e7f2-eddd-35c9-2d9c09c1a602.md)|Brings the shape or selected shapes forward one position in the z-order.|
| [BringToFront](91689605-16b4-eda5-2513-3e04f78fc13e.md)|Brings the shape or selected shapes to the front of the z-order.|
| [CenterDrawing](2ac35f58-2f9d-4139-6477-7e865713c863.md)|Centers a page's, master's, or group's shapes with respect to the extent of the page, master, or group. .|
| [ChangePicture](9193d802-cebd-2bfd-5f8e-400fac36c1a5.md)|Replaces the specified shape's current picture with a new picture.|
| [ConnectedShapes](7f5a0ac9-d0a7-d9fe-9ecb-8e8070ab5951.md)|Returns an array that contains the identifiers (IDs) of the shapes that are connected to the shape.|
| [ConvertToGroup](080db7d0-4283-f8d0-eeca-a6495e6f0536.md)|Converts a selection or an object from another application (a linked or embedded object) to a group.|
| [Copy](2579682b-1dd3-7579-271d-a9994b91a933.md)|Copies a shape to the Clipboard.|
| [CreateSelection](205efec7-afa7-87e8-9c31-22395b283496.md)|Creates various types of  **Selection** objects.|
| [CreateSubProcess](efb26247-777f-d468-a8e6-19a9e9c4a343.md)|Creates and returns a new sub-process page that is linked to the shape.|
| [Cut](fda7a58c-233b-5864-880e-cfa17f20c175.md)|Deletes an object or selection and places it on the Clipboard.|
| [Delete](0960d9e1-b091-ea8c-0724-e10a68d8821a.md)|Deletes an object or selection.|
| [DeleteEx](df4c164d-576a-acce-3322-7f166eb81e4f.md)|Deletes the additional shapes that are associated with the shape, such as connectors and unselected container members, when the shape is deleted.|
| [DeleteRow](892ca523-679d-c707-4aba-e43c011cb718.md)|Deletes a row from a section in a ShapeSheet spreadsheet.|
| [DeleteSection](e07981f3-5efe-f4ad-0517-1af4913c3f70.md)|Deletes a ShapeSheet section.|
| [Disconnect](ece61baa-dfe7-7b61-5c45-49de4cf0e394.md)|Unglues the specified connector end points and offsets them the specified amount from the shapes to which they were joined.|
| [DrawArcByThreePoints](9c00cca4-548e-8f15-1747-897fa5482340.md)|Creates a shape whose path consists of an arc defined by the three points passed as parameters.|
| [DrawBezier](d38b56a5-2366-e418-206f-db39bd8e2c82.md)|Creates a shape whose path is defined by the supplied sequence of Bezier control points.|
| [DrawCircularArc](538ee927-c34a-c697-8bf1-f134355e6060.md)|Creates a new shape whose path consists of a circular arc defined by its center, radius, and start and end angles.|
| [DrawLine](8793104a-0ded-e2ca-54e8-acf987b9c797.md)|Adds a line to the  **Shapes** collection of a group shape.|
| [DrawNURBS](e1209142-3902-3231-a019-f6e091978847.md)|Creates a new shape whose path consists of a single NURBS (nonuniform rational B-spline) segment.|
| [DrawOval](7f561251-251e-6aa9-5291-5919ccce1a9e.md)|Adds an oval (ellipse) to the  **Shapes** collection of a group shape.|
| [DrawPolyline](79bd8e58-097e-6af3-cc52-435acbeececd.md)|Creates a shape whose path is a polyline along a given set of points.|
| [DrawQuarterArc](7bc281ea-eac8-cdb6-ac4b-c71c93a81827.md)|Creates a new shape whose path consists of an elliptical arc defined by the two points and the flag passed in as arguments.|
| [DrawRectangle](2d02da32-0938-b019-0fa0-c4ef07d1a318.md)|Adds a rectangle to the  **Shapes** collection of a page, master, or group.|
| [DrawSpline](02a66d00-2309-b508-7867-90980addb309.md)|Creates a new shape whose path follows a given sequence of points.|
| [Drop](bce5f9d1-8684-0ff5-a4a3-3c2adb972310.md)|Creates one or more new  **Shape**objects by dropping an object onto a receiving object such as a master, drawing page, shape, or group.|
| [DropMany](def60b35-ce19-ec34-9654-355856e26b37.md)|Creates one or more new  **Shape** objects in a group. It returns an array of the IDs of the **Shape** objects it produces.|
| [DropManyU](b3e18874-bb90-334f-e633-3e20133242a1.md)|Creates one or more new  **Shape** objects on a page, in a master, or in a group. It returns an array of the IDs of the **Shape** objects it produces.|
| [Duplicate](a45fd247-e4ad-8149-3656-af9588f076ef.md)|Duplicates an object.|
| [Export](f4051560-8719-ea9c-30eb-33230c95786c.md)|Exports an object from Microsoft Visio to a file format such as .bmp, .dib, .dwg, .dxf, .emf, .emz, .gif, .htm, .jpg, .png, .svg, .svgz, .tif, or .wmf.|
| [FitCurve](9055ee19-a021-35b5-1993-6f00c8a5f859.md)|Reduces the number of geometry segments in a shape or shapes by replacing them with similar spline, arc, and line segments that approximate the paths of the initial segments. Typically, this reduces the number of segments in the shape.|
| [FlipHorizontal](a1f308a7-1f00-9432-ea26-bc1d784b8451.md)|Flips an object horizontally.|
| [FlipVertical](d83d29fb-4292-61c3-b2b4-ba6aed6fe7ad.md)|Flips an object vertically.|
| [GetCustomPropertiesLinkedToData](8a0d783d-f5ee-d6c0-adbd-377cbe65e5f5.md)|Gets the IDs of the shape-data-item (custom property) rows in the Shape Data section of the shape's ShapeSheet spreadsheet linked to the specified data recordset.|
| [GetCustomPropertyLinkedColumn](0d6e3577-d918-1d33-135a-37a3f09f3eaa.md)|Gets the name of the data column linked to the shape data (custom property) row in the shape's ShapeSheet spreadsheet specified by the custom property index.|
| [GetFormulas](51ff9731-802c-2001-c5e6-6f7aeb9d6839.md)|Returns the formulas of many cells.|
| [GetFormulasU](f478abfa-d576-fcb2-5126-464b874355a0.md)|Returns the formulas of many cells.|
| [GetLinkedDataRecordsetIDs](1ce55d6c-02ae-8d5d-f581-b368e830bcf5.md)|Gets the IDs of all the data recordsets that contain data rows linked to the shape.|
| [GetLinkedDataRow](55e578a5-da95-9a5c-3d1d-5cc5edeb57a7.md)|Gets the ID of the data row in the specified data recordset linked to the shape.|
| [GetResults](7380f8b4-ec22-2271-f4ce-19869264ec25.md)|Gets the results or formulas of many cells.|
| [GluedShapes](0c9c551d-ce28-f7c6-4656-8120962e8d34.md)|Returns an array that contains the identifiers of the shapes that are glued to a shape.|
| [Group](fe19f27f-47ad-93ef-1d82-4010d8cb6e47.md)|Groups the objects that are selected in a selection, or it converts a shape into a group.|
| [HasCategory](91115794-31ab-73b1-d1ec-ca249a57a61f.md)|Returns  **True** if the specified category is in the shape categories list.|
| [HitTest](1250ac1d-32f8-d078-3a01-6e2ce045d254.md)|Determines if a given x,y position hits outside, inside, or on the boundary of a shape.|
| [Import](07c858ee-0bbc-5ac1-37be-1e853cdf2361.md)|Imports a file into the current document.|
| [InsertFromFile](894f69fc-65a7-d0a8-a2ae-e56a73843bc2.md)|Adds a linked or embedded object to a page, master, or group.|
| [InsertObject](7abc6e96-6822-7237-b695-36f297a076fc.md)|Adds a new embedded object or ActiveX control to a page, master, or group.|
| [IsCustomPropertyLinked](e75b910f-fb21-3e39-2ca3-ac0913adccf0.md)|Returns whether the shape data (custom property) row in the Shape Data section of the shape's ShapeSheet spreadsheet is linked to a data row in the specified data recordset.|
| [Layout](f70dfdbb-6501-b9b7-3444-7fa35e98637e.md)|Lays out the shapes or reroutes the connectors (or both) for the page, master, group, or selection.|
| [LinkToData](75dd1543-e643-0c7d-a89a-f0dd09d6d323.md)|Links a shape to a data row in a data recordset.|
| [MoveToSubprocess](3688c9d5-5b28-23d7-369a-332649267ffe.md)|Moves the shape to the specified page and drops a replacement shape on the source page, then links it to the target page. Returns the selection of moved shapes on the target page.|
| [Offset](0a82ed87-cc11-77d3-4337-2694f8431a79.md)|Offsets a shape a specified amount.|
| [OpenDrawWindow](5e4106a3-ba72-9e3c-1189-9587d39edd00.md)|Opens a new drawing window that displays a group.|
| [OpenSheetWindow](744b72f5-381a-48fc-407f-20ffe815c54e.md)|Opens a ShapeSheet window for a  **Shape** object.|
| [Paste](ce5892be-b5e7-2ca0-7ee1-aa4e602641a2.md)|Pastes the contents of the Clipboard into an object.|
| [PasteSpecial](0e3a1006-1664-3b60-5d75-d7d4f77d364d.md)|Inserts the contents of the Clipboard, allowing you to control the format of the pasted information and (optionally) establish a link to the source file (for example, a Microsoft Word document).|
| [RemoveFromContainers](b9dbf604-01f0-675a-a0e1-7b30841ec5c5.md)|Removes the shape from all lists and containers of which it is a member.|
| [ReplaceShape](b330a63d-4e3f-0c4d-c38c-6ee806670225.md)|Replaces the specified shape with an instance of the master passed as the first parameter, and returns the new shape.|
| [Resize](ce8e9253-e1bb-e542-30eb-f9ac2e4305da.md)|Resizes the shape by moving shape handles as specified.|
| [ReverseEnds](f2e450fa-0f86-6c90-cf58-8ee57f055577.md)|Reverses an object by flipping it both horizontally and vertically.|
| [Rotate90](1c7d526e-f053-d9f5-232a-61cf12de2c6e.md)|Rotates an object 90 degrees counterclockwise.|
| [SendBackward](9e43cfd9-c2d3-9042-46e3-39e209ae54aa.md)|Moves a shape or selected shapes back one position in the z-order.|
| [SendToBack](faa9cab5-0b2f-8331-e0df-8b4f4be1e69f.md)|Moves the shape or selected shapes to the back of the z-order.|
| [SetBegin](257a6ec4-b9c4-4c42-3c57-6e53c1d4d526.md)|Moves the begin point of a one-dimensional (1-D) shape to the coordinates represented by xPos andyPos.|
| [SetCenter](9a3c0597-c255-44ab-9268-938acd3c5a69.md)|Moves a shape so that its pin is positioned at the coordinates represented by xPos andyPos. .|
| [SetEnd](5f2c7b85-52b3-9147-a989-b2dce61c3493.md)|Moves the endpoint of a one-dimensional (1-D) shape to the coordinates represented by xPos andyPos.|
| [SetFormulas](b2371ff1-4742-e178-3606-133c9c8a1937.md)|Sets the formulas of one or more cells.|
| [SetQuickStyle](aebe80cb-fae9-0be7-e903-882f6eb58b63.md)|Sets the quick style of the specified shape.|
| [SetResults](b5dccaf0-776a-3f0c-ca45-2efff3d3f95b.md)|Sets the results or formulas of one or more cells.|
| [SwapEnds](54096674-eb4f-4f07-a1cf-701faf3b5fae.md)|Swaps the begin and endpoints of a one-dimensional (1-D) shape.|
| [TransformXYFrom](4676e464-83c7-7ff6-e742-becc41436259.md)|Transforms a point expressed in the local coordinate system of one  **Shape** object from an equivalent point expressed in the local coordinate system of another **Shape** object.|
| [TransformXYTo](dc85cf08-0d83-34ff-8389-94a0f5f05c5e.md)|Transforms a point expressed in the local coordinate system of one  **Shape** object to an equivalent point expressed in the local coordinate system of another **Shape** object.|
| [Ungroup](a4ff17b9-6bad-aaf4-ce00-2b529c73f48b.md)|Ungroups a group.|
| [UpdateAlignmentBox](7076ee5f-f536-77ec-a1f7-518195e3e897.md)|Updates the alignment box for a shape.|
| [VisualBoundingBox](http://msdn.microsoft.com/library/6a7d4622-8ba5-c598-4aaa-c6297cb4c008%28Office.15%29.aspx)||
| [XYFromPage](85b04e0b-04e1-a5b5-f6ff-393c57751946.md)|Transforms a point expressed in the local coordinate system of its  **Page** or **Master** object to an equivalent point expressed in the local coordinate system of the **Shape** object.|
| [XYToPage](4a230d63-57a8-3b69-6425-2dca6a2014eb.md)|Transforms a point expressed in the local coordinate system of a  **Shape** object to an equivalent point expressed in the local coordinate system of its **Page** or **Master** object.|

## Properties
<a name="sectionSection2"> </a>



|**Name**|**Description**|
|:-----|:-----|
| [Application](01ad1b62-5a69-9c70-3735-f678a6fa537d.md)|Returns the instance of Microsoft Visio that is associated with an object. Read-only.|
| [AreaIU](a9982cd2-9a91-f5e5-7297-360b6d9a1f29.md)|Returns the area of a  **Shape** object in internal units (square inches). Read-only.|
| [CalloutsAssociated](c1e32bb2-c946-3919-4f6e-064b5be50d6c.md)|Returns an array of  **Long** values that represent the collection of identifiers for all of the callout shapes that are associated with the target shape by a callout relationship. Read-only.|
| [CalloutTarget](4366753a-c8e2-ba85-54fd-9c74cd21d762.md)|Gets or sets the target shape that is associated with the callout shape by a callout relationship. Read/write.|
| [CellExists](479c4d99-0282-3ab0-2e6f-4a17e48adfab.md)|Determines whether a particular ShapeSheet cell exists in the scope of the search. Read-only.|
| [CellExistsU](da26e913-39c5-7af5-194d-3bb5dca76678.md)|Determines whether a particular ShapeSheet cell exists in the scope of the search. Read-only.|
| [Cells](2d90b848-ee2c-d69c-e44e-9c30b04bf776.md)|Returns a  **Cell** object that represents a ShapeSheet cell. Read-only.|
| [CellsRowIndex](7415afcb-9d98-5981-bd33-6ca18116470e.md)|Returns the index of a row to which a cell belongs. Read-only.|
| [CellsRowIndexU](425fbf08-d44c-2631-7400-55620fd429ee.md)|Returns the index of a row to which a cell belongs. Read-only.|
| [CellsSRC](8fb6fd7b-e0ca-c694-3b9d-5390d4192565.md)|Returns a  **Cell** object that represents a ShapeSheet cell identified by section, row, and column indices. Read-only.|
| [CellsSRCExists](7d614820-2a64-c3ee-b61c-a7c0dcfb90c8.md)|Determines whether a ShapeSheet cell exists in the scope of a search. Read-only.|
| [CellsU](bee20521-6515-8a3b-e861-104f7cc71c26.md)|Returns a  **Cell** object that represents a ShapeSheet cell. Read-only.|
| [Characters](dcb7fa7b-61ff-df09-8128-2d1ef4e17770.md)|Returns a  **Characters** object that represents the text of a shape. Read-only.|
| [CharCount](2da9c359-d86c-bdf6-3553-01ded11d9208.md)|Returns the number of characters in an object. Read-only.|
| [ClassID](b3cb2f9c-1247-9799-69f3-5374a112af95.md)|Returns the class ID string of a shape that represents an ActiveX control or an embedded or linked OLE object. Read-only.|
| [Comments](498eca91-beb9-b764-0262-a935e5205710.md)|Returns a  [Comments](7cd0ee53-6b8d-a03b-ecd6-f6f6dda0f2d4.md) object that represents the collection of all the reviewer comments on the shape. Read-only.|
| [Connects](9edaac59-f52e-67ee-6e5a-e11572907785.md)|Returns a  **Connects** collection for a shape, page, or master. Read-only.|
| [ContainerProperties](bc375912-f624-dbdc-3b02-2edf3bf5d8a2.md)|Returns the  ** [ContainerProperties](b94f758f-58f7-f1ef-c03b-761e26c11017.md)** object associated with the shape. Read-only.|
| [ContainingMaster](ca262f68-472e-3412-f620-ca837c40378c.md)|Returns the  **Master** object that contains an object. Read-only.|
| [ContainingMasterID](e194cd7c-d7c0-2c08-a0df-764398efa447.md)|Returns the ID of the  **Master** object that contains an object. Read-only.|
| [ContainingPage](18fe6146-34eb-9369-603b-b3b316aa23d7.md)|Returns the page that contains an object.|
| [ContainingPageID](fd33d0d6-571d-47b5-28a7-6fa4aa671312.md)|Returns the ID of the page that contains an object. Read-only.|
| [ContainingShape](b09bc382-de6c-368e-53bd-c8b01fbc0ae1.md)|Returns the  **Shape** object that contains an object or collection. Read-only.|
| [Data1](ca9dda75-4ae2-70f0-46bd-ff5afbba84fc.md)|Gets or sets the value of the  **Data1** field for a **Shape** object. Read/write.|
| [Data2](59499252-ee61-d158-5ad8-ceece33a8e9e.md)|Gets or sets the value of the  **Data2** field for a **Shape** object. Read/write.|
| [Data3](0d02964d-0296-5142-e7c3-e319ea80c224.md)|Gets or sets the value of the  **Data3** field for a **Shape** object. Read/write.|
| [DataGraphic](09c804fe-d0ec-ac88-6620-1a41fc8a507a.md)|Gets or sets the data graphic master ( **Master** of type **visTypeDataGraphic**) that is associated with the shape. Read/write.|
| [DistanceFrom](2df9e60f-b138-4dde-09ca-af4ee2f6a8d0.md)|Returns the distance from one shape to another, measured between the closest points on the two shapes. Both shapes must be on the same page or in the same master. Read-only.|
| [DistanceFromPoint](262b5814-3b86-c3eb-9526-96ec73836ad6.md)|Returns the distance from a shape to a point. Read-only.|
| [Document](235e9100-dd91-cb6b-01e6-893b4f7acdd8.md)|Gets the  **Document** object that is associated with an object. Read-only.|
| [EventList](513838c2-f00e-06e3-f08b-b23295f7f0d3.md)|Returns the  **EventList** collection of an object or the **EventList** collection that contains an **Event** object. Read-only.|
| [FillStyle](f674da21-deac-4636-608c-c26241a7b125.md)|Returns or sets the fill style for an shape. Read/write.|
| [FillStyleKeepFmt](39fc0329-322e-fd96-2c42-43bdcd170c02.md)|Applies a fill style to an object while preserving local formatting. Read/write.|
| [ForeignData](c7d826fd-b411-3403-a7ec-9fe4e44f41a3.md)|Returns metafile, bitmap, or OLE data for a shape that represents a foreign object. Read-only.|
| [ForeignType](a6cda280-bf0c-b8b0-0750-0ec5fbad90e0.md)|Returns the subtype of a  **Shape** object that represents a foreign object. Read-only.|
| [FromConnects](feb80221-c5d9-f72e-2f79-5153ed375383.md)|Returns a  **Connects** collection of the shapes connected to a shape. Read-only.|
| [GeometryCount](4dffe649-3629-6e3e-bcc0-d860eb1efdbe.md)|Returns the number of Geometry sections for a shape. Read-only.|
| [Help](12784797-c42b-deee-9ae1-6115cd014ac8.md)|Gets or sets the Help string for a shape. Read/write.|
| [Hyperlinks](c1f04a6f-032b-f626-c2e9-6688528052f6.md)|Returns the  **Hyperlinks** collection for a **Shape** object. Read-only.|
| [ID](948982c0-a872-802f-a2d3-69c6539ca3f2.md)|Gets the ID of an object. Read-only.|
| [Index](7fb67e8b-76a7-c2ac-e7aa-89635ca7622c.md)|Gets the ordinal position of a  **Shape** object in the **Shapes** collection. Read-only.|
| [IsCallout](6977e383-41c5-effe-4ac9-88cfc0476813.md)|Indicates whether the shape is a callout shape. Read-only.|
| [IsDataGraphicCallout](dedf6880-e597-8582-12e5-18bfe6286e66.md)|Specifes whether a shape is a data graphic callout. Read-only.|
| [IsOpenForTextEdit](6a4525f2-2532-083d-87f7-390ae7034a65.md)|Indicates whether a shape is currently open for interactive text editing. Read-only.|
| [Language](6c7ab4ca-8813-9cbc-d433-a3991a0b450f.md)|Represents the language ID of the version of the Microsoft Visio instance represented by the parent object. Read/write.|
| [Layer](fb076dda-fa1f-a1fe-c97b-03ba3c7041f0.md)|Returns the layer to which a shape is assigned. Read-only.|
| [LayerCount](0ebcdf53-ebf3-8e26-236f-086f2c9f3c08.md)|Returns the number of layers to which a shape is assigned. Read-only.|
| [LengthIU](11d57f17-5285-6b45-1da1-dc58db087395.md)|Returns the length (perimeter) of the shape in internal units. Read-only.|
| [LineStyle](1d1f2b2e-705d-6547-f6d6-0c5693e426d6.md)|Specifies the line style for an object. Read/write.|
| [LineStyleKeepFmt](4dd4ee1e-5201-1602-39f1-bcda85f96bd0.md)|Applies a line style to an object while preserving local formatting. Read/write.|
| [Master](698e205b-3cfc-2ee1-4fa1-73bc3d018b78.md)|Returns the master from which the  **Shape** object was created. Read-only.|
| [MasterShape](bf710d8b-11f6-145d-a306-658dc23dedbf.md)|If this shape is part of a master instance, returns the shape in the master that this shape inherits from. Read-only.|
| [MemberOfContainers](e8ed57eb-4031-5718-07ce-3641bda00186.md)|Returns an array of  **Long** values that represent the identifiers of the container shapes that include the shape as a member. Read-only.|
| [Name](a0708af0-a813-7539-c43f-049009f1ab62.md)|Specifies the name of an object. Read-only.|
| [NameID](ae658ed9-124f-22f2-53be-5c9b6ebaa382.md)|Returns a unique name for a shape. Read-only.|
| [NameU](1f645016-86a5-f8e4-d5e4-00b8d02cc523.md)|Specifies the universal name of a  **Shape** object. Read/write.|
| [Object](a2e8644a-ac7b-1bb7-9b6b-1515fb9126d2.md)|Returns an  **IDispatch** interface on the ActiveX control or embedded or linked OLE 2.0 object represented by a **Shape** object or an **OLEObject** object. Read-only.|
| [ObjectIsInherited](5bb91e7a-f28e-f169-2e4a-87d46aacdccc.md)|Indicates if a shape represents an ActiveX or OLE object that is inherited from the shape's master. Read-only.|
| [ObjectType](d5711c8e-14a5-6e6b-e8f4-5501a483c9b9.md)|Returns an object's type. Read-only.|
| [OneD](f1511393-4402-ecf8-82a2-2026c56622d0.md)|Determines whether an object behaves as a one-dimensional (1-D) object. Read-only.|
| [Parent](aada0bc1-75e3-8357-3ef9-597a10250860.md)|Determines the parent of a  **Shape** object. Read/write.|
| [Paths](8a179059-7cab-728a-c7b8-a4d8b31476ee.md)|Returns a  **Paths** collection that reports the coordinates of a shape's paths in the coordinate system of the shape's parent. Read-only.|
| [PathsLocal](aa5da0de-ca06-69c0-1fbf-b19ea02d0088.md)|Returns a  **Paths** collection that reports the coordinates of a shape's paths in the shape's local coordinate system. Read-only.|
| [PersistsEvents](6bfa4b18-b4f3-0ac0-de21-ed18600ff473.md)|Indicates whether an object is capable of containing persistent events in its  **EventList** collection. Read-only.|
| [Picture](0ccd2df9-fd84-dee0-0d89-5b7115e418d6.md)|Returns a picture that represents an enhanced metafile (EMF) contained in a master, shape, selection, or page. Read-only.|
| [ProgID](2cd96dd5-7d73-77ea-9e7e-3d1dcd98a21a.md)|Returns the programmatic identifier of a shape that represents an ActiveX control, an embedded object, or linked object. Read-only.|
| [RootShape](c2e91d43-4968-cfee-e53b-4df115d171f6.md)|Returns the top-level shape of an instance if this shape is part of a master instance. Read-only.|
| [RowCount](358f07c8-f72a-134a-53d8-9b70f2400484.md)|Returns the number of rows in a ShapeSheet section. Read-only.|
| [RowExists](bd89deb9-eda3-18d8-6305-bd380d9e649f.md)|Determines whether a ShapeSheet row exists. Read-only.|
| [RowsCellCount](bb9c1990-5ead-e56b-7b09-a49a2b7ad111.md)|Returns the number of cells in a row of a ShapeSheet section. Read-only.|
| [RowType](416b77f1-6cec-de5b-c2b8-c6e5b239c54c.md)|Gets or sets the type of a row in a Geometry, Connection Points, Controls, or Tabs ShapeSheet section. Read/write.|
| [Section](e87823aa-fd7c-e222-417b-a167d2e0898a.md)|Returns the requested  **Section** object belonging to a shape. Read-only.|
| [SectionExists](588a3b17-4831-b7bb-455f-12bc5c3620fc.md)|Determines whether a ShapeSheet section exists for a particular shape. Read-only.|
| [Shapes](83fea91a-19a6-f600-7d03-ba2f03f62d28.md)|Returns the  **Shapes** collection for a page, master, or group. Read-only.|
| [SpatialNeighbors](98069519-d788-c34f-ac25-64bda73324d5.md)|Returns a  **Selection** object that represents the shapes that meet certain criteria in relation to a specified shape. Read-only.|
| [SpatialRelation](7e9f26b5-2887-493f-01c1-5e3900ea8c05.md)|Returns an integer that represents the spatial relationship of one shape to another shape. Both shapes must be on the same page or in the same master. Read-only.|
| [SpatialSearch](360b48b0-783a-7282-b3fe-83f424c393d4.md)|Returns a  **Selection** object whose shapes meet certain criteria in relation to a point that is expressed in the coordinate space of a page, master, or group. Read-only.|
| [Stat](c9d9d8bf-6e64-5231-b870-fcc5de7fdc7b.md)|Returns status information for an object. Read-only.|
| [Style](beba03ba-6926-d2db-4e36-652d05c2925c.md)|Gets or sets the style for a  **Shape** object. Read/write.|
| [StyleKeepFmt](22403064-fa5d-c0cf-8ee7-0f8ae2895d3c.md)|Applies a style to an object while preserving local formatting. Read/write.|
| [Text](5c002c5d-f5ce-7f89-d799-4fc6ccb1a1f7.md)|Returns all of the shape's text. Read/write.|
| [TextStyle](9436ba1b-f792-aed6-3936-b2d88a6dd2ea.md)|Gets or sets the text style for an object. Read/write.|
| [TextStyleKeepFmt](add41319-8b81-a803-46e2-697df37eb731.md)|Applies a text style to an object while preserving local formatting. Read/write.|
| [Type](0d7438d2-e2df-2045-1a2f-608eca530bc1.md)|Returns the type of the object. Read-only.|
| [UniqueID](a82e1175-4536-8919-6531-593d57c3b2f5.md)|Gets, deletes, or makes the GUID that uniquely identifies the shape within the scope of the application. Read-only.|
