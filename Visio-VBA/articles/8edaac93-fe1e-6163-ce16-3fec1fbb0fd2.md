
# Master Members (Visio)
Represents a master in a stencil.

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
| [BeforeMasterDelete](46b455db-9165-0ed4-ebf3-15e1794313be.md)|Occurs before a master is deleted from a document.|
| [BeforeSelectionDelete](e2f86944-6ca2-6535-ee08-889af9694fd6.md)|Occurs before selected objects are deleted.|
| [BeforeShapeDelete](21921e16-3e05-6232-ed89-76217b76149f.md)|Occurs before a shape is deleted.|
| [BeforeShapeTextEdit](1d39001d-6efa-7d58-1eaa-f6c2531e2018.md)|Occurs before a shape is opened for text editing in the user interface.|
| [CellChanged](53323234-8e92-de8b-65b8-20eb867748dd.md)|Occurs after the value changes in a cell in a document.|
| [ConnectionsAdded](15c772fe-d5fb-901e-f1d4-1d3eb0cb7c64.md)|Occurs after connections have been established between shapes.|
| [ConnectionsDeleted](dc043012-d653-8f37-372e-f7532047aa81.md)|Occurs after connections between shapes have been removed.|
| [ConvertToGroupCanceled](b585e434-fd81-93ae-92a6-5cc1d21c1afa.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelConvertToGroup** event.|
| [FormulaChanged](6d2a9ab6-778e-cbba-0b63-f7d38116dc85.md)|Occurs after a formula changes in a cell in the object that receives the event.|
| [GroupCanceled](ec87e679-2b8f-de85-81b9-ccb4a9df7ae2.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelGroup** event.|
| [MasterChanged](922120cc-56e0-143b-7a8b-754bc368af47.md)|Occurs after properties of a master are changed and propagated to its instances.|
| [MasterDeleteCanceled](a682fab6-1fc9-65ba-83a1-408d048ee81e.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelMasterDelete** event.|
| [QueryCancelConvertToGroup](c23d7ed0-0ad4-fa20-4b4f-fa453716fbd5.md)|Occurs before the application converts a selection of shapes to a group in response to a user action in the interface. If any event handler returns  **True**, the operation is canceled.|
| [QueryCancelGroup](37625c3b-49e2-d3ba-5270-2dcb65062f08.md)|Occurs before the application groups a selection of shapes in response to a user action in the interface. If any event handler returns  **True**, the operation is canceled.|
| [QueryCancelMasterDelete](33690e0f-821e-42cd-ec52-3ade1a1ceadc.md)|Occurs before the application deletes a master in response to a user action in the interface. If any event handler returns  **True**, the operation is canceled.|
| [QueryCancelSelectionDelete](c85569ca-b802-7a7e-6b24-d89852d2d0bc.md)|Occurs before the application deletes a selection of shapes in response to a user action in the interface. If any event handler returns  **True**, the operation is canceled.|
| [QueryCancelUngroup](44ce0f2e-e877-ec7f-b5ec-1c3ff3b9749a.md)|Occurs before the application ungroups a selection of shapes in response to a user action in the interface. If any event handler returns  **True**, the operation is canceled.|
| [SelectionAdded](c004e65c-1770-edf1-9d1e-a1a02a15fc39.md)|Occurs after one or more shapes are added to a document.|
| [SelectionDeleteCanceled](87ecdfcb-616f-0b47-bfa4-216ef456deaa.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelSelectionDelete** event.|
| [ShapeAdded](d679f866-c939-faff-d8da-cdddb2131054.md)|Occurs after one or more shapes are added to a document.|
| [ShapeChanged](e1a2a7bf-bfe1-acfc-ae04-308f9fda7c0a.md)|Occurs after a property of a shape that is not stored in a cell is changed in a document.|
| [ShapeDataGraphicChanged](74eb2604-bcb2-0cba-37e2-50ad896991ca.md)|Occurs after a data graphic is applied to or deleted from a shape.|
| [ShapeExitedTextEdit](401f6d32-d1fb-f019-52a3-d553b8516ecf.md)|Occurs after a shape is no longer open for interactive text editing.|
| [ShapeParentChanged](37de7351-969b-5b24-fde2-e4473e92b344.md)|Occurs after shapes are grouped or a group is ungrouped.|
| [TextChanged](9224577c-a285-c26f-60be-3adbf3285ef3.md)|Occurs after the text of a shape is changed in a document.|
| [UngroupCanceled](0bbe537e-9bae-62a9-7e29-aea71ab3c8f9.md)|Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelUngroup** event.|

## Methods
<a name="sectionSection1"> </a>



|**Name**|**Description**|
|:-----|:-----|
| [AddGuide](7beba614-244b-f559-50c7-5156ca4510b1.md)|Adds a guide to a master.|
| [BoundingBox](23ef5e08-fcb4-93e6-2ed5-818d34f99a8e.md)|Returns a rectangle that tightly encloses the shapes of a master.|
| [CenterDrawing](1bf660a3-30eb-4a0b-fcea-66d0e0574ae0.md)|Centers a page's, master's, or group's shapes with respect to the extent of the page, master, or group. .|
| [Close](69607a2c-dc59-d170-733a-3557a996a67e.md)|Closes a master.|
| [CreateSelection](52db8b1b-e253-549f-c3ba-d661fa7b675e.md)|Creates various types of  **Selection** objects.|
| [CreateShortcut](e808ba09-b85a-52bb-55e2-ced37f426a3b.md)|Creates a shortcut for a master.|
| [DataGraphicDelete](aa84af70-975c-3747-1976-b872a6c2fa36.md)|Deletes the  **Master** of type **visTypeDataGraphic** from the **Masters** collection of the document.|
| [Delete](8f71e69e-7d7d-7732-738c-ad262b0367ae.md)|Deletes an object.|
| [DrawArcByThreePoints](d2df1c41-8164-d941-21a8-2e1b00de6199.md)|Creates a shape whose path consists of an arc defined by the three points passed as parameters.|
| [DrawBezier](4cbefabf-530e-2c6d-0751-45efa2bb0980.md)|Creates a shape whose path is defined by the supplied sequence of Bezier control points.|
| [DrawCircularArc](f9557127-8470-2968-3056-0e295cd05633.md)|Creates a new shape whose path consists of a circular arc defined by its center, radius, and start and end angles.|
| [DrawLine](c29810a2-c1eb-82cc-ab19-236a89baf7b0.md)|Adds a line to the  **Shapes** collection of a master.|
| [DrawNURBS](7dcfef4a-5b69-9a8b-3966-9b3089bdaac3.md)|Creates a new shape whose path consists of a single NURBS (nonuniform rational B-spline) segment.|
| [DrawOval](092a59d6-1b43-c094-e2ae-480ee7b32b73.md)|Adds an oval (ellipse) to the  **Shapes** collection of a master.|
| [DrawPolyline](a599e60c-ccd6-ce6b-7e54-f65f8500447d.md)|Creates a shape whose path is a polyline along a given set of points.|
| [DrawQuarterArc](6c728c0c-8317-6114-70e5-e5cb68a5729f.md)|Creates a shape whose path consists of an elliptical arc defined by the two points and the flag passed in as arguments.|
| [DrawRectangle](e41ec411-ccd7-0fe6-f560-cf3934d18b59.md)|Adds a rectangle to the  **Shapes** collection of a page, master, or group.|
| [DrawSpline](a255978d-5479-ba7e-4520-0a8d18390ea6.md)|Creates a new shape whose path follows a given sequence of points.|
| [Drop](13abc8fc-7b3c-98cf-3965-3ac7b3d15e85.md)|Creates one or more new  **Shape**objects by dropping an object onto a receiving object such as a master, drawing page, shape, or group.|
| [DropMany](fb0ef035-c1ce-5703-e2e8-0f9b63b186bf.md)|Creates one or more new  **Shape** objects in a master. It returns an array of the IDs of the **Shape** objects it produces.|
| [DropManyU](467356ff-d2d9-71d9-d533-b88099bf2fae.md)|Creates one or more new  **Shape** objects on a page, in a master, or in a group. It returns an array of the IDs of the **Shape** objects it produces.|
| [Export](212bcc8e-646c-37df-9387-4605b72b6edd.md)|Exports an object from Microsoft Visio to a file format such as .bmp, .dib, .dwg, .dxf, .emf, .emz, .gif, .htm, .jpg, .png, .svg, .svgz, .tif, or .wmf.|
| [ExportIcon](8b13f92f-537a-1efb-b2b0-531a8054e89b.md)|Exports the icon for a  **Master** object to a named file or the Clipboard.|
| [GetFormulas](09ee33a3-41fc-3ac2-4f5e-1e857f685049.md)|Returns the formulas of many cells.|
| [GetFormulasU](d5a419e2-9630-a724-af44-f2f1b0166c80.md)|Returns the formulas of many cells.|
| [GetResults](d532a2ed-2246-8c90-2d77-df2df05a395f.md)|Gets the results or formulas of many cells.|
| [Import](3b13025f-1a83-0dcf-41e1-03cd83dfc7be.md)|Imports a file into the current document.|
| [ImportIcon](886d724d-9d02-ab6f-8049-80fa04f8caec.md)|Imports the icon for a  **Master** object from a named file.|
| [InsertFromFile](5a24e289-675a-d08b-36f7-0cfaedac5aaf.md)|Adds a linked or embedded object to a page, master, or group.|
| [InsertObject](7b663eef-ed40-486b-2b5b-e7c7066c2300.md)|Adds a new embedded object or ActiveX control to a page, master, or group.|
| [Layout](acab2dc3-daf8-57c2-cbf8-edf647a12a09.md)|Lays out the shapes and/or reroutes the connectors for the page, master, group, or selection.|
| [Open](3f14f3b2-1cfb-ccf9-b344-7fbf80ae9a26.md)|Opens an existing master so that it can be edited.|
| [OpenDrawWindow](5f17d4a0-6b5d-bb85-cff7-047bd18ff1ee.md)|Opens a new drawing window that displays a master.|
| [OpenIconWindow](5e2b2437-05cc-4855-e0bb-96b097c98d3c.md)|Opens an icon window that shows a master's icon.|
| [Paste](ee8a4c79-9a10-d852-70d3-4856627efb8a.md)|Pastes the contents of the Clipboard into an object.|
| [PasteSpecial](6ca1994b-feb4-6b0d-c2c4-8a134eb284f1.md)|Inserts the contents of the Clipboard, allowing you to control the format of the pasted information and (optionally) establish a link to the source file (for example, a Microsoft Word document).|
| [PasteToLocation](c5c94265-23ee-5516-525d-ed3f34d2e7bf.md)|Pastes a shape to the specified location.|
| [ResizeToFitContents](982fa4c4-014c-319d-a73e-f6bbc28f16e8.md)|Resizes the page, or the master's page, to fit tightly around the shapes or master that are on it.|
| [SetFormulas](fb419eb5-6bd3-cfc7-d358-cef9e68dddbf.md)|Sets the formulas of one or more cells.|
| [SetResults](6be7dd71-55a7-777c-e1b7-8f41c028e843.md)|Sets the results or formulas of one or more cells.|
| [VisualBoundingBox](http://msdn.microsoft.com/library/478d636f-e741-cf6b-3e16-b5faf70a9f14%28Office.15%29.aspx)||

## Properties
<a name="sectionSection2"> </a>



|**Name**|**Description**|
|:-----|:-----|
| [AlignName](5df055eb-ddb1-2d2a-1d94-93781960b3a9.md)|Gets or sets the position of a master name in a stencil window. Read/write.|
| [Application](88b2fd6e-8f7e-3caa-5316-35a6a0060793.md)|Returns the instance of Microsoft Visio that is associated with an object. Read-only.|
| [BaseID](85ca3c0d-5015-b303-7102-144768acb6a8.md)|Returns a base ID for a master. Read-only.|
| [Connects](72c01ae0-9134-d384-b860-dbb333a498fe.md)|Returns a  **Connects** collection for a shape, page, or master. Read-only.|
| [DataGraphicHidden](adcf1867-8541-785b-d8ad-dd44583473b9.md)|Hides or displays a data graphic in the  **Data Graphics** task pane in the Microsoft Visio user interface. Read/write.|
| [DataGraphicHidesText](c1a08780-0873-3d8b-1872-edc8a6515840.md)|Displays or hides the text of a shape or of the primary shape in a selection when a data graphic is applied to the shape or to the selection. Read/write.|
| [DataGraphicHorizontalPosition](d9c98a41-ffc0-152e-2150-0915bd38bcac.md)|Gets or sets the default horizontal callout position for members of the  **GraphicItems** collection of the **Master** object of type **visTypeDataGraphic**. Read/write.|
| [DataGraphicShowBorder](203d631c-d838-ea0a-f67a-39de513e738e.md)|Gets or sets whether a border is displayed around the graphic items contained in the data graphic that are in default positions. Read/write.|
| [DataGraphicVerticalPosition](779f360e-7529-7fe6-87e7-f41cc9334c83.md)|Gets or sets the default vertical callout position for members of the  **GraphicItems** collection of the **Master** object of type **visTypeDataGraphic**. Read/write.|
| [Document](b95000f8-67df-99f4-bbfc-020b14ae73b8.md)|Gets the  **Document** object that is associated with an object. Read-only.|
| [EditCopy](69d13b8f-c5af-d9c9-b92e-00e6eadf660a.md)|Returns a master that is open for editing and originally copied from this master. Read-only.|
| [EventList](02a4d80f-fbc6-6491-5f8b-ce98dd5c2aa8.md)|Returns the  **EventList** collection of an object or the **EventList** collection that contains an **Event** object. Read-only.|
| [GraphicItems](615b4909-c248-3ebd-c7c1-53151464cee9.md)|Returns the  **GraphicItems** collection that the master contains. Read-only.|
| [Hidden](d28eb888-75d7-bbd2-e6d3-3e412cca85d4.md)|Hides or shows a master on a stencil or a style in the user interface. Read/write.|
| [Icon](2e9c7bbd-d8fd-e932-4a6b-bbd845aef4f0.md)|Returns the icon contained in a master. Read/write.|
| [IconSize](c6516b30-642d-1e61-22b4-f95d6c47a8ec.md)|Gets or sets the size of a master icon. Read/write.|
| [IconUpdate](3978c650-47d5-e961-53c2-d99dd4c2ca7c.md)|Determines whether a master icon is updated manually or automatically. Read/write.|
| [ID](9064e708-f939-9522-b8f7-24488d780bc0.md)|Gets the ID of an object. Read-only.|
| [Index](48a90dee-ce11-ef81-e58a-e4a3cdb899dc.md)|Gets the ordinal position of a  **Master** object in the **Masters** collection. Read-only.|
| [IndexInStencil](3c2c12c4-0233-4aa3-c3d7-a3613bb391ad.md)|Gets or sets the index of a master or master shortcut object within its stencil. Read/write.|
| [IsChanged](8e557655-3e16-3e96-99a2-b097fa6abd75.md)|Indicates whether a master has changed since it was opened. Read-only.|
| [Layers](6c78d629-506c-54aa-e0cc-7fd807cdfffb.md)|Returns the  **Layers** collection of an object. Read-only.|
| [MatchByName](4edb0e5f-7e87-c66d-b842-318cd0eba5d5.md)|Determines how the application decides if a document master is already present when an instance of a master is dropped on the drawing page. It allows changes made to a document master to apply to new instances of the master, even if the instances are dragged from a stand-alone stencil file. Read/write.|
| [Name](66ca8cd6-c784-efbb-a2b6-2b0fcce7d5b1.md)|Specifies the name of an object. Read-only.|
| [NameU](87530cb6-5ac1-55c4-9210-9989c5f589c3.md)|Specifies the universal name of a  **Master** object. Read/write.|
| [NewBaseID](bee59c61-06de-ebb9-a8aa-599fc788e4e1.md)|Generates a new base ID for a master. Read-only.|
| [ObjectType](958b08f3-a52b-d6cb-2360-ca2ddf758e3c.md)|Returns an object's type. Read-only.|
| [OLEObjects](b51fbdc2-a236-4733-5a2e-b8e75d457d64.md)|Returns the  **OLEObjects** collection of a master. Read-only.|
| [OneD](917f8cfc-a2fc-7572-936a-69956d139131.md)|Determines whether an object behaves as a one-dimensional (1-D) object. Read-only.|
| [Original](33636aa0-2b2b-9edb-3738-ac193eaab212.md)|Returns the original master that produced this open master. Read-only.|
| [PageSheet](8ec4d38a-79fe-018d-9bc8-3a9c0221f018.md)|Returns the page sheet (an object that represents the ShapeSheet spreadsheet) of a master. Read-only.|
| [PatternFlags](cf7d5e0e-802e-c65b-6260-eaf68dfe6eb4.md)|Determines whether a master behaves as a custom pattern. Read/write.|
| [PersistsEvents](6840a242-85d8-b93e-242b-90c584a9b422.md)|Indicates whether an object is capable of containing persistent events in its  **EventList** collection. Read-only.|
| [Picture](b882b05f-5e54-aab8-db88-1e66cf825581.md)|Returns a picture that represents an enhanced metafile (EMF) contained in a master, shape, selection, or page. Read-only.|
| [Prompt](7467c2dd-5cf6-0af0-bc4d-522889d69707.md)|Gets or sets the prompt string for a master or master shortcut. Read/write.|
| [Shapes](56db5c02-9b55-dfe1-993b-c23e93e84577.md)|Returns the  **Shapes** collection for a page, master, or group. Read-only.|
| [SpatialSearch](d71b05b7-32e1-d3c8-668e-6e96595acd59.md)|Returns a  **Selection** object whose shapes meet certain criteria in relation to a point that is expressed in the coordinate space of a page, master, or group. Read-only.|
| [Stat](1cc33fe9-e317-ab3d-1ce1-a7f8c619c4f2.md)|Returns status information for an object. Read-only.|
| [Type](4688ff5d-2f9a-fcaf-6a73-0aa50562b24a.md)|Returns the type of the  **Master** object. Read-only.|
| [UniqueID](99d0655c-da5c-9d0a-4936-2fa24821e097.md)|Returns the unique ID of a master. Read-only.|
