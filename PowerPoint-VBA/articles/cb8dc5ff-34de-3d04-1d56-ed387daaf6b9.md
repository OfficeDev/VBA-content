
# TextRange Members (PowerPoint)
Contains the text that's attached to a shape, and properties and methods for manipulating the text.

 **Last modified:** July 28, 2015


## Methods



|**Name**|**Description**|
|:-----|:-----|
| [AddPeriods](597592ba-6c26-7645-33b8-19ccb375a098.md)|Adds a period at the end of each paragraph in the specified text.|
| [ChangeCase](a14edb26-7ec3-5fb5-7590-cd67a75c1f03.md)|Changes the case of the specified text.|
| [Characters](019c15d3-349d-ab10-7448-70bf81176150.md)|Returns a  ** [TextRange](7c234107-c423-7ec9-e8bd-a82cc3b345de.md)** object that represents the specified subset of text characters. For information about counting or looping through the characters in a text range, see the ** [TextRange](7c234107-c423-7ec9-e8bd-a82cc3b345de.md)** object.|
| [Copy](c8d1edf7-68ef-aaa4-e2db-717263df8dd3.md)|Copies the specified object to the Clipboard.|
| [Cut](9be71668-1486-0466-f87b-47792d402102.md)|Deletes the specified object and places it on the Clipboard.|
| [Delete](2baac89b-d7b1-2209-b17f-b65b71b5aed4.md)|Deletes the specified  **TextRange** object.|
| [Find](24186821-3a0a-efd5-c35a-8b553e00f92b.md)|Finds the specified text in a text range, and returns a  ** [TextRange](7c234107-c423-7ec9-e8bd-a82cc3b345de.md)** object that represents the first text range where the text is found. Returns **Nothing** if no match is found.|
| [InsertAfter](2af4e134-c205-fbe6-a006-3fc1ca8d6a50.md)|Appends a string to the end of the specified text range. Returns a  **TextRange** object that represents the appended text. When used without an argument, this method returns a zero-length string at the end of the specified range.|
| [InsertBefore](fbadcecd-a31b-8c8d-3281-63d70286bcff.md)|Appends a string to the beginning of the specified text range. Returns a  **TextRange** object that represents the appended text. When used without an argument, this method returns a zero-length string at the end of the specified range.|
| [InsertDateTime](b1f6c2db-2524-f76e-eee2-8f177b08dcde.md)|Inserts the date and time in the specified text range. Returns a  **TextRange** object that represents the inserted text.|
| [InsertSlideNumber](07489db8-9db1-9721-845a-7895ad316aca.md)|Inserts the slide number of the current slide into the specified text range. Returns a  **TextRange** object that represents the slide number.|
| [InsertSymbol](a424e011-1bfe-f690-cbc0-604f89718831.md)|Returns a  ** [TextRange](7c234107-c423-7ec9-e8bd-a82cc3b345de.md)** object that represents a symbol inserted into the specified text range.|
| [Lines](8e9f344b-2e74-5a9d-06e8-3e6ff9ca6bd0.md)|Returns a  ** [TextRange](7c234107-c423-7ec9-e8bd-a82cc3b345de.md)** object that represents the specified subset of text lines. For information about counting or looping through the lines in a text range, see the ** [TextRange](7c234107-c423-7ec9-e8bd-a82cc3b345de.md)** object.|
| [LtrRun](5c6787cc-d37c-8aec-b49e-12418291e006.md)|Sets the direction of text in a text range to read from left to right.|
| [Paragraphs](5062eccf-4db2-692f-501e-b7d214181171.md)|Returns a  **TextRange** object that represents the specified subset of text paragraphs.|
| [Paste](4bbaa1f3-206e-2009-11f0-5abde24517c6.md)|Pastes the text on the Clipboard into the specified text range, and returns a  **TextRange** object that represents the pasted text.|
| [PasteSpecial](97bfd298-f8e8-32f0-b05c-6a93ed651954.md)|Replaces the text range with the contents of the Clipboard in the format specified. |
| [RemovePeriods](66562cc7-e25b-e110-000e-c01b62caf761.md)|Removes the period at the end of each paragraph in the specified text.|
| [Replace](046d1c3d-fd3e-7871-e31e-6529b77fcd60.md)|Finds specific text in a text range, replaces the found text with a specified string, and returns a  **TextRange** object that represents the first occurrence of the found text. Returns **Nothing** if no match is found.|
| [RotatedBounds](33a4393e-3b87-a4d3-0e16-8230a4beabe3.md)|Returns the coordinates of the vertices of the text bounding box for the specified text range.|
| [RtlRun](eb474c9b-d789-f741-9ba9-0514f0a5b0be.md)|Sets the direction of text in a text range to read from right to left.|
| [Runs](0bf2724a-0735-bd79-31e5-894d1320b9b2.md)|Returns a  **TextRange** object that represents the specified subset of text runs. A text run consists of a range of characters that share the same font attributes.|
| [Select](cd6fb1ba-ac49-a7d8-2777-fda2ce2746a4.md)|Selects the specified object.|
| [Sentences](c3640cb8-f78a-2934-bbe0-506cb8d2534c.md)|Returns a  **TextRange** object that represents the specified subset of text sentences.|
| [TrimText](8566ed9d-c73a-d699-bcb7-edcd9a375afe.md)|Returns a  **TextRange** object that represents the specified text minus any trailing spaces.|
| [Words](b8cd8dca-bf10-1041-dd9e-adc04b2df42d.md)|Returns a  ** [TextRange](7c234107-c423-7ec9-e8bd-a82cc3b345de.md)** object that represents the specified subset of text words.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [ActionSettings](7a66ca28-d6b9-2066-4c88-a04888d5ba1e.md)|Returns an  ** [ActionSettings](8914c203-6b8d-fa80-16ad-7015595657b7.md)**object that contains information about what action occurs when the user clicks or moves the mouse over the specified shape or text range during a slide show. Read-only.|
| [Application](53b4f6fc-7e1b-7045-e59d-eec668a75d3e.md)|Returns an  ** [Application](978c2b99-4271-b953-4283-73b5f3d96f41.md)**object that represents the creator of the specified object.|
| [BoundHeight](8f3b9947-5ee3-260d-3d44-0ad2da422724.md)|Returns the height (in points) of the text bounding box for the specified text frame. Read-only.|
| [BoundLeft](2641e084-6b6e-ff6e-c6a6-27cb84cbd4dd.md)|Returns the distance (in points) from the left edge of the text bounding box for the specified text frame to the left edge of the slide. Read-only.|
| [BoundTop](cfc3baec-06c4-da2f-a233-afcb5301302a.md)|Returns the distance (in points) from the top of the of the text bounding box for the specified text frame to the top of the slide. Read-only.|
| [BoundWidth](409d1c66-8956-cdd0-2328-f1cbe584f778.md)|Returns the width (in points) of the text bounding box for the specified text frame. Read-only.|
| [Count](9c514376-18ef-1eac-661a-c1fc46514b32.md)|Returns the number of objects in the specified collection. Read-only.|
| [Font](234c8843-3c0d-a425-0173-02c3910ba400.md)|Returns a  ** [Font](ad62daaa-01a5-36cc-5451-e0da0134ac95.md)**object that represents character formatting. Read-only.|
| [IndentLevel](3ba39fc4-6fc4-62ca-0e87-a7605d6c8bea.md)|Returns or sets the the indent level for the specified text as an integer from 1 to 5, where 1 indicates a first-level paragraph with no indentation. Read/write.|
| [LanguageID](f6744845-5125-239e-65d1-7db8dacdaecd.md)|Returns or sets the language for the specified text range. Read/write.|
| [Length](4eb64830-f8e4-5226-57c1-80df7f4bd39f.md)|Returns the length of the specified text range, in characters. Read-only.|
| [ParagraphFormat](41d3f0f3-70e3-ad1a-efcb-de849d4a03d4.md)|Returns a  ** [ParagraphFormat](15d495cf-16e2-5cfb-e99c-a551876e3a8a.md)**object that represents paragraph formatting for the specified text. Read-only.|
| [Parent](303cc0cf-8c1c-60af-648e-fea4d25abb36.md)|Returns the parent object for the specified object.|
| [Start](1e37b589-842e-b03b-09eb-a19ce03f6a72.md)|Returns the position of the first character in the specified text range relative to the first character in the shape that contains the text. Read-only.|
| [Text](c80c8b19-73e2-0820-abd6-c44f4b2644b2.md)|Returns or sets a  **String** that represents the text contained in the specified object. Read/write.|
