
# TextRange2 Members (Office)
Represents the text frame in a  **Shape** or **ShapeRange** objects.

 **Last modified:** July 28, 2015


## Methods



|**Name**|**Description**|
|:-----|:-----|
| [AddPeriods](3c706017-1d13-6f15-a111-7e05647ed5d4.md)|Adds period (.) punctuation to the right side of the text contained in TextRange2 object for left-to-right languages and on the left side for right-to-left languages.|
| [ChangeCase](c59fd653-02e6-0e9a-a7a7-3806a75fc146.md)|Changes the case of a  **TextRange2** object to one of the values in the **MsoTextChangeCase** enumeration.|
| [Copy](ad247113-31b4-424c-b62d-ab427081b46c.md)|Copies a  **TextRange2** object.|
| [Cut](64f09c8a-a4cb-2770-0efc-a79e19f51e05.md)|Removes a portion or all of the text from a range of text.|
| [Delete](876c315d-4b97-1489-9d12-f1f0f1fecb74.md)|Deletes a  **TextRange2** object.|
| [Find](ad5bc61a-a7f1-485a-0fc8-a3bd6707f956.md)|Searches a  **TextRange2** object for a subset of text.|
| [InsertAfter](67ebed89-0090-98cb-882a-c1eaf701d182.md)|Inserts text to the right of the existing text in the  **TextRange2** object.|
| [InsertBefore](f75709bd-1239-1736-9cb0-0092dd720860.md)|Inserts text to the left of the existing text in the  **TextRange2** object.|
| [InsertChartField](3ced5d2c-b3a4-6bf3-3d3c-b1145e7b9eab.md)|Inserts a field into the body of a data label in a chart. |
| [InsertSymbol](74642859-0d84-5de9-494a-a58b9d93de88.md)|Inserts a symbol from the specified font set into the range of text represented by the  **TextRange2** object.|
| [Item](8ea600ad-31b0-5b6e-6391-c954bbc97245.md)|Gets the range of text specified by the index number from the  **TextRange2** object.|
| [LtrRun](519f23a7-550f-b155-9e49-113933ce0434.md)|Returns a  **TextRange2** object that represents the specified subset of left-to-right text runs. A text run consists of a range of characters that share the same font attributes.|
| [Paste](b22e0628-f137-9018-5b50-a804c07933dd.md)|Pastes the contents of the Clipboard into the  **TextRange2** object.|
| [PasteSpecial](79f88454-2f95-ea10-6ec4-5fb78ca8036d.md)|Replaces the text range with the contents of the Clipboard in the format specified. If the paste succeeds, this method returns a  **TextRange2** object including the text range that was pasted.|
| [RemovePeriods](581c9be1-94f4-d73b-c274-16b2575cac60.md)|Removes all period (.) punctuation from the text in the  **TextRange2** object.|
| [Replace](e14f0ad0-3b9c-d9f5-a13d-d3bbdcae50e1.md)|Finds specific text in a text range, replaces the found text with a specified string, and returns a  **TextRange2** object that represents the first occurrence of the found text. Returns **Nothing** if no match is found.|
| [RotatedBounds](e8e1b0dc-426f-c804-e91a-8cd5345186de.md)|Gets the coordinates of the vertices of the text bounding box for the specified text range. Read-only.|
| [RtlRun](3c18d756-768f-292f-31c0-efbcf5800f63.md)|Returns a  **TextRange2** object that represents the specified subset of right-to-left text runs. A text run consists of a range of characters that share the same font attributes.|
| [Select](17c6e340-6522-d6b2-f4b7-137dacb666da.md)|Selects the  **TextRange2** object.|
| [TrimText](304c6b05-febf-4ebe-2d26-326ffff613b6.md)|Returns a **TextRange2** object that represents the specified text that has the whitespace removed.|

## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Application](3883561f-229b-92f9-eaea-83f00ac33f06.md)|Used without an object qualifier, this property returns an  **Application**object that represents the current instance of the Microsoft Office application. Used with an object qualifier, this property returns an  **Application** object that represents the creator of the **TextRange2** object. When used with an OLE Automation object, it returns the object's application. Read-only.|
| [BoundHeight](078ff3f3-745d-05f7-c81e-f78f603a45df.md)|Gets the height, in points, of the text bounding box for the specified text. Read-only.|
| [BoundLeft](8af6b9b9-4ecf-c127-87db-b87cabe9184b.md)|Gets the left coordinate, in points, of the text bounding box for the specified text. Read-only.|
| [BoundTop](b225b65e-04a0-1938-9520-ea71eed13b04.md)|Gets the top coordinate, in points, of the text bounding box for the specified text. Read-only.|
| [BoundWidth](a5668c93-0206-c26f-41bc-771c1ceef7e6.md)|Gets the width, in points, of the text bounding box for the specified text. Read-only.|
| [Characters](9b264529-e538-4480-e629-822d5056f148.md)|Read-only.|
| [Count](3bb6408f-acc0-05cb-ef45-9f9a4bae4ebc.md)|Gets a  **Long** indicating the number of items in the **TextRange2** collection. Read-only.|
| [Creator](5158865d-13b7-960c-4bdc-8c0d5711a6c4.md)|Gets a 32-bit integer that indicates the application in which the **TextRange2** object was created. Read-only.|
| [Font](005fa6bf-2dd5-32ec-18e8-30ff6260e55d.md)|Returns a  **Font**object that represents character formatting for the  **TextRange2** object. Read-only.|
| [LanguageID](3fc73136-6107-ae4c-7f18-0c6ec944591a.md)|Gets or sets the  **MsoLanguageID** value of the **TextRange2** object. Read/write.|
| [Length](3b873f1f-5120-3832-1d34-b8c0f668bba3.md)|Get a Long that represents the length of a text range. Read-only.|
| [Lines](5e20f089-c345-e22a-c136-483d13f7f658.md)|Returns a TextRange2 object that represents the specified subset of text lines. Read-only.|
| [MathZones](277aa819-d717-e2f5-5bc7-607abfce20a4.md)|Sets the starting point and length of a math zone within a text range. Read-only|
| [ParagraphFormat](68818c1a-9503-4f3f-77e1-28ac6b049c3b.md)|Returns a  **ParagraphFormat**object that represents paragraph formatting for the specified text. Read-only.|
| [Paragraphs](15479f9e-f261-7ea6-0460-861ccea08440.md)|Gets a  **TextRange2** object that represents the specified subset of text paragraphs. Read-only.|
| [Parent](692dc869-1525-ffa5-023d-83cea9cec19e.md)|Gets the  **Parent** object for the **TextRange2** object. Read-only.|
| [Runs](5398a676-67a9-315f-193c-62602f27c377.md)|Gets a  **TextRange2** object that represents the specified subset of text runs. A text run consists of a range of characters that share the same font attributes. Read-only.|
| [Sentences](236196a7-97b3-f3d5-b483-c42bc60bd9ed.md)|Returns a  **TextRange2** object that represents the specified subset of text sentences. Read-only.|
| [Start](53f7731d-2e98-28c7-981e-64b2e6616636.md)|Gets a  **Long** value indicating the starting point of the specified text range. Read-only.|
| [Text](b071a9fb-f657-0bc2-9c07-6b1ef604a525.md)|Gets or sets a  **String** value that represents the text in a text range. Read/write.|
| [Words](bab78b31-ebd6-649e-0b05-5b21552f8f22.md)|Gets a  **TextRange2** object that represents the specified subset of text words. Read-only.|
