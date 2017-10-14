---
title: About Units of Measure (Visio)
ms.prod: visio
ms.assetid: b6140312-b8e6-0cf2-9fe0-b14e800216bf
ms.date: 06/08/2017
---


# About Units of Measure (Visio)

When you insert fields into text or build formulas, you often specify units of measure for the values you type.

Visio evaluates the result of a formula differently depending on the cell in which you enter the formula. In general, cells that represent shape position, a dimension, or an angle require a number-unit pair that consists of a number and the qualifying units needed to interpret the number. Many other cells don't require units and evaluate to a string, to TRUE or FALSE, or to an index. For example, the same formula that in the FillForegnd cell means color 5 from the drawing's color palette means TRUE (and locks the shape's width) in the LockWidth cell.

Always specify a unit of measure when you enter a formula in a cell that expects a dimensional value. If you do not specify a unit of measure, Visio uses the default unit for that cell, which can be page units, drawing units, type units, duration units, or angular units.


## Units of measure

When indicating units of measure in ShapeSheet formulas, use the abbreviations listed in the following table.



|**To specify these units of measure**|**Use**|**Automation constant**|
|:-----|:-----|:-----|
| Centimeters| cm| **visCentimeters** (69)|
| Ciceros| c| **visCiceros** (54)|
| Date or time| date| **visDate** (40)|
| Degrees| deg| **visDegrees** (81)|
| Didots| d| **visDidots** (53)|
| Elapsed weeks| ew| **visElapsedWeek** (43)|
| Elapsed days| ed| **visElapsedDay** (44)|
| Elapsed hours| eh| **visElapsedHour** (45)|
| Elapsed minutes| em| **visElapsedMin** (46)|
| Elapsed seconds| es| **visElapsedSec** (47)|
| Feet| ft| **visFeet** (66)|
| Inches| in| **visInches** (65)|
| Kilometers| km| **visKilometers** (72)|
| Meters| m| **visMeters** (71)|
| Miles| mi| **visMiles** (68)|
| Millimeters| mm| **visMillimeters** (70)|
| Minutes| '| **visMin** (84)|
| Nautical miles| nm| **visNautMiles** (76)|
| Percent| %| **visPercent** (33)|
| Picas| p| **visPicas** (51)|
| Points| pt| **visPoints** (50)|
| Radians| rad| **visRadians** (83)|
| Seconds| "| **visSec** (85)|
| Yards| yd| **visYards** (75)|

## Compound units of measure

In formulas, you can express units of measure for compound numbers using the abbreviations in the following table. Visio simplifies the results and displays them in the compound units.

For example, if you enter  _45.635°_, Visio displays the equivalent value as 45° 38' 6".



|**To specify units**|**Use this abbreviation**|**Automation constant**|
|:-----|:-----|:-----|
| Ciceros and didots| CICERO/DIDOT| **visCicerosAndDidots** (52)|
| Degrees, minutes, and seconds| °| **visDegreeMinSec** (82)|
| Feet and inches| FEET/INCH| **visFeetAndInches** (67)|
| Picas and points| PICAPOINTS| **visPicasAndPoints** (49)|

## Fractional units of measure

You can specify fractional units of measure in the DrawingScale cell to affect the number of ruler subdivisions that Visio displays in the drawing window. By default, Visio divides distances into tenths when drawing its rulers. If you use fractional units of measure in the DrawingScale cell, Visio divides distance into the following:




- Eighths for  **visInchFrac** and **visMileFrac**
    
- Twelfths for  **visFeetAndInches**
    


Fractional units of measure have no effect in cells other than in the DrawingScale cell.



|**To specify fractional units**|**Use this abbreviation**|**Automation constant**|
|:-----|:-----|:-----|
| Inches in fractions| IN_F| **visInchFrac** (73)|
| Miles in fractions| MI_F| **visMileFrac** (74)|
| Feet and inches| FEET/INCH| **visFeetAndInches** (67)|

## Multidimensional units of measure

In formulas, you can express units of measure for multidimensional numbers using the abbreviations in the following table. Visio simplifies the results and displays them in the multidimensional units.



|**To specify multidimensional units**|**Use this abbreviation**|**Automation constant**|
|:-----|:-----|:-----|
| Acre| ACRE| **visAcre** (36)|
| Centimeters| SQ. CM., SQ CM, CM.^2, CM^2| **visCentimeters** (69)|
| Feet| SQ. FT., SQ FT, FT.^2, FT^2| **visFeet** (66)|
| Hectare| HECTARES, HECTARE, HA., HA| **visHectare** (37)|
| Inches| SQ. IN., SQ IN, IN.^2, IN^2| **visInches** (65)|
| Kilometers| SQ. KM., SQ KM, KM.^2, KM ^2| **visKilometers** (72)|
| Meters| SQ. M., SQ M, M.^2, M ^2| **visMeters** (71)|
| Miles| SQ. MI., SQ MI, MI.^2, MI ^2| **visMiles** (68)|
| Millimeters| SQ. MM., SQ MM, MM.^2, MM ^2| **visMillimeters** (70)|
| Yards| SQ. YD., SQ YD, YD.^2, YD^2| **visYards** (75)|

## Universal strings

In localized versions of Visio, the set of recognized strings changes with the language. If you want your program to work with multiple languages, use the universal strings for units of measure.



|**For**|**Use**|
|:-----|:-----|
| Centimeters| CM|
| Ciceros| C|
| Ciceros and didots| CICERO/DIDOT|
| Date or time| DATE|
| Degrees| DEG|
| Degrees, minutes, seconds| °|
| Didots| D|
| Elapsed week| EW|
| Elapsed day| ED|
| Elapsed hour| EH|
| Elapsed minute| EM|
| Elapsed second| ES|
| Feet| FT|
| Feet and inches| FEET/INCH|
| Inches| IN|
| Inches in fractions| IN_F|
| Kilometers| KM|
| Meters| M|
| Miles| MI|
| Miles in fractions| MI_F|
| Millimeters| MM|
| Minutes| '|
| Nautical miles| NM|
| Percent| %|
| Picas| P|
| Picas and points| PICAPOINTS|
| Points| PT|
| Radians| RAD|
| Seconds| "|
| Yards| YD|

## Implicit units of measure

When Visio parses and stores a number-unit pair, it can use explicit units or implicit units. A number expressed in explicit units always is displayed in the units of measure that were originally entered. A number expressed in implicit units always converts to the equivalent value in the drawing, page, or angular units appropriate for the cell.

For example, suppose you enter the equivalent of 1 inch in cell A using explicit units and in cell B using implicit units, and that both cell A and cell B use drawing units. Next, you change the default units for the page to centimeters. Cell A still displays 1 in., because it uses explicit units that don't change with the defaults. Cell B now displays 2.54 cm, the equivalent value in the default units.

To enter units implicitly, use the following syntax.




```
number [unit, flag]  
```



| _number_|The original value, such as 3.7, 1.7E-4, or 5 1/2.|
| _unit_|The units in which  _number_ originally is expressed.|
| _flag_|The measurement system to use when the implicit-value unit is displayed. See below for values.|
The element  _flag_ is one of the following letters (either uppercase or lowercase) indicating the measurement system that should be used when the implicit-value unit is displayed.



|**_flag_**|**Measurement system**|**Example**|
|:-----|:-----|:-----|
| a, A| Angular| =5[deg,A]|
| d, D| Drawing| =5[in,D]|
| e, E| Duration| =5[eh,E]|
| p, P| Page| =5[in,P]|
| t, T| Type| =5[pt,T]|
Additionally, you can use the implicit units DL, DP, DT, DA, DE for implicit drawing-, page-, text-, angular-, and time-units, respectively. These units assume the associated value is internal units. For example, if the current measurement system is centimeters,  _=2 DL_ would be interpreted as 2 internal units (inches) and displayed as 5.08 cm.

Using the implicit syntax described above, this expression (=2 DL) is equivalent to 2[in,d]. The implicit syntax gives you the choice of how to interpret the value, so you could also specify 2[ft,d], which would be interpreted as 2 feet, and displayed as 60.96 cm. The implicit units DL, DP, DT, DA, and DE are universal, and do not have localized counterparts.


## Default units of measure

Following are the default units of measure along with their equivalent settings in the user interface.



|**Default unit of measure**|**User interface equivalent**|
|:-----|:-----|
| **visDrawingUnits**|The units in the DrawingScale cell of the page or master containing the cell. |
| **visPageUnits**|The units selected in the  **Measurement units** box on the **Page Properties** tab of the **Page Setup** dialog box (on the **Design** tab, click the **Page Setup** arrow).|
| **visTypeUnits**|The units selected in the  **Text** box under **Display** on the **Advanced** tab of the **Visio Options** dialog box.|
| **visAngleUnits**| The units selected in the **Angle** box under **Display** on the **Advanced** tab of the **Visio Options** dialog box.|
| **visDurationUnits**| The units selected in the **Duration** box under **Display** on the **Advanced** tab of the **Visio Options** dialog box.|

