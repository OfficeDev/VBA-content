---
title: TextRange.LanguageID Property (PowerPoint)
keywords: vbapp10.chm569037
f1_keywords:
- vbapp10.chm569037
ms.prod: powerpoint
api_name:
- PowerPoint.TextRange.LanguageID
ms.assetid: f6744845-5125-239e-65d1-7db8dacdaecd
ms.date: 06/08/2017
---


# TextRange.LanguageID Property (PowerPoint)

Returns or sets the language for the specified text range. Read/write.


## Syntax

 _expression_. **LanguageID**

 _expression_ A variable that represents a **TextRange** object.


### Return Value

MsoLanguageID


## Remarks

The  **LanguageID** property is used for tagging portions of text written in a different language than the **[DefaultLanguageID](presentation-defaultlanguageid-property-powerpoint.md)** property specifies. This allows Microsoft PowerPoint to check spelling and grammar according to the language for each text range. This property is not related to the application interface language.

The value of the  **LanguageID** property can be one of these **MsoLanguageID** constants.


||
|:-----|
|**msoLanguageIDAfrikaans**|
|**msoLanguageIDAlbanian**|
|**msoLanguageIDAmharic**|
|**msoLanguageIDArabic**|
|**msoLanguageIDArabicAlgeria**|
|**msoLanguageIDArabicBahrain**|
|**msoLanguageIDArabicEgypt**|
|**msoLanguageIDArabicIraq**|
|**msoLanguageIDArabicJordan**|
|**msoLanguageIDArabicKuwait**|
|**msoLanguageIDArabicLebanon**|
|**msoLanguageIDArabicLibya**|
|**msoLanguageIDArabicMorocco**|
|**msoLanguageIDArabicOman**|
|**msoLanguageIDArabicQatar**|
|**msoLanguageIDArabicSyria**|
|**msoLanguageIDArabicTunisia**|
|**msoLanguageIDArabicUAE**|
|**msoLanguageIDArabicYemen**|
|**msoLanguageIDArmenian**|
|**msoLanguageIDAssamese**|
|**msoLanguageIDAzeriCyrillic**|
|**msoLanguageIDAzeriLatin**|
|**msoLanguageIDBasque**|
|**msoLanguageIDBelgianDutch**|
|**msoLanguageIDBelgianFrench**|
|**msoLanguageIDBengali**|
|**msoLanguageIDBrazilianPortuguese**|
|**msoLanguageIDBulgarian**|
|**msoLanguageIDBurmese**|
|**msoLanguageIDByelorussian**|
|**msoLanguageIDCatalan**|
|**msoLanguageIDCherokee**|
|**msoLanguageIDChineseHongKong**|
|**msoLanguageIDChineseMacao**|
|**msoLanguageIDChineseSingapore**|
|**msoLanguageIDCroatian**|
|**msoLanguageIDCzech**|
|**msoLanguageIDDanish**|
|**msoLanguageIDDutch**|
|**msoLanguageIDEnglishAUS**|
|**msoLanguageIDEnglishBelize**|
|**msoLanguageIDEnglishCanadian**|
|**msoLanguageIDEnglishCaribbean**|
|**msoLanguageIDEnglishIreland**|
|**msoLanguageIDEnglishJamaica**|
|**msoLanguageIDEnglishNewZealand**|
|**msoLanguageIDEnglishPhilippines**|
|**msoLanguageIDEnglishSouthAfrica**|
|**msoLanguageIDEnglishTrinidad**|
|**msoLanguageIDEnglishUK**|
|**msoLanguageIDEnglishUS**|
|**msoLanguageIDEnglishZimbabwe**|
|**msoLanguageIDEstonian**|
|**msoLanguageIDFaeroese**|
|**msoLanguageIDFarsi**|
|**msoLanguageIDFinnish**|
|**msoLanguageIDFrench**|
|**msoLanguageIDFrenchCameroon**|
|**msoLanguageIDFrenchCanadian**|
|**msoLanguageIDFrenchCotedIvoire**|
|**msoLanguageIDFrenchLuxembourg**|
|**msoLanguageIDFrenchMali**|
|**msoLanguageIDFrenchMonaco**|
|**msoLanguageIDFrenchReunion**|
|**msoLanguageIDFrenchSenegal**|
|**msoLanguageIDFrenchWestIndies**|
|**msoLanguageIDFrenchZaire**|
|**msoLanguageIDFrisianNetherlands**|
|**msoLanguageIDGaelicIreland**|
|**msoLanguageIDGaelicScotland**|
|**msoLanguageIDGalician**|
|**msoLanguageIDGeorgian**|
|**msoLanguageIDGerman**|
|**msoLanguageIDGermanAustria**|
|**msoLanguageIDGermanLiechtenstein**|
|**msoLanguageIDGermanLuxembourg**|
|**msoLanguageIDGreek**|
|**msoLanguageIDGujarati**|
|**msoLanguageIDHebrew**|
|**msoLanguageIDHindi**|
|**msoLanguageIDHungarian**|
|**msoLanguageIDIcelandic**|
|**msoLanguageIDIndonesian**|
|**msoLanguageIDInuktitut**|
|**msoLanguageIDItalian**|
|**msoLanguageIDJapanese**|
|**msoLanguageIDKannada**|
|**msoLanguageIDKashmiri**|
|**msoLanguageIDKazakh**|
|**msoLanguageIDKhmer**|
|**msoLanguageIDKirghiz**|
|**msoLanguageIDKonkani**|
|**msoLanguageIDKorean**|
|**msoLanguageIDLao**|
|**msoLanguageIDLatvian**|
|**msoLanguageIDLithuanian**|
|**msoLanguageIDMacedonian**|
|**msoLanguageIDMalayalam**|
|**msoLanguageIDMalayBruneiDarussalam**|
|**msoLanguageIDMalaysian**|
|**msoLanguageIDMaltese**|
|**msoLanguageIDManipuri**|
|**msoLanguageIDMarathi**|
|**msoLanguageIDMexicanSpanish**|
|**msoLanguageIDMixed**|
|**msoLanguageIDMongolian**|
|**msoLanguageIDNepali**|
|**msoLanguageIDNone**|
|**msoLanguageIDNoProofing**|
|**msoLanguageIDNorwegianBokmol**|
|**msoLanguageIDNorwegianNynorsk**|
|**msoLanguageIDOriya**|
|**msoLanguageIDPolish**|
|**msoLanguageIDPunjabi**|
|**msoLanguageIDRhaetoRomanic**|
|**msoLanguageIDRomanian**|
|**msoLanguageIDRomanianMoldova**|
|**msoLanguageIDRussian**|
|**msoLanguageIDRussianMoldova**|
|**msoLanguageIDSamiLappish**|
|**msoLanguageIDSanskrit**|
|**msoLanguageIDSerbianCyrillic**|
|**msoLanguageIDSerbianLatin**|
|**msoLanguageIDSesotho**|
|**msoLanguageIDSimplifiedChinese**|
|**msoLanguageIDSindhi**|
|**msoLanguageIDSlovak**|
|**msoLanguageIDSlovenian**|
|**msoLanguageIDSorbian**|
|**msoLanguageIDSpanish**|
|**msoLanguageIDSpanishArgentina**|
|**msoLanguageIDSpanishBolivia**|
|**msoLanguageIDSpanishChile**|
|**msoLanguageIDSpanishColombia**|
|**msoLanguageIDSpanishCostaRica**|
|**msoLanguageIDSpanishDominicanRepublic**|
|**msoLanguageIDSpanishEcuador**|
|**msoLanguageIDSpanishElSalvador**|
|**msoLanguageIDSpanishGuatemala**|
|**msoLanguageIDSpanishHonduras**|
|**msoLanguageIDSpanishModernSort**|
|**msoLanguageIDSpanishNicaragua**|
|**msoLanguageIDSpanishPanama**|
|**msoLanguageIDSpanishParaguay**|
|**msoLanguageIDSpanishPeru**|
|**msoLanguageIDSpanishPuertoRico**|
|**msoLanguageIDSpanishUruguay**|
|**msoLanguageIDSpanishVenezuela**|
|**msoLanguageIDSutu**|
|**msoLanguageIDSwahili**|
|**msoLanguageIDSwedish**|
|**msoLanguageIDSwedishFinland**|
|**msoLanguageIDSwissFrench**|
|**msoLanguageIDSwissGerman**|
|**msoLanguageIDSwissItalian**|
|**msoLanguageIDTajik**|
|**msoLanguageIDTamil**|
|**msoLanguageIDTatar**|
|**msoLanguageIDTelugu**|
|**msoLanguageIDThai**|
|**msoLanguageIDTibetan**|
|**msoLanguageIDTraditionalChinese**|
|**msoLanguageIDTsonga**|
|**msoLanguageIDTswana**|
|**msoLanguageIDTurkish**|
|**msoLanguageIDTurkmen**|
|**msoLanguageIDUkrainian**|
|**msoLanguageIDUrdu**|
|**msoLanguageIDUzbekCyrillic**|
|**msoLanguageIDUzbekLatin**|
|**msoLanguageIDVenda**|
|**msoLanguageIDVietnamese**|
|**msoLanguageIDWelsh**|
|**msoLanguageIDXhosa**|
|**msoLanguageIDZulu**|
|**msoLanguageIDPortuguese**|

## Example

This example sets the language for the specified text in shape one to US English.


```vb
ActivePresentation.Slides(1).Shapes(1).TextFrame.TextRange.LanguageID = msoLanguageIDEnglishUS
```


## See also


#### Concepts


[TextRange Object](textrange-object-powerpoint.md)

