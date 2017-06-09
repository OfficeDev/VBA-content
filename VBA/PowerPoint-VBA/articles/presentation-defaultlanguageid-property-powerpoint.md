---
title: Presentation.DefaultLanguageID Property (PowerPoint)
keywords: vbapp10.chm583050
f1_keywords:
- vbapp10.chm583050
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.DefaultLanguageID
ms.assetid: 8568c96c-b997-6a92-e93b-0f3d091383e2
ms.date: 06/08/2017
---


# Presentation.DefaultLanguageID Property (PowerPoint)

Returns or sets the default language of a presentation. Read/write.


## Syntax

 _expression_. **DefaultLanguageID**

 _expression_ A variable that represents a **Presentation** object.


### Return Value

MsoLanguageID


## Remarks

When you set the  **DefaultLanguageID** property for a presentation, you set it for all subsequent new presentations as well.

The value of the  **DefaultLanguageID** property can be one of these **MsoLanguageID** constants.


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

This example sets the default language for the active presentation, and all subsequent new presentations, to German.


```vb
ActivePresentation.DefaultLanguageID = msoLanguageIDGerman
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

