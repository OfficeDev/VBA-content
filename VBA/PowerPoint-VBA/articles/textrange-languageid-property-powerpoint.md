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
|<strong>msoLanguageIDAfrikaans</strong>|
|
<strong>msoLanguageIDAlbanian</strong>|
|
<strong>msoLanguageIDAmharic</strong>|
|
<strong>msoLanguageIDArabic</strong>|
|
<strong>msoLanguageIDArabicAlgeria</strong>|
|
<strong>msoLanguageIDArabicBahrain</strong>|
|
<strong>msoLanguageIDArabicEgypt</strong>|
|
<strong>msoLanguageIDArabicIraq</strong>|
|
<strong>msoLanguageIDArabicJordan</strong>|
|
<strong>msoLanguageIDArabicKuwait</strong>|
|
<strong>msoLanguageIDArabicLebanon</strong>|
|
<strong>msoLanguageIDArabicLibya</strong>|
|
<strong>msoLanguageIDArabicMorocco</strong>|
|
<strong>msoLanguageIDArabicOman</strong>|
|
<strong>msoLanguageIDArabicQatar</strong>|
|
<strong>msoLanguageIDArabicSyria</strong>|
|
<strong>msoLanguageIDArabicTunisia</strong>|
|
<strong>msoLanguageIDArabicUAE</strong>|
|
<strong>msoLanguageIDArabicYemen</strong>|
|
<strong>msoLanguageIDArmenian</strong>|
|
<strong>msoLanguageIDAssamese</strong>|
|
<strong>msoLanguageIDAzeriCyrillic</strong>|
|
<strong>msoLanguageIDAzeriLatin</strong>|
|
<strong>msoLanguageIDBasque</strong>|
|
<strong>msoLanguageIDBelgianDutch</strong>|
|
<strong>msoLanguageIDBelgianFrench</strong>|
|
<strong>msoLanguageIDBengali</strong>|
|
<strong>msoLanguageIDBrazilianPortuguese</strong>|
|
<strong>msoLanguageIDBulgarian</strong>|
|
<strong>msoLanguageIDBurmese</strong>|
|
<strong>msoLanguageIDByelorussian</strong>|
|
<strong>msoLanguageIDCatalan</strong>|
|
<strong>msoLanguageIDCherokee</strong>|
|
<strong>msoLanguageIDChineseHongKong</strong>|
|
<strong>msoLanguageIDChineseMacao</strong>|
|
<strong>msoLanguageIDChineseSingapore</strong>|
|
<strong>msoLanguageIDCroatian</strong>|
|
<strong>msoLanguageIDCzech</strong>|
|
<strong>msoLanguageIDDanish</strong>|
|
<strong>msoLanguageIDDutch</strong>|
|
<strong>msoLanguageIDEnglishAUS</strong>|
|
<strong>msoLanguageIDEnglishBelize</strong>|
|
<strong>msoLanguageIDEnglishCanadian</strong>|
|
<strong>msoLanguageIDEnglishCaribbean</strong>|
|
<strong>msoLanguageIDEnglishIreland</strong>|
|
<strong>msoLanguageIDEnglishJamaica</strong>|
|
<strong>msoLanguageIDEnglishNewZealand</strong>|
|
<strong>msoLanguageIDEnglishPhilippines</strong>|
|
<strong>msoLanguageIDEnglishSouthAfrica</strong>|
|
<strong>msoLanguageIDEnglishTrinidad</strong>|
|
<strong>msoLanguageIDEnglishUK</strong>|
|
<strong>msoLanguageIDEnglishUS</strong>|
|
<strong>msoLanguageIDEnglishZimbabwe</strong>|
|
<strong>msoLanguageIDEstonian</strong>|
|
<strong>msoLanguageIDFaeroese</strong>|
|
<strong>msoLanguageIDFarsi</strong>|
|
<strong>msoLanguageIDFinnish</strong>|
|
<strong>msoLanguageIDFrench</strong>|
|
<strong>msoLanguageIDFrenchCameroon</strong>|
|
<strong>msoLanguageIDFrenchCanadian</strong>|
|
<strong>msoLanguageIDFrenchCotedIvoire</strong>|
|
<strong>msoLanguageIDFrenchLuxembourg</strong>|
|
<strong>msoLanguageIDFrenchMali</strong>|
|
<strong>msoLanguageIDFrenchMonaco</strong>|
|
<strong>msoLanguageIDFrenchReunion</strong>|
|
<strong>msoLanguageIDFrenchSenegal</strong>|
|
<strong>msoLanguageIDFrenchWestIndies</strong>|
|
<strong>msoLanguageIDFrenchZaire</strong>|
|
<strong>msoLanguageIDFrisianNetherlands</strong>|
|
<strong>msoLanguageIDGaelicIreland</strong>|
|
<strong>msoLanguageIDGaelicScotland</strong>|
|
<strong>msoLanguageIDGalician</strong>|
|
<strong>msoLanguageIDGeorgian</strong>|
|
<strong>msoLanguageIDGerman</strong>|
|
<strong>msoLanguageIDGermanAustria</strong>|
|
<strong>msoLanguageIDGermanLiechtenstein</strong>|
|
<strong>msoLanguageIDGermanLuxembourg</strong>|
|
<strong>msoLanguageIDGreek</strong>|
|
<strong>msoLanguageIDGujarati</strong>|
|
<strong>msoLanguageIDHebrew</strong>|
|
<strong>msoLanguageIDHindi</strong>|
|
<strong>msoLanguageIDHungarian</strong>|
|
<strong>msoLanguageIDIcelandic</strong>|
|
<strong>msoLanguageIDIndonesian</strong>|
|
<strong>msoLanguageIDInuktitut</strong>|
|
<strong>msoLanguageIDItalian</strong>|
|
<strong>msoLanguageIDJapanese</strong>|
|
<strong>msoLanguageIDKannada</strong>|
|
<strong>msoLanguageIDKashmiri</strong>|
|
<strong>msoLanguageIDKazakh</strong>|
|
<strong>msoLanguageIDKhmer</strong>|
|
<strong>msoLanguageIDKirghiz</strong>|
|
<strong>msoLanguageIDKonkani</strong>|
|
<strong>msoLanguageIDKorean</strong>|
|
<strong>msoLanguageIDLao</strong>|
|
<strong>msoLanguageIDLatvian</strong>|
|
<strong>msoLanguageIDLithuanian</strong>|
|
<strong>msoLanguageIDMacedonian</strong>|
|
<strong>msoLanguageIDMalayalam</strong>|
|
<strong>msoLanguageIDMalayBruneiDarussalam</strong>|
|
<strong>msoLanguageIDMalaysian</strong>|
|
<strong>msoLanguageIDMaltese</strong>|
|
<strong>msoLanguageIDManipuri</strong>|
|
<strong>msoLanguageIDMarathi</strong>|
|
<strong>msoLanguageIDMexicanSpanish</strong>|
|
<strong>msoLanguageIDMixed</strong>|
|
<strong>msoLanguageIDMongolian</strong>|
|
<strong>msoLanguageIDNepali</strong>|
|
<strong>msoLanguageIDNone</strong>|
|
<strong>msoLanguageIDNoProofing</strong>|
|
<strong>msoLanguageIDNorwegianBokmol</strong>|
|
<strong>msoLanguageIDNorwegianNynorsk</strong>|
|
<strong>msoLanguageIDOriya</strong>|
|
<strong>msoLanguageIDPolish</strong>|
|
<strong>msoLanguageIDPunjabi</strong>|
|
<strong>msoLanguageIDRhaetoRomanic</strong>|
|
<strong>msoLanguageIDRomanian</strong>|
|
<strong>msoLanguageIDRomanianMoldova</strong>|
|
<strong>msoLanguageIDRussian</strong>|
|
<strong>msoLanguageIDRussianMoldova</strong>|
|
<strong>msoLanguageIDSamiLappish</strong>|
|
<strong>msoLanguageIDSanskrit</strong>|
|
<strong>msoLanguageIDSerbianCyrillic</strong>|
|
<strong>msoLanguageIDSerbianLatin</strong>|
|
<strong>msoLanguageIDSesotho</strong>|
|
<strong>msoLanguageIDSimplifiedChinese</strong>|
|
<strong>msoLanguageIDSindhi</strong>|
|
<strong>msoLanguageIDSlovak</strong>|
|
<strong>msoLanguageIDSlovenian</strong>|
|
<strong>msoLanguageIDSorbian</strong>|
|
<strong>msoLanguageIDSpanish</strong>|
|
<strong>msoLanguageIDSpanishArgentina</strong>|
|
<strong>msoLanguageIDSpanishBolivia</strong>|
|
<strong>msoLanguageIDSpanishChile</strong>|
|
<strong>msoLanguageIDSpanishColombia</strong>|
|
<strong>msoLanguageIDSpanishCostaRica</strong>|
|
<strong>msoLanguageIDSpanishDominicanRepublic</strong>|
|
<strong>msoLanguageIDSpanishEcuador</strong>|
|
<strong>msoLanguageIDSpanishElSalvador</strong>|
|
<strong>msoLanguageIDSpanishGuatemala</strong>|
|
<strong>msoLanguageIDSpanishHonduras</strong>|
|
<strong>msoLanguageIDSpanishModernSort</strong>|
|
<strong>msoLanguageIDSpanishNicaragua</strong>|
|
<strong>msoLanguageIDSpanishPanama</strong>|
|
<strong>msoLanguageIDSpanishParaguay</strong>|
|
<strong>msoLanguageIDSpanishPeru</strong>|
|
<strong>msoLanguageIDSpanishPuertoRico</strong>|
|
<strong>msoLanguageIDSpanishUruguay</strong>|
|
<strong>msoLanguageIDSpanishVenezuela</strong>|
|
<strong>msoLanguageIDSutu</strong>|
|
<strong>msoLanguageIDSwahili</strong>|
|
<strong>msoLanguageIDSwedish</strong>|
|
<strong>msoLanguageIDSwedishFinland</strong>|
|
<strong>msoLanguageIDSwissFrench</strong>|
|
<strong>msoLanguageIDSwissGerman</strong>|
|
<strong>msoLanguageIDSwissItalian</strong>|
|
<strong>msoLanguageIDTajik</strong>|
|
<strong>msoLanguageIDTamil</strong>|
|
<strong>msoLanguageIDTatar</strong>|
|
<strong>msoLanguageIDTelugu</strong>|
|
<strong>msoLanguageIDThai</strong>|
|
<strong>msoLanguageIDTibetan</strong>|
|
<strong>msoLanguageIDTraditionalChinese</strong>|
|
<strong>msoLanguageIDTsonga</strong>|
|
<strong>msoLanguageIDTswana</strong>|
|
<strong>msoLanguageIDTurkish</strong>|
|
<strong>msoLanguageIDTurkmen</strong>|
|
<strong>msoLanguageIDUkrainian</strong>|
|
<strong>msoLanguageIDUrdu</strong>|
|
<strong>msoLanguageIDUzbekCyrillic</strong>|
|
<strong>msoLanguageIDUzbekLatin</strong>|
|
<strong>msoLanguageIDVenda</strong>|
|
<strong>msoLanguageIDVietnamese</strong>|
|
<strong>msoLanguageIDWelsh</strong>|
|
<strong>msoLanguageIDXhosa</strong>|
|
<strong>msoLanguageIDZulu</strong>|
|
<strong>msoLanguageIDPortuguese</strong>|

## Example

This example sets the language for the specified text in shape one to US English.


```vb
ActivePresentation.Slides(1).Shapes(1).TextFrame.TextRange.LanguageID = msoLanguageIDEnglishUS
```


## See also


#### Concepts


[TextRange Object](textrange-object-powerpoint.md)

