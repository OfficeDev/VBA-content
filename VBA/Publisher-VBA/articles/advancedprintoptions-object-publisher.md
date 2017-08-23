---
title: "Объект AdvancedPrintOptions (издатель)"
keywords: vbapb10.chm7143423
f1_keywords: vbapb10.chm7143423
ms.prod: publisher
api_name: Publisher.AdvancedPrintOptions
ms.assetid: 61f776cc-dc3e-61b6-057a-125ad15146c8
ms.date: 06/08/2017
ms.openlocfilehash: 142a970642eeb359f29aac6a399d31ecc7d178af
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="advancedprintoptions-object-publisher"></a>Объект AdvancedPrintOptions (издатель)

Представляет Дополнительные параметры печати для публикации.
 


## <a name="remarks"></a>Заметки

Свойства объекта **AdvancedPrintOptions** соответствующие параметры, доступные на вкладках диалогового окна **Дополнительные настройки печати** .
 

 

## <a name="example"></a>Пример

Используйте свойство **AdvancedPrintOptions** объекта **Document** для возврата объекта **AdvancedPrintOptions** . Следующий пример проверяет, чтобы определить, установлено ли active публикации для печати цветоделение. Если Да, оно установлено для печати форм только для красок, используемые в публикации, а также не печатать формы для всех страниц, где не используется цвет.
 

 

```
Sub PrintOnlyInksUsed 
 With ActiveDocument.AdvancedPrintOptions 
 If .PrintMode = pbPrintModeSeparations Then 
 .InksToPrint = pbInksToPrintUsed 
 .PrintBlankPlates = False 
 End If 
 End With 
End Sub
```


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[AllowBleeds](advancedprintoptions-allowbleeds-property-publisher.md)|
|[Приложения](advancedprintoptions-application-property-publisher.md)|
|[BackSideInsertFaceUp](advancedprintoptions-backsideinsertfaceup-property-publisher.md)|
|[GraphicsResolution](advancedprintoptions-graphicsresolution-property-publisher.md)|
|[HorizontalFlip](advancedprintoptions-horizontalflip-property-publisher.md)|
|[IsPostscriptPrinter](advancedprintoptions-ispostscriptprinter-property-publisher.md)|
|[ManualFeedAlign](advancedprintoptions-manualfeedalign-property-publisher.md)|
|[ManualFeedDirection](advancedprintoptions-manualfeeddirection-property-publisher.md)|
|[NegativeImage](advancedprintoptions-negativeimage-property-publisher.md)|
|[PageRotated](advancedprintoptions-pagerotated-property-publisher.md)|
|[Родительский раздел](advancedprintoptions-parent-property-publisher.md)|
|[PrintBleedMarks](advancedprintoptions-printbleedmarks-property-publisher.md)|
|[PrintCropMarks](advancedprintoptions-printcropmarks-property-publisher.md)|
|[PrintDensityBars](advancedprintoptions-printdensitybars-property-publisher.md)|
|[PrintJobInformation](advancedprintoptions-printjobinformation-property-publisher.md)|
|[PrintRegistrationMarks](advancedprintoptions-printregistrationmarks-property-publisher.md)|
|[Решение](advancedprintoptions-resolution-property-publisher.md)|
|[UseOnlyPublicationFonts](advancedprintoptions-useonlypublicationfonts-property-publisher.md)|
|[VerticalFlip](advancedprintoptions-verticalflip-property-publisher.md)|

