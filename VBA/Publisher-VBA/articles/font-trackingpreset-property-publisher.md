---
title: "Свойство Font.TrackingPreset (издатель)"
keywords: vbapb10.chm5373986
f1_keywords: vbapb10.chm5373986
ms.prod: publisher
api_name: Publisher.Font.TrackingPreset
ms.assetid: 818e6efd-a1b3-1ccd-1dc1-29c0a8ded7f2
ms.date: 06/08/2017
ms.openlocfilehash: c1c1d12997e7405228324b6331ce3ccf11f0912f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fonttrackingpreset-property-publisher"></a>Свойство Font.TrackingPreset (издатель)

Возвращает или задает значение константы **PbTrackingPresetType** , представляющее тип предварительно отслеживания для символов шрифта, указанного в диапазон текста. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **TrackingPreset**

 переменная _expression_A, представляющий объект **Font** .


### <a name="return-value"></a>Возвращаемое значение

PbTrackingPresetType


## <a name="remarks"></a>Заметки

Значение свойства **TrackingPreset** может иметь одно из следующих констант **PbTrackingPresetType** .



| **pbTrackingCustom**|| **pbTrackingLoose**|| **pbTrackingMixed**|| **pbTrackingNormal**|| **pbTrackingTight**|| **pbTrackingVeryLoose**|| **pbTrackingVeryTight**| Отслеживание свободном и очень широкий покидает достаточно между символами, в то время как строгий контроль и очень узкий отслеживания может осуществлять перекрытие символ.


## <a name="example"></a>Пример

В этом примере указывается строгий контроль отслеживания виде стиля символов во второй материал.


```vb
Sub TrackingType() 
 
 Application.ActiveDocument.Stories(2).TextRange _ 
 .Font.TrackingPreset = pbTrackingTight 
 
End Sub
```


