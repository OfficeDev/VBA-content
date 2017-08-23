---
title: "Метод WebPageOptions.SetBackgroundSoundRepeat (издатель)"
keywords: vbapb10.chm544777
f1_keywords: vbapb10.chm544777
ms.prod: publisher
api_name: Publisher.WebPageOptions.SetBackgroundSoundRepeat
ms.assetid: a699fa92-a36a-6722-431d-a0ce8413cfcf
ms.date: 06/08/2017
ms.openlocfilehash: 2aff8b18d5ded2b24a941e2f1093f970cb743a44
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webpageoptionssetbackgroundsoundrepeat-method-publisher"></a>Метод WebPageOptions.SetBackgroundSoundRepeat (издатель)

Указывает, необходимо будет воспроизводиться фоновый звук, подключенного к веб-страницы, бесконечно после того, как страница загружается в веб-браузере и противном случае, при необходимости указывает, сколько раз необходимо будет воспроизводиться фоновый звук.


## <a name="syntax"></a>Синтаксис

 _выражение_. **SetBackgroundSoundRepeat** ( **_RepeatForever_**, **_RepeatTimes_**)

 переменная _expression_A, представляет собой объект- **WebPageOptions** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|RepeatForever|Обязательное свойство.| **Boolean**|Указывает, следует ли фоновый звук будет воспроизводиться бесконечно. Значение этого параметра используется для заполнения значение ** [Свойство BackgroundSoundLoopForever](webpageoptions-backgroundsoundloopforever-property-publisher.md)** свойство.|
|RepeatTimes|Необязательный| **Длинный**|Указывает, сколько раз необходимо будет воспроизводиться фоновый звук. Значение этого параметра используется для заполнения значение ** [Свойство BackgroundSoundLoopCount](webpageoptions-backgroundsoundloopcount-property-publisher.md)** свойство.|

## <a name="remarks"></a>Заметки

Если параметр **_RepeatForever_** имеет значение **True**, нельзя указывать необязательный параметр **_RepeatTimes_** . Попытка задать **_RepeatTimes_** , если **_RepeatForever_** имеет **значение True,** приводит к ошибке времени выполнения.

Если параметр **_RepeatForever_** имеет значение **False**, должен быть указан необязательный параметр **_RepeatTimes_** . Пропуск **_RepeatTimes_** , если **_RepeatForever_** имеет **значение False,** приводит к ошибке времени выполнения.


## <a name="example"></a>Пример

В следующем примере задается фон звука для страницы четыре active веб-публикации для WAV-файл на локальном компьютере. Если **BackgroundSoundLoopForever** имеет **значение False**, метод **SetBackgroundSoundRepeat** используется для указания, что звуковое сопровождение повторяться бесконечно (Обратите внимание, заменяют параметр **_RepeatTimes_** ). Если **BackgroundSoundLoopForever** имеет **значение True**, метод **SetBackgroundSoundRepeat** используется для указания, что звуковое сопровождение не повторяться бесконечно, однако, его необходимо выполнять два раза.


```vb
Dim theWPO As WebPageOptions 
 
Set theWPO = ActiveDocument.Pages(4).WebPageOptions 
 
With theWPO 
 .BackgroundSound = "C:\CompanySounds\corporate_jingle.wav" 
 If .BackgroundSoundLoopForever = False Then 
 .SetBackgroundSoundRepeat RepeatForever:=True 
 ElseIf .BackgroundSoundLoopForever = True Then 
 .SetBackgroundSoundRepeat RepeatForever:=False, RepeatTimes:=2 
 End If 
 
End With
```


