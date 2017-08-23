---
title: "Свойство WebPageOptions.BackgroundSoundLoopCount (издатель)"
keywords: vbapb10.chm544776
f1_keywords: vbapb10.chm544776
ms.prod: publisher
api_name: Publisher.WebPageOptions.BackgroundSoundLoopCount
ms.assetid: 34d34a04-5fdb-3d43-9140-fcf10b420efd
ms.date: 06/08/2017
ms.openlocfilehash: 6b7aa9b6aa37f566e2c5d346672b50cbf11bcd70
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webpageoptionsbackgroundsoundloopcount-property-publisher"></a>Свойство WebPageOptions.BackgroundSoundLoopCount (издатель)

Возвращает значение типа **Long** , указывает, сколько раз звук фона, подключенного к веб-страницы будет воспроизводиться при загрузке страницы в веб-браузере. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **BackgroundSoundLoopCount**

 переменная _expression_A, представляет собой объект- **WebPageOptions** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="remarks"></a>Заметки

Метод **[SetBackgroundSoundRepeat](webpageoptions-setbackgroundsoundrepeat-method-publisher.md)** можно использовать для указания количество отправок воспроизведения звукового файла фона при загрузке страницы. Если с помощью метода **SetBackgroundSoundRepeat** , чтобы указать, сколько раз воспроизвести файл фонового, свойство **BackgroundSoundLoopCount** будет равен, заданному значению. Обратите внимание, что допустимые значения в диапазоне от 1 до 999 включительно. При попытке установить это значение за пределами этого диапазона приведет к ошибке времени выполнения.

Пока метод **SetBackgroundSoundRepeat** используется для изменения номера отправок воспроизведения звукового файла фона, **BackgroundSoundLoopCount** равен 1.


## <a name="example"></a>Пример

В следующем примере задается фон звука для страницы четыре active веб-публикации для WAV-файл на локальном компьютере. Если **BackgroundSoundLoopCount** меньше, чем три, метод **SetBackgroundSoundRepeat** используется для указания, что звуковое сопровождение повторяться три раза. Свойство **BackgroundSoundLoopCount** будут три.


```vb
Dim theWPO As WebPageOptions 
 
Set theWPO = ActiveDocument.Pages(4).WebPageOptions 
 
With theWPO 
 .BackgroundSound = "C:\CompanySounds\corporate_jingle.wav" 
 If .BackgroundSoundLoopCount < 3 Then 
 .SetBackgroundSoundRepeat RepeatForever:=False, RepeatTimes:=3 
 End If 
End With
```


