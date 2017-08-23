---
title: "Свойство WebPageOptions.BackgroundSoundLoopForever (издатель)"
keywords: vbapb10.chm544775
f1_keywords: vbapb10.chm544775
ms.prod: publisher
api_name: Publisher.WebPageOptions.BackgroundSoundLoopForever
ms.assetid: f2e90665-09e9-5215-59e4-f93e4469d0df
ms.date: 06/08/2017
ms.openlocfilehash: 7d88eac8fa0c5b14b811917f52ba0c2bf1b1b36d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webpageoptionsbackgroundsoundloopforever-property-publisher"></a>Свойство WebPageOptions.BackgroundSoundLoopForever (издатель)

Возвращает **логическое** значение, указывающее, является ли звуковое сопровождение, подключенного к веб-странице должен повторяться бесконечно. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **BackgroundSoundLoopForever**

 переменная _expression_A, представляет собой объект- **WebPageOptions** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Метод **[SetBackgroundSoundRepeat](webpageoptions-setbackgroundsoundrepeat-method-publisher.md)** используется для указания, необходимо ли выполнять звуковое сопровождение бесконечно после загрузки страницы. Пока метод **SetBackgroundSoundRepeat** используется для указания, следует ли фоновый звук будет воспроизводиться бесконечно, **BackgroundSoundLoopForever** имеет **значение False**.


## <a name="example"></a>Пример

В следующем примере задается фон звука для страницы четыре active веб-публикации для WAV-файл на локальном компьютере. Если **BackgroundSoundLoopForever** имеет **значение False**, метод **SetBackgroundSoundRepeat** используется для указания, что звуковое сопровождение должен повторяться бесконечно. Свойство **BackgroundSoundLoopForever** теперь будет иметь **значение True**.


```vb
Dim theWPO As WebPageOptions 
 
Set theWPO = ActiveDocument.Pages(4).WebPageOptions 
 
With theWPO 
 .BackgroundSound = "C:\CompanySounds\corporate_jingle.wav" 
 If .BackgroundSoundLoopForever = False Then 
 .SetBackgroundSoundRepeat RepeatForever:=True 
 End If 
End With
```


