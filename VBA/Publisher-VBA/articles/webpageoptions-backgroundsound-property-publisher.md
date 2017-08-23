---
title: "Свойство WebPageOptions.BackgroundSound (издатель)"
keywords: vbapb10.chm544774
f1_keywords: vbapb10.chm544774
ms.prod: publisher
api_name: Publisher.WebPageOptions.BackgroundSound
ms.assetid: c6be30e0-28ea-e269-c546-48e0eb284ac4
ms.date: 06/08/2017
ms.openlocfilehash: 880e202eb3282f532185b70f1184dbd0431303de
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webpageoptionsbackgroundsound-property-publisher"></a>Свойство WebPageOptions.BackgroundSound (издатель)

Возвращает или задает **строку** , которая указывает путь к звуковой файл, который будет воспроизводиться при загрузке веб-страницы в веб-браузере. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **'' Фоновый звук ''**

 переменная _expression_A, представляет собой объект- **WebPageOptions** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="remarks"></a>Заметки

Путь к фоновый звуковой файл должен быть сетевой или локальный путь; адрес http:// не будут работать.

Если указано **'' фоновый звук ''** фоновый звук будет воспроизводиться один раз по умолчанию. Метод **[SetBackgroundSoundRepeat](webpageoptions-setbackgroundsoundrepeat-method-publisher.md)** можно использовать для указания ли звуковое сопровождение для прослушивания бесконечно и противном случае, чтобы указать, сколько раз звукового файла фона для прослушивания.

Звуковое сопровождение может быть любой из следующих типов файлов:



|*.wav | |*. Mid | | *.midi | |*. RMI | | *.au | |*. AIF | | * .aiff |

## <a name="example"></a>Пример

В следующем примере задается фон звука для страницы четыре active веб-публикации для WAV-файл на локальном компьютере. Этот файл WAV будет воспроизведен один раз при загрузке страницы в веб-браузере.


```vb
Dim theWPO As WebPageOptions 
 
Set theWPO = ActiveDocument.Pages(4).WebPageOptions 
 
With theWPO 
 .BackgroundSound = "C:\CompanySounds\corporate_jingle.wav" 
End With
```


