---
title: "Свойство WebOptions.EmailAsImg (издатель)"
keywords: vbapb10.chm8257545
f1_keywords: vbapb10.chm8257545
ms.prod: publisher
api_name: Publisher.WebOptions.EmailAsImg
ms.assetid: c44d3b07-2030-4901-b9df-4dcfe08c985c
ms.date: 06/08/2017
ms.openlocfilehash: 25b0acb5b1ef844324be277e570ee3d0f1d13f0d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="weboptionsemailasimg-property-publisher"></a>Свойство WebOptions.EmailAsImg (издатель)

 **Значение true,** Чтобы отправить страницу публикации в виде одного изображения в формате JPEG. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **EmailAsImg**

 переменная _expression_A, представляющий объект **WebOptions** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Это свойство можно увеличить совместимость с устаревшими клиентами электронной почты, но может привести к увеличению размера файла.

Это свойство доступно для печати публикаций в дополнение к веб-публикации.

Свойства объекта **[WebOptions](weboptions-object-publisher.md)** используются для указания режима веб-публикации. Это означает, что если какие-либо из этих свойств изменяются, только что созданный веб-публикации будет наследовать измененных свойств.

Это свойство соответствует флажок в разделе **Параметры электронной почты** на вкладке **веб** диалогового окна **Параметры** .


## <a name="example"></a>Пример

В следующем примере задается Microsoft Publisher по электронной почте страниц публикации как изображений в формате JPEG.


```vb
Application.WebOptions.EmailAsImg = True
```


