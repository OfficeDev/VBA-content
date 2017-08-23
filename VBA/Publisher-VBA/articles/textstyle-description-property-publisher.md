---
title: "Свойство TextStyle.Description (издатель)"
keywords: vbapb10.chm5963779
f1_keywords: vbapb10.chm5963779
ms.prod: publisher
api_name: Publisher.TextStyle.Description
ms.assetid: 278d647e-c4bc-218c-417d-b01caf2d98a3
ms.date: 06/08/2017
ms.openlocfilehash: 3137aab4268dd33d0a57b287ea05b77393bf2c23
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textstyledescription-property-publisher"></a>Свойство TextStyle.Description (издатель)

Возвращает **строку** , представляющую описание указанного стиля. К примеру, могут быть типичное описание обычный стиль «Roman нового времени (по умолчанию), Mincho мс (восточно-азиатский), 10 пунктов, Main (черный) кернинг 14 pt слева, строки разделитель 1 sp.» Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Описание**

 переменная _expression_A, представляющий объект **стиля текста** .


## <a name="example"></a>Пример

В этом примере отображается описание для стиля Обычный.


```vb
Sub ShowStyleDescription() 
 MsgBox "The Normal style has the following formatting attributes: " &; _ 
 vbLf &; ActiveDocument.TextStyles("Normal").Description 
End Sub
```


