---
title: "Свойство WebOptions.ShowOnlyWebFonts (издатель)"
keywords: vbapb10.chm8257544
f1_keywords: vbapb10.chm8257544
ms.prod: publisher
api_name: Publisher.WebOptions.ShowOnlyWebFonts
ms.assetid: d18197f4-9abe-d523-77fd-f33a8ecc8076
ms.date: 06/08/2017
ms.openlocfilehash: 46fc9cc110f35ba71fc24a2f60125d947b67cdfd
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="weboptionsshowonlywebfonts-property-publisher"></a>Свойство WebOptions.ShowOnlyWebFonts (издатель)

Возвращает или задает **логическое** значение, указывающее, следует ли использовать только безопасные шрифты и схемы шрифтов при просмотре веб-сайт в браузере. Если **значение True**, только безопасные шрифты и схемы шрифтов используются. Если **значение False**, отображаемое не ограничивается безопасные шрифты и схемы шрифтов. Значение по умолчанию — **False**. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ShowOnlyWebFonts**

 переменная _expression_A, представляет собой объект- **WebOptions** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="remarks"></a>Заметки

Это свойство применяется к на основе латиница только шрифты.


## <a name="example"></a>Пример

Следующий пример указывает, что только безопасные шрифты и схемы шрифтов должен использоваться при просмотре веб-сайт в браузере.


```vb
Dim theWO As WebOptions 
 
Set theWO = Application.WebOptions 
 
With theWO 
 .ShowOnlyWebFonts = True 
End With
```


