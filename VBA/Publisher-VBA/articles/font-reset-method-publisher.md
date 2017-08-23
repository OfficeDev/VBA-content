---
title: "Метод Font.Reset (издатель)"
keywords: vbapb10.chm5373993
f1_keywords: vbapb10.chm5373993
ms.prod: publisher
api_name: Publisher.Font.Reset
ms.assetid: 7a81d7f9-4db9-3ce1-188d-2b4719b57fff
ms.date: 06/08/2017
ms.openlocfilehash: 9e8544adac1b312f86ab81df35d474bdec711926
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontreset-method-publisher"></a>Метод Font.Reset (издатель)

Удаляет вручную абзаца или текст из указанного объекта и оставляет только форматирование, указанного идентификатором текущего стиля текста.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Сброс**

 переменная _expression_A, представляющий объект **Font** .


### <a name="return-value"></a>Возвращаемое значение

Значение Nothing


## <a name="example"></a>Пример

В следующем примере сбрасывается символ форматирования текста в форму одно на странице один из активных публикации для форматирования для текущего стиля текста знаков по умолчанию.


```vb
ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Font.Reset
```

В следующем примере сбрасывается форматирование абзаца текст в фигуре одно на странице один из активных публикации с форматированием для текущего стиля текста абзаца по умолчанию.




```vb
ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.ParagraphFormat.Reset
```


