---
title: "Метод ParagraphFormat.Reset (издатель)"
keywords: vbapb10.chm5439509
f1_keywords: vbapb10.chm5439509
ms.prod: publisher
api_name: Publisher.ParagraphFormat.Reset
ms.assetid: 8ef5c799-cace-133c-33d3-3454df2c2f24
ms.date: 06/08/2017
ms.openlocfilehash: e3c6cdae4889f075c0311298142e8ec6086abfa4
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatreset-method-publisher"></a>Метод ParagraphFormat.Reset (издатель)

Удаляет вручную абзаца или текст из указанного объекта и оставляет только форматирование, указанного идентификатором текущего стиля текста.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Сброс**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


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


