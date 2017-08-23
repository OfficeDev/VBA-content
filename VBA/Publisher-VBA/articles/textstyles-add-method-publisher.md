---
title: "Метод TextStyles.Add (издатель)"
keywords: vbapb10.chm5898244
f1_keywords: vbapb10.chm5898244
ms.prod: publisher
api_name: Publisher.TextStyles.Add
ms.assetid: 56bb84a2-5632-1baa-4b97-3c48d43367bf
ms.date: 06/08/2017
ms.openlocfilehash: 98e70bcfea9c115d20b9813b85f512d0245abf65
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textstylesadd-method-publisher"></a>Метод TextStyles.Add (издатель)

Добавляет новый объект **стиля текста** на указанный объект **TextStyles** и возвращает новый объект **стиля текста** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Добавление** ( **_Шрифт_**, **_ParagraphFormat_** **_StyleName_**, **_BasedOn_**)

 переменная _expression_A, представляет собой объект- **TextStyles** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|StyleName|Обязательное свойство.| **String**|Имя нового стиля текста. Если имя соответствует существующего стиля текста, будет перезаписан существующий стиль текста.|
|BasedOn|Необязательный| **String**|Имя стиля текста, лежащие в основе нового стиля текста. Если имя не соответствует существующего стиля текста, возникает ошибка.|
|Font|Необязательный| **Шрифт**|Параметры шрифта для применения нового стиля текста.|
|ParagraphFormat|Необязательный| **ParagraphFormat**|Чтобы применить новый стиль текста форматирование абзаца.|

### <a name="return-value"></a>Возвращаемое значение

Стиля текста


## <a name="example"></a>Пример

В следующем примере добавляется новый стиль текста для активной публикации на основании стиля Обычный текст.


```vb
Dim tsNew As TextStyle 
 
Set tsNew = ActiveDocument.TextStyles _ 
 .Add(StyleName:="Title", BasedOn:="Normal")
```


