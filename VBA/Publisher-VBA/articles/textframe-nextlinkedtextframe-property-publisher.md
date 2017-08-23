---
title: "Свойство TextFrame.NextLinkedTextFrame (издатель)"
keywords: vbapb10.chm3866648
f1_keywords: vbapb10.chm3866648
ms.prod: publisher
api_name: Publisher.TextFrame.NextLinkedTextFrame
ms.assetid: 5ba08ab5-8515-4efe-59a3-79a11f6a7c4e
ms.date: 06/08/2017
ms.openlocfilehash: e664fd02645ab8f675e9c7c51c0fa2f7adfa2e74
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textframenextlinkedtextframe-property-publisher"></a>Свойство TextFrame.NextLinkedTextFrame (издатель)

Возвращает или задает объект **[TextFrame](textframe-object-publisher.md)** , представляющий рамки для каких потоки текста из рамки указанный текст. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **NextLinkedTextFrame**

 переменная _expression_A, представляет собой объект- **TextFrame** .


### <a name="return-value"></a>Возвращаемое значение

TextFrame


## <a name="remarks"></a>Заметки

Если указанный текст frame не является частью цепочки связанных рамок или в конце цепочки связанных рамок, данное свойство возвращает значение nothing.


## <a name="example"></a>Пример

Следующий пример возвращает следующий кадр связанный текст фигуры три на странице один активный публикации и задает его шрифт Times New Roman.


```vb
Dim txtFrame As TextFrame 
 
Set txtFrame = ActiveDocument.Pages(1) _ 
 .Shapes(3).TextFrame.NextLinkedTextFrame 
 
txtFrame.TextRange.Font = "Times New Roman"
```


