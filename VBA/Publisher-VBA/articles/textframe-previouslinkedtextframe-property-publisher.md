---
title: "Свойство TextFrame.PreviousLinkedTextFrame (издатель)"
keywords: vbapb10.chm3866656
f1_keywords: vbapb10.chm3866656
ms.prod: publisher
api_name: Publisher.TextFrame.PreviousLinkedTextFrame
ms.assetid: 00947ec3-fcff-4451-491b-5b7748ccb74e
ms.date: 06/08/2017
ms.openlocfilehash: c8042d644b3265ca91c05f557a5ce21925ae7226
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textframepreviouslinkedtextframe-property-publisher"></a>Свойство TextFrame.PreviousLinkedTextFrame (издатель)

Возвращает объект **[TextFrame](textframe-object-publisher.md)** , представляющий рамки с какой текст располагается frame указанный текст.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PreviousLinkedTextFrame**

 переменная _expression_A, представляет собой объект- **TextFrame** .


### <a name="return-value"></a>Возвращаемое значение

TextFrame


## <a name="remarks"></a>Заметки

Если указанный текст frame не является частью цепочки связанных рамок или является первым в цепочке связанных рамок, данное свойство возвращает значение nothing.


## <a name="example"></a>Пример

В следующем примере возвращает ранее связанного рамки фигуры три на странице один из активных публикации и задает его шрифт Times New Roman.


```vb
Dim txtFrame As TextFrame 
 
Set txtFrame = ActiveDocument.Pages(1) _ 
 .Shapes(3).TextFrame.PreviousLinkedTextFrame 
 
txtFrame.TextRange.Font = "Times New Roman"
```


