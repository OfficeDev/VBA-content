---
title: "Свойство Application.Selection (издатель)"
keywords: vbapb10.chm131109
f1_keywords: vbapb10.chm131109
ms.prod: publisher
api_name: Publisher.Application.Selection
ms.assetid: b4a542a7-cb54-476b-9ccf-004ce4b9ec47
ms.date: 06/08/2017
ms.openlocfilehash: e571e7c7e876e9c33ce37f44f530606c2c713e0a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationselection-property-publisher"></a>Свойство Application.Selection (издатель)

Возвращает объект **[Selection](selection-object-publisher.md)** , который представляет выбранный диапазон или курсор.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Выбор**

 переменная _expression_A, представляющий объект **приложения** .


## <a name="example"></a>Пример

В этом примере проверяется, является ли текущий выделенный фрагмент текста. Если это текст, выбранный текст отображается в окне сообщения.


```vb
Sub Selectable() 
 
 If Selection.Type = pbSelectionText Then MsgBox Selection.TextRange 
 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

