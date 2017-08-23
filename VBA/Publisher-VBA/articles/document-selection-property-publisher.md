---
title: "Свойство Document.Selection (издатель)"
keywords: vbapb10.chm196658
f1_keywords: vbapb10.chm196658
ms.prod: publisher
api_name: Publisher.Document.Selection
ms.assetid: b1098cdb-8fb7-0906-b193-6dc572ac2993
ms.date: 06/08/2017
ms.openlocfilehash: 2747400785cfd806562acb12ef967c91ebbe24d2
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentselection-property-publisher"></a>Свойство Document.Selection (издатель)

Возвращает объект **[Selection](selection-object-publisher.md)** , который представляет выбранный диапазон или курсор.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Выбор**

 переменная _expression_A, представляющий объект **Document** .


## <a name="example"></a>Пример

В этом примере проверяется, является ли текущий выделенный фрагмент текста. Если это текст, выбранный текст отображается в окне сообщения.


```vb
Sub Selectable() 
 
 If Selection.Type = pbSelectionText Then MsgBox Selection.TextRange 
 
End Sub
```


