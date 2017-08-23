---
title: "Свойство TextFrame.Story (издатель)"
keywords: vbapb10.chm3866663
f1_keywords: vbapb10.chm3866663
ms.prod: publisher
api_name: Publisher.TextFrame.Story
ms.assetid: 7bbe0967-83aa-745b-ad13-8a7dfe61811c
ms.date: 06/08/2017
ms.openlocfilehash: 52c2eaae2f54185d9f119eb08f5e800e14810a42
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textframestory-property-publisher"></a>Свойство TextFrame.Story (издатель)

Возвращает объект **материал** , который представляет свойства материала в диапазон текста.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Статья**

 переменная _expression_A, представляет собой объект- **TextFrame** .


## <a name="example"></a>Пример

В этом примере возвращает История в диапазоне выделенный текст и, если фрагмент текста Вставка текста в диапазон текста.


```vb
Sub AddTextToStory() 
 With Selection.TextRange.Story 
 If .HasTextFrame Then 
 .TextRange.InsertAfter NewText:=vbLf &; "This is a test." 
 End If 
 End With 
End Sub
```


