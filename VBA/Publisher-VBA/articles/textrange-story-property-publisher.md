---
title: "Свойство TextRange.Story (издатель)"
keywords: vbapb10.chm5308470
f1_keywords: vbapb10.chm5308470
ms.prod: publisher
api_name: Publisher.TextRange.Story
ms.assetid: 833f9537-5c11-a4d5-907a-777eaecb89d2
ms.date: 06/08/2017
ms.openlocfilehash: 32753305a83ba5302f037d6b55a0e1d00385b37f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangestory-property-publisher"></a>Свойство TextRange.Story (издатель)

Возвращает объект **материал** , который представляет свойства материала в диапазон текста.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Статья**

 переменная _expression_A, представляющий объект **TextRange** .


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


