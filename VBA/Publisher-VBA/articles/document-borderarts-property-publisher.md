---
title: "Свойство Document.BorderArts (издатель)"
keywords: vbapb10.chm196721
f1_keywords: vbapb10.chm196721
ms.prod: publisher
api_name: Publisher.Document.BorderArts
ms.assetid: 5639ffce-f711-71b6-78f8-2de63fe50a3c
ms.date: 06/08/2017
ms.openlocfilehash: 137071883d12c4d19226fb34cb200d48932d59b6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentborderarts-property-publisher"></a>Свойство Document.BorderArts (издатель)

Возвращает коллекцию **[BorderArts](borderarts-object-publisher.md)** , представляющий Узорные типы, доступные для использования в указанной публикации. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **BorderArts**

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

BorderArts


## <a name="remarks"></a>Заметки

Узорные, границы изображения, которые можно применять для текстовых полей, рамки рисунков или прямоугольники. 

Коллекция **BorderArts** включает все пользовательские типы Узорные, создаваемые пользователем для указанной публикации.


## <a name="example"></a>Пример

В следующем примере возвращается коллекция BorderArts и перечислены имена всех типов Узорные, доступных для использования в активной публикации.


```vb
Sub ListBorderArt() 
Dim bdaTemp As BorderArts 
Dim bdaLoop As BorderArt 
 
Set bdaTemp = ActiveDocument.BorderArts 
 
For Each bdaLoop In bdaTemp 
 Debug.Print "The name of this BorderArt is " &; bdaLoop.Name 
Next bdaLoop 
End Sub
```


