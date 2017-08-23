---
title: "Свойство Font.AllCaps (издатель)"
keywords: vbapb10.chm5373959
f1_keywords: vbapb10.chm5373959
ms.prod: publisher
api_name: Publisher.Font.AllCaps
ms.assetid: e8394f91-de31-0075-51ac-8a372023f0ce
ms.date: 06/08/2017
ms.openlocfilehash: 805e2fb05c3e991eb8c158994b277bf4d1ee05c7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontallcaps-property-publisher"></a>Свойство Font.AllCaps (издатель)

Возвращает или задает **msoTrue** , если шрифт имеет формат прописных букв или возвращает один или другие константы **MsoTriState** , если он не установлен. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **AllCaps**

 переменная _expression_A, представляющий объект **Font** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Установка для свойства **AllCaps** **msoTrue** задает свойство **SmallCaps** **msoFalse**и наоборот.

Значение свойства **AllCaps** может иметь одно из ** [MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.


## <a name="example"></a>Пример

В этом примере проверяется выбранного текста в активном документе для текста в формате прописных букв. Для работы этого примера необходимо быть активной публикации с выделенного текста.


```vb
Public Sub Caps() 
 
 If Publisher.ActiveDocument.Selection _ 
 .TextRange.Font.AllCaps = msoTrue Then 
 MsgBox "Text is all caps." 
 Else 
 MsgBox "Text is not all caps." 
 End If 
 
End Sub
```

В этом примере Форматирует выбранный текст в виде прописных букв. Для правильного выполнения этого кода активный документ должен существовать с выделенного текста.




```vb
Public Sub MakeCaps() 
 
 If Publisher.ActiveDocument.Selection.TextRange _ 
 .Font.AllCaps = msoFalse Then 
 Selection.TextRange.Font.AllCaps = msoTrue 
 Else 
 MsgBox "You need to select some text or it is already all caps." 
 End If 
 
End Sub
```


