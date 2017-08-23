---
title: "Свойство Document.Find (издатель)"
keywords: vbapb10.chm196725
f1_keywords: vbapb10.chm196725
ms.prod: publisher
api_name: Publisher.Document.Find
ms.assetid: e9b31937-4504-79b5-5913-b2ef0a23f2a7
ms.date: 06/08/2017
ms.openlocfilehash: f7138c7a5f32f58edcdc93259e1501689675de1b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentfind-property-publisher"></a>Свойство Document.Find (издатель)

## <a name="syntax"></a>Синтаксис

 _выражение_. **Поиск**

 переменная _expression_A, представляющий объект **Document** .


## <a name="example"></a>Пример

Применительно к объекта **Document** .

В следующем примере задается объектную переменную объекту **FindReplace** для активных документов. Выполняет операцию поиска, который применяет жирное форматирование для каждого вхождения слово «важно».




```vb
Dim objFind as FindReplace 
Dim fFound as Boolean 
 
Set objFind = ActiveDocument.Find 
fFound = True 
 
With objFind 
 .Clear 
 .FindText = "important" 
 Do While fFound = True 
 fFound = .Execute 
 If Not .FoundTextRange Is Nothing Then 
 .FoundTextRange.Font.Bold = True 
 End If 
 Loop 
End With 
```

Применительно к объекту **TextRange** .

В следующем примере задается объектную переменную объекту **FindReplace** текстового диапазона первой фигуры в активный документ. Выполняет операцию поиска, которое применяется для каждого вхождения слово «срочно» в диапазоне текст полужирным шрифтом.




```vb
Dim objFind as FindReplace 
Dim fFound as Boolean 
 
Set objFind = ActiveDocument.Pages(1) _ 
 .Shapes(1).TextFrame.TextRange.Find 
fFound = True 
 
With objFind 
 .Clear 
 .FindText = "urgent" 
 Do While fFound = True 
 fFound = .Execute 
 If Not .FoundTextRange Is Nothing Then 
 .FoundTextRange.Font.Bold = True 
 End If 
 Loop 
End With
```


