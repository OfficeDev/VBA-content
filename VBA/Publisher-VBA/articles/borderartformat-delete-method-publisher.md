---
title: "Метод BorderArtFormat.Delete (издатель)"
keywords: vbapb10.chm7602184
f1_keywords: vbapb10.chm7602184
ms.prod: publisher
api_name: Publisher.BorderArtFormat.Delete
ms.assetid: 3ec0576f-8304-2647-7309-b014b586c1b6
ms.date: 06/08/2017
ms.openlocfilehash: 184e243f872b4315f4b06444fcae7b67dde3d2c9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="borderartformatdelete-method-publisher"></a>Метод BorderArtFormat.Delete (издатель)

Удаляет указанный объект.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Удаление**

 переменная _expression_A, представляет собой объект- **BorderArtFormat** .


## <a name="remarks"></a>Заметки

Если указанный объект не существует, возникает ошибка времени выполнения.


## <a name="example"></a>Пример

Следующий пример проверяет наличие Узорные на каждой фигуры для каждой страницы публикации active. Если Узорные существует, она удаляется.


```vb
Sub DeleteBorderArt() 
 
Dim anyPage As Page 
Dim anyShape As Shape 
 
For Each anyPage in ActiveDocument.Pages 
 For Each anyShape in anyPage.Shapes 
 With anyShape.BorderArt 
 If .Exists = True Then 
 .Delete 
 End If 
 End With 
 Next anyShape 
 Next anyPage 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект BorderArtFormat](borderartformat-object-publisher.md)

