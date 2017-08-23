---
title: "Метод PageBackground.Create (издатель)"
keywords: vbapb10.chm8126469
f1_keywords: vbapb10.chm8126469
ms.prod: publisher
api_name: Publisher.PageBackground.Create
ms.assetid: a9b699c4-067a-2c68-5f9b-ee7ba0c22cbd
ms.date: 06/08/2017
ms.openlocfilehash: cd1613e1da52a52e44e75bface4e4215ab7bcf6d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagebackgroundcreate-method-publisher"></a>Метод PageBackground.Create (издатель)

Создает новый объект **PageBackground** для указанного объекта **страницы** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **Создание**

 переменная _expression_A, представляет собой объект- **PageBackground** .


## <a name="remarks"></a>Заметки

Использование PageBackground.Exists для проверки, если страница уже фон перед попыткой создать новую. Возвращает «отказано в разрешении "ошибка возникает, если фона уже существует. 


## <a name="example"></a>Пример

Следующий пример проверяет наличие фона на первой странице активных документов. Если не существует фона затем он будет создан. 


```vb
If ActiveDocument.Pages(1).Background.Exists = False Then 
 ActiveDocument.Pages(1).Background.Create 
End If
```


