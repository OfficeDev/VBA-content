---
title: "Метод LinkFormat.Update (издатель)"
keywords: vbapb10.chm4390916
f1_keywords: vbapb10.chm4390916
ms.prod: publisher
api_name: Publisher.LinkFormat.Update
ms.assetid: a167a463-56bd-2c4e-ded5-70ea38b2ed2f
ms.date: 06/08/2017
ms.openlocfilehash: 88fd4b0e26576a31c01f8da404eb484610edace4
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="linkformatupdate-method-publisher"></a>Метод LinkFormat.Update (издатель)

Обновляет указанный связанный объект OLE.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Обновление**

 переменная _expression_A, представляет собой объект- **LinkFormat** .


## <a name="example"></a>Пример

В этом примере обновляются все связанные объекты OLE в активной публикации.


```vb
Dim pageLoop As Page 
Dim shpLoop As Shape 
 
For Each pageLoop In ActiveDocument.Pages 
 For Each shpLoop In pageLoop.Shapes 
 
 With shpLoop 
 If .Type = pbLinkedOLEObject Then 
 .LinkFormat.Update 
 End If 
 End With 
 
 Next shpLoop 
Next pageLoop
```


