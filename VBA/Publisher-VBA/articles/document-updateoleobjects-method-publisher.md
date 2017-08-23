---
title: "Метод Document.UpdateOLEObjects (издатель)"
keywords: vbapb10.chm196706
f1_keywords: vbapb10.chm196706
ms.prod: publisher
api_name: Publisher.Document.UpdateOLEObjects
ms.assetid: 2c07e755-6f5c-5fd8-091c-fbe3bfae6692
ms.date: 06/08/2017
ms.openlocfilehash: 458279fc1ab5a4f92dc469b7dc17cda07062ea25
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentupdateoleobjects-method-publisher"></a>Метод Document.UpdateOLEObjects (издатель)

Обновления связанных и внедренных объектов OLE.


## <a name="syntax"></a>Синтаксис

 _выражение_. **UpdateOLEObjects**

 переменная _expression_A, представляющий объект **Document** .


## <a name="example"></a>Пример

В этом примере обновляются все объекты OLE в активной публикации.


```vb
Sub SearchAndUpdateOLEObjects() 
 ActiveDocument.UpdateOLEObjects 
End Sub
```


