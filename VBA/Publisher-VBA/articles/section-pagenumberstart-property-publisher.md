---
title: "Свойство Section.PageNumberStart (издатель)"
keywords: vbapb10.chm7405572
f1_keywords: vbapb10.chm7405572
ms.prod: publisher
api_name: Publisher.Section.PageNumberStart
ms.assetid: f4124b7d-4a90-2489-13da-947df0c34a3d
ms.date: 06/08/2017
ms.openlocfilehash: e502ea64b72ba0a354938e920843288814755f4f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="sectionpagenumberstart-property-publisher"></a>Свойство Section.PageNumberStart (издатель)

Задает или возвращает номер страницы, указанного раздела начинается с. Чтение и запись **времени**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PageNumberStart**

 переменная _expression_A, представляет собой объект **раздела** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="example"></a>Пример

В следующем примере задается начальный номер страницы для первого раздела активного документа для 45.


```vb
ActiveDocument.Sections(1).PageNumberStart = 45 

```


