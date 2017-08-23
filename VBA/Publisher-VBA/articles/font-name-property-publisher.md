---
title: "Свойство Font.Name (издатель)"
keywords: vbapb10.chm5373952
f1_keywords: vbapb10.chm5373952
ms.prod: publisher
api_name: Publisher.Font.Name
ms.assetid: 03561991-5456-aee3-4c04-56a2520a4d6e
ms.date: 06/08/2017
ms.openlocfilehash: ca9aaf6760112b792937400419e14abd8fe32a18
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fontname-property-publisher"></a>Свойство Font.Name (издатель)

Указывает имя выбранного шрифта. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Имя**

 _expression_An выражение, возвращающее объект **Font** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

В этом примере форматов фрагмент текста на одну страницу вместе с Arial полужирным шрифтом.


```vb
With ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Font 
 .Name = "Arial" 
 .Bold = True 
End With 

```


