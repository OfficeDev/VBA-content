---
title: "Объект ScratchArea (издатель)"
keywords: vbapb10.chm1245183
f1_keywords: vbapb10.chm1245183
ms.prod: publisher
api_name: Publisher.ScratchArea
ms.assetid: 41856866-c1d8-2550-1b4c-28886ed2b714
ms.date: 06/08/2017
ms.openlocfilehash: e46ff691f269ed997d42137a6d47089b00101e33
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="scratcharea-object-publisher"></a>Объект ScratchArea (издатель)

Представляет область за пределами границ страницы публикации, хранения элементов layout не оказывает влияния на выходные данные публикации.
 


## <a name="example"></a>Пример

Используйте свойство **[ScratchArea](document-scratcharea-property-publisher.md)** объекта **Document** для возврата вспомогательной области. Используйте свойство **фигур** объекта **ScratchArea** для возврата коллекции фигур, которые в настоящее время на вспомогательной области.
 

 

 

 
В этом примере присваивает переменной первую фигуру на вспомогательной области активных документов.
 

 



```
Dim saPage As ScratchArea 
Dim objFirst As Object 
 
saPage = Application.ActiveDocument.ScratchArea 
objFirst = saPage.Shapes(1)
```


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](scratcharea-application-property-publisher.md)|
|[Родительский раздел](scratcharea-parent-property-publisher.md)|
|[Фигур](scratcharea-shapes-property-publisher.md)|

