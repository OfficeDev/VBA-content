---
title: "Перечисление PbFixedFormatIntent (издатель)"
keywords: vbapb10.chm65637
f1_keywords: vbapb10.chm65637
ms.prod: publisher
api_name: Publisher.PbFixedFormatIntent
ms.assetid: bddb023b-181f-7805-434f-128f27d609e4
ms.date: 06/08/2017
ms.openlocfilehash: 2384923b6db4967130fbc3505d38656b60b49112
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pbfixedformatintent-enumeration-publisher"></a>Перечисление PbFixedFormatIntent (издатель)

Константы, переданной в метод **[ExportAsFixedFormat](document-exportasfixedformat-method-publisher.md)** , укажите, как пользователь намеревается совместно использовать созданный файл.



|**Имя**|**Значение**|**Описание**|
|:-----|:-----|:-----|
| **pbIntentCommercial**|4|Отправка публикации для профессиональной печати.|
| **pbIntentMinimum**|1|Нажмите публикации до наименьшего размера файла. Это должна удовлетворять на экране просмотров сценарий, где отображаются публикации на экране.|
| **pbIntentPrinting**|3|Печать публикации на настольном принтере или в копии хранилища, например Kinko.|
| **pbIntentStandard**|2|Распространение публикации по электронной почте или с веб-сайта. Обратите внимание на то, что пользователь не знает, как просматривать публикации: на экране или с помощью настольного принтера при выводе на печать. Рабочего стола сценарий печати и на экране Просмотр сценария должны быть выполнены с этой целью.|

