RS Matrix NDW / Nasdaq Style

Что внутри:
- index.html — основной сайт
- sample-data.xlsx — пример Excel файла

Как загрузить на GitHub / Vercel:
1. Распакуйте архив
2. Загрузите в репозиторий только эти файлы
3. На Vercel выберите Other / Static
4. Никакие npm install не нужны
5. Build command не нужен
6. Output directory не нужен

Локально:
- просто откройте index.html
или
- запустите любой простой static server

Формат Excel:
- 1-я строка: заголовки
- колонка A: DateTime
- колонки B..: цены активов

Логика:
- weekly Friday closes
- P&F relative strength matrix
- коды BX / BO / SO / SX
- сортировка weak -> strong
