Программа для конвертации Excel в различные форматы
Описание программы
Эта программа на языке Python предназначена для конвертации файлов Excel в различные форматы, такие как CSV, JSON, Markdown, YAML и другие. Программа также поддерживает применение различных фильтров для обработки данных перед конвертацией.
Основные возможности:
Конвертация Excel в CSV:
Простая и быстрая конвертация файлов Excel в формат CSV.
Конвертация Excel в JSON:
Поддержка различных форматов JSON, включая списки, объекты и массивы.
Конвертация Excel в Markdown:
Преобразование данных Excel в формат Markdown для удобного чтения и публикации.
Конвертация Excel в YAML:
Конвертация данных в формат YAML для структурированного хранения.
Применение фильтров:
Возможность применения различных фильтров для обработки данных перед конвертацией.

Инструкция по установке и сборке программы
Для установки программы выполните следующие шаги:
Установите зависимости:
В вашем Python окружении выполните команду:
pip install -r requirements.txt
Это установит все необходимые зависимости, перечисленные в файле requirements.txt.
Соберите исполняемый файл:
После установки зависимостей, соберите ваш основной скрипт в исполняемый файл с помощью PyInstaller. Выполните следующую команду:
pyinstaller your_script.py
Эта команда создаст каталог с необходимыми файлами для запуска исполняемого файла.
Запуск исполняемого файла:
После успешной сборки вы можете найти исполняемый файл в созданном каталоге, например, dist/your_script.

Конвертация иконок для macOS
mkdir MyIcon.iconset

iconutil -c icns MyIcon.iconset -o icon.icns

sips -z 16 16 icon.png --out MyIcon.iconset/icon_16x16.png
sips -z 32 32 icon.png --out MyIcon.iconset/icon_32x32.png
sips -z 128 128 icon.png --out MyIcon.iconset/icon_128x128.png
sips -z 256 256 icon.png --out MyIcon.iconset/icon_256x256.png
sips -z 512 512 icon.png --out MyIcon.iconset/icon_512x512.png

Сборка программы для macOS происзводится командой
pyinstaller --windowed --icon=icon.icns --add-data "icon.icns:." --name "Excel-data-tool" main.py 
Команда для сбора в dmg архиы
hdiutil create -volname "Excel-data-tool" -srcfolder Excel-data-tool.app -ov -format UDZO Excel-data-tool.dmg

Сборка приложения под Windows
pyinstaller --onefile --windowed --icon=icon.ico --add-data "icon.ico;." monitoring_app.py

Лицензия
Программа распространяется под лицензией GNU General Public License (GPL). Это означает, что вы можете свободно использовать, изменять и распространять программу, при условии, что любые производные работы также будут распространяться под лицензией GPL.
/*
 * Этот программный продукт является свободным программным обеспечением,
 * распространяемым в соответствии с условиями Стандартной общественной
 * лицензии GNU, версии 3 или любой более поздней версии.
 */

Copyright (C) 2024 Vladimir Babushkin
Для вопросов, предложений и сотрудничества обращайтесь через GitHub Issues.
