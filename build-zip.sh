#!/usr/bin/env bash
# ============================================================
# build-zip.sh — сборка access-course-project.zip (Linux/macOS)
# Использование:
#   cd /path/to/access-course
#   chmod +x build-zip.sh && ./build-zip.sh
# Зависимость: zip (apt-get install zip | brew install zip)
# ============================================================
set -euo pipefail

# --- 1. Каталог проекта (где лежит скрипт) ---
PROJECT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$PROJECT_DIR"
echo "Каталог проекта: $PROJECT_DIR"

# --- 2. Имя архива, удаление старого ---
ZIP_NAME="access-course-project.zip"
ZIP_PATH="$PROJECT_DIR/$ZIP_NAME"
if [ -f "$ZIP_PATH" ]; then
  rm -f "$ZIP_PATH"
  echo "Старый архив удалён: $ZIP_NAME"
fi

# --- 3. Проверка утилиты zip ---
if ! command -v zip >/dev/null 2>&1; then
  echo "ОШИБКА: утилита zip не найдена." >&2
  echo "  Debian/Ubuntu : sudo apt-get install zip" >&2
  echo "  macOS         : brew install zip (или используйте встроенный ditto)" >&2
  exit 1
fi

# --- 4. Исключения: git, временные, системные, кэш, данные ---
EXCLUDES=(
  ".git/*"            ".git"
  "node_modules/*"    "node_modules"
  ".vscode/*"         ".vscode"
  ".idea/*"           ".idea"
  ".cache/*"          ".cache"
  "import/*"          "import"
  "dist/*"            "build/*"
  "*.tmp" "*.temp" "*.swp" "*.swo" "*.bak" "*~" "*.log"
  ".DS_Store" "Thumbs.db" "Desktop.ini"
  "*.zip" "*.accdb" "*.mdb" "*.laccdb" "*.ldb" "*.csv" "*.xlsx"
)

# --- 5. Создать ZIP (все файлы проекта, с папкой access-course внутри) ---
zip -r -q "$ZIP_PATH" . -x "${EXCLUDES[@]}"

# --- 5b. CSV репозитория — часть учебника, возвращаются в архив ---
# datasets/calendar.csv (761 дата) и starter-data/*.csv (стартовый комплект, 6 файлов)
# являются частью курса; пользовательские *.csv (выгрузки Olist, import/) исключены выше.
REPO_CSVS=( "datasets/calendar.csv" "starter-data/categories.csv" "starter-data/products.csv" "starter-data/customers.csv" "starter-data/employees.csv" "starter-data/orders.csv" "starter-data/order-details.csv" )
for f in "${REPO_CSVS[@]}"; do
  if [ -f "$f" ]; then
    zip -q "$ZIP_PATH" "$f"
    echo "Добавлен $f"
  else
    echo "ПРЕДУПРЕЖДЕНИЕ: $f не найден — архив собран без него." >&2
  fi
done

# --- 6. Проверить наличие архива ---
if [ ! -f "$ZIP_PATH" ]; then
  echo "ОШИБКА: архив не создан." >&2
  exit 1
fi

# --- 7. Вывести список файлов в архиве ---
echo ""
echo "=== Содержимое архива ==="
unzip -l "$ZIP_PATH" | awk 'NR>3 {print} /files$/ {print}'
COUNT=$(unzip -l "$ZIP_PATH" | awk '/files$/{print $2}')
[ -z "$COUNT" ] && COUNT=$(unzip -l "$ZIP_PATH" | grep -c 'access-course/')

# --- 8. Размер архива ---
SIZE_BYTES=$(stat -c%s "$ZIP_PATH" 2>/dev/null || stat -f%z "$ZIP_PATH")
SIZE_KB=$((SIZE_BYTES / 1024))
if [ "$SIZE_BYTES" -ge 1048576 ]; then
  # сотые доли МБ без bc (чистый bash): 4500000 -> 4.29
  SIZE_CENTI=$((SIZE_BYTES * 100 / 1048576))
  echo "Размер архива   : $((SIZE_CENTI / 100)).$((SIZE_CENTI % 100)) МБ (${SIZE_BYTES} байт)"
else
  echo "Размер архива   : ${SIZE_KB} КБ (${SIZE_BYTES} байт)"
fi

# --- 9. Тест целостности и успешное завершение ---
if unzip -t "$ZIP_PATH" >/dev/null 2>&1; then
  echo ""
  echo "ГОТОВО: $ZIP_NAME создан успешно (файлов: $COUNT)."
  echo "Проверка: unzip -t access-course-project.zip && unzip access-course-project.zip -d ../check"
  exit 0
else
  echo "ОШИБКА: архив повреждён." >&2
  exit 1
fi
