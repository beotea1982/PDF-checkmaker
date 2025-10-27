#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
PDF Generator Script
Генерирует PDF-документы на основе CSV данных и HTML-шаблонов
"""

import os
import sys
import pandas as pd
import platform
from pathlib import Path
from string import Template

# Настройка кодировки для Windows
if platform.system() == 'Windows':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8')
    
    # Подавляем предупреждения GTK+ и GLib
    os.environ['G_MESSAGES_DEBUG'] = '0'
    import warnings
    warnings.filterwarnings('ignore', category=UserWarning)
    warnings.filterwarnings('ignore', category=RuntimeWarning)

# Попытаемся использовать WeasyPrint, если не работает - используем альтернативу
USE_WEASYPRINT = False
USE_XHTML2PDF = False
pisa = None

try:
    from weasyprint import HTML, CSS
    from weasyprint.text.fonts import FontConfiguration
    USE_WEASYPRINT = True
except (ImportError, OSError) as e:
    USE_WEASYPRINT = False
    print(f"⚠️  WeasyPrint недоступен: {e}")
    print("📌 Ищу альтернативную библиотеку...")
    try:
        from xhtml2pdf import pisa
        print("✅ Использую xhtml2pdf")
        USE_XHTML2PDF = True
    except ImportError:
        print("❌ Не найдена альтернативная библиотека")
        USE_XHTML2PDF = False
        print("\n💡 Установите одну из библиотек:")
        print("   pip install xhtml2pdf")
        print("   или установите GTK+ для WeasyPrint")
        sys.exit(1)

# Директории для работы
DATA_DIR = Path('data')
TEMPLATES_DIR = Path('templates')
OUTPUT_DIR = Path('output')


def clear_screen():
    """Очищает экран консоли"""
    if platform.system() == 'Windows':
        os.system('cls')
    else:
        os.system('clear')


def print_menu_header(title):
    """Выводит заголовок меню"""
    print("\n" + "=" * 60)
    print(f"  {title}")
    print("=" * 60)


def print_menu(items):
    """Выводит нумерованный список"""
    for i, item in enumerate(items, 1):
        print(f"  {i}. {item}")
    print()


def get_choice(prompt, max_value):
    """Получает выбор пользователя с валидацией"""
    while True:
        try:
            choice = input(prompt)
            choice_num = int(choice)
            if 1 <= choice_num <= max_value:
                return choice_num
            else:
                print(f"❌ Пожалуйста, введите число от 1 до {max_value}")
        except ValueError:
            print("❌ Пожалуйста, введите корректное число")
        except KeyboardInterrupt:
            print("\n\n👋 Программа завершена пользователем")
            sys.exit(0)


def get_csv_files():
    """Получает список CSV файлов из директории data"""
    csv_files = list(DATA_DIR.glob('*.csv'))
    if not csv_files:
        print("❌ Ошибка: Не найдено CSV файлов в директории 'data'")
        sys.exit(1)
    return csv_files


def get_html_templates():
    """Получает список HTML-шаблонов из директории templates"""
    html_files = list(TEMPLATES_DIR.glob('*.html'))
    if not html_files:
        print("❌ Ошибка: Не найдено HTML-шаблонов в директории 'templates'")
        sys.exit(1)
    return html_files


def load_csv_data(csv_path):
    """Загружает данные из CSV файла"""
    try:
        df = pd.read_csv(csv_path, encoding='utf-8')
        return df
    except Exception as e:
        print(f"❌ Ошибка при чтении файла {csv_path}: {e}")
        return None


def read_html_template(template_path):
    """Читает HTML-шаблон"""
    try:
        with open(template_path, 'r', encoding='utf-8') as f:
            return f.read()
    except Exception as e:
        print(f"❌ Ошибка при чтении шаблона {template_path}: {e}")
        return None


def render_act_template(template_html, positions_df):
    """Подставляет данные в шаблон акта"""
    if len(positions_df) == 0:
        print("❌ Нет данных для генерации акта")
        return None
    
    # Берем данные из первой строки для заголовка
    first_row = positions_df.iloc[0]
    
    # Генерируем строки таблицы
    positions_rows = ""
    total_amount = 0
    
    for idx, row in enumerate(positions_df.itertuples(), 1):
        try:
            total = float(row.total)
            total_amount += total
        except:
            total = 0
        
        positions_rows += f"""
      <tr>
        <td>{idx}</td>
        <td>{row.work_name}</td>
        <td>{row.quantity}</td>
        <td>{row.unit}</td>
        <td>{row.unit_price}</td>
        <td>{row.total}</td>
      </tr>"""
    
    # Заполняем шаблон
    template = Template(template_html)
    html = template.safe_substitute(
        act_number=first_row.act_number,
        act_date=first_row.act_date,
        contract_number=first_row.contract_number,
        object_address=first_row.object_address,
        positions_rows=positions_rows,
        total_amount=f"{total_amount:.2f}",
        client=first_row.client,
        contractor=first_row.contractor
    )
    
    return html


def render_certificate_template(template_html, data_row):
    """Подставляет данные в шаблон сертификата"""
    template = Template(template_html)
    html = template.safe_substitute(
        student_name=data_row.student_name,
        course_name=data_row.course_name,
        cert_date=data_row.cert_date
    )
    return html


def render_template(template_path, template_html, data_df, selected_id):
    """Подставляет данные в шаблон в зависимости от типа"""
    # Фильтруем данные по ID
    filtered_df = data_df[data_df['id'] == selected_id]
    
    if len(filtered_df) == 0:
        print(f"❌ Не найдено данных для ID {selected_id}")
        return None
    
    template_name = template_path.stem
    
    if template_name == 'act':
        # Для актов: группируем все строки с одинаковым ID
        return render_act_template(template_html, filtered_df)
    elif template_name == 'certificate':
        # Для сертификатов: берем первую строку
        return render_certificate_template(template_html, filtered_df.iloc[0])
    else:
        print(f"❌ Неизвестный тип шаблона: {template_name}")
        return None


def generate_pdf(html_content, output_path):
    """Генерирует PDF из HTML"""
    try:
        if USE_WEASYPRINT:
            # Используем WeasyPrint
            font_config = FontConfiguration()
            HTML(string=html_content).write_pdf(
                output_path,
                font_config=font_config
            )
        elif USE_XHTML2PDF:
            # Используем xhtml2pdf
            with open(output_path, 'wb') as result_file:
                pdf = pisa.CreatePDF(
                    html_content,
                    dest=result_file,
                    encoding='utf-8'
                )
            if pdf.err:
                raise Exception(f"Ошибка xhtml2pdf: {pdf.err}")
        else:
            raise Exception("Не найдена библиотека для генерации PDF")
        
        print(f"✅ PDF успешно создан: {output_path}")
        return True
    except Exception as e:
        print(f"❌ Ошибка при генерации PDF: {e}")
        return False


def open_pdf(pdf_path):
    """Открывает PDF в системной программе"""
    try:
        if platform.system() == 'Windows':
            os.startfile(pdf_path)
        elif platform.system() == 'Darwin':  # macOS
            os.system(f'open "{pdf_path}"')
        else:  # Linux
            os.system(f'xdg-open "{pdf_path}"')
    except Exception as e:
        print(f"⚠️  Не удалось открыть PDF автоматически: {e}")


def main():
    """Главная функция"""
    clear_screen()
    
    # Проверяем наличие необходимых директорий
    for directory in [DATA_DIR, TEMPLATES_DIR, OUTPUT_DIR]:
        if not directory.exists():
            directory.mkdir(exist_ok=True)
            print(f"📁 Создана директория: {directory}")
    
    print("\n🔧 PDF Генератор - Запуск\n")
    
    # ШАГ 1: Выбираем шаблон
    html_templates = get_html_templates()
    print_menu_header("📄 Доступные HTML-шаблоны")
    template_filenames = [f.name for f in html_templates]
    print_menu(template_filenames)
    
    template_choice = get_choice(f"Выберите шаблон (1-{len(html_templates)}): ", len(html_templates))
    selected_template = html_templates[template_choice - 1]
    
    # Читаем шаблон
    print(f"\n📝 Загрузка шаблона {selected_template.name}...")
    template_html = read_html_template(selected_template)
    if template_html is None:
        return
    
    # ШАГ 2: Выбираем CSV файл
    csv_files = get_csv_files()
    print_menu_header("📊 Доступные файлы с данными")
    csv_filenames = [f.name for f in csv_files]
    print_menu(csv_filenames)
    
    csv_choice = get_choice(f"Выберите файл с данными (1-{len(csv_files)}): ", len(csv_files))
    selected_csv = csv_files[csv_choice - 1]
    
    # Загружаем данные
    print(f"\n📖 Загрузка данных из {selected_csv.name}...")
    data_df = load_csv_data(selected_csv)
    if data_df is None:
        return
    
    # ШАГ 3: Выбираем режим генерации
    print_menu_header("🔢 Режим генерации")
    print_menu([
        "Создать один PDF для конкретного ID",
        "Создать PDF для всех ID в файле"
    ])
    
    mode_choice = get_choice("Выберите режим (1-2): ", 2)
    
    # Получаем уникальные ID
    unique_ids = sorted(data_df['id'].unique())
    
    if mode_choice == 1:
        # Режим 1: Один PDF
        print_menu_header(f"🔢 Доступные ID в файле {selected_csv.name}")
        print_menu([f"ID: {id_val} ({len(data_df[data_df['id'] == id_val])} записей)" for id_val in unique_ids])
        
        id_choice = get_choice(f"Выберите ID (1-{len(unique_ids)}): ", len(unique_ids))
        selected_id = unique_ids[id_choice - 1]
        
        # Генерируем один PDF
        print(f"\n🔨 Генерация PDF для ID {selected_id}...")
        rendered_html = render_template(selected_template, template_html, data_df, selected_id)
        if rendered_html is None:
            return
        
        output_filename = f"{selected_template.stem}_{selected_csv.stem}_id{selected_id}.pdf"
        output_path = OUTPUT_DIR / output_filename
        
        if generate_pdf(rendered_html, output_path):
            print("\n🚀 Открытие PDF файла...")
            open_pdf(output_path)
            print(f"\n✅ Готово! PDF файл создан и открыт: {output_filename}")
        else:
            print("\n❌ Произошла ошибка при генерации PDF.")
    
    else:
        # Режим 2: Несколько PDF
        print(f"\n🔨 Генерация PDF для всех ID...")
        print(f"Будет создано {len(unique_ids)} файлов\n")
        
        created_files = []
        errors = []
        
        for idx, selected_id in enumerate(unique_ids, 1):
            print(f"  [{idx}/{len(unique_ids)}] Генерация ID {selected_id}...", end=" ")
            
            rendered_html = render_template(selected_template, template_html, data_df, selected_id)
            if rendered_html is None:
                errors.append(f"ID {selected_id}")
                print("❌")
                continue
            
            output_filename = f"{selected_template.stem}_{selected_csv.stem}_id{selected_id}.pdf"
            output_path = OUTPUT_DIR / output_filename
            
            if generate_pdf(rendered_html, output_path):
                created_files.append(output_filename)
                print("✅")
            else:
                errors.append(f"ID {selected_id}")
                print("❌")
        
        print(f"\n✅ Создано PDF файлов: {len(created_files)}")
        if created_files:
            print("\n📁 Созданные файлы:")
            for fname in created_files:
                print(f"   • {fname}")
        
        if errors:
            print(f"\n⚠️  Ошибки при генерации {len(errors)} файлов: {', '.join(errors)}")
        
        print(f"\n📂 Все файлы сохранены в директории: output/")


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n👋 Программа завершена пользователем")
        sys.exit(0)
    except Exception as e:
        print(f"\n❌ Произошла ошибка: {e}")
        sys.exit(1)

