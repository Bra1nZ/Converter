import sys
import os
from docx2pdf import convert

OUTPUT_DIR = r"C:\MT"


def main():
    if len(sys.argv) < 2:
        print("ОШИБКА: Ты не перетащил файлы!")
        print("Выдели нужные .docx файлы и перетащи их прямо на иконку скрипта.")
        input("\nНажмите Enter для выхода...")
        return

    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        print(f"Создана новая папка для сохранения: {OUTPUT_DIR}")

    print("Начинаю конвертацию...\n")

    for file_path in sys.argv[1:]:
        if not file_path.lower().endswith(".docx"):
            print(f"[-] Пропускаю: {os.path.basename(file_path)} (это не .docx)")
            continue

        file_name = os.path.basename(file_path)
        pdf_name = file_name[:-5] + ".pdf"  
        output_path = os.path.join(OUTPUT_DIR, pdf_name)

        print(f"[+] Конвертирую: {file_name} ...")
        
        try:
            convert(file_path, output_path)
        except Exception as e:
            print(f"Ошибка при конвертации {file_name}: {e}")

    print(f"\nГОТОВО! Все PDF сохранены в папку:\n{OUTPUT_DIR}")
    input("\nНажмите Enter для выхода...")

if __name__ == "__main__":
    main()