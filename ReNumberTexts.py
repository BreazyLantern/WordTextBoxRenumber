#!/usr/bin/python3
import File_Accessor as F_access

fm = F_access.FileManip()

word_files = None
try:
    word_files = F_access.sys.argv[1:]
    if len(word_files) != 0:
        fm.add_to_list(word_files)
    print()
except Exception as e:
    print(f"Something went wrong with retrieving the file: {e}")

try:
    fm.output_paths()

except Exception as e:
    print(e)

try:
    fm.work_on_all_files()
    #fm.testing()
    #fm.edit_text_in_documents("25_07 Dec, 2024.docx")
except Exception as e:
    print(e)

input("Type to quit")
