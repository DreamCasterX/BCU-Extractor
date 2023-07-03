import prettytable as pt
import os
from prettytable.colortable import ColorTable, Themes


tb1 = pt.PrettyTable()
tb1.align = "l"   # 靠左對齊


# 選擇風格
tb1.set_style(pt.SINGLE_BORDER)
# tb1.set_style(pt.DOUBLE_BORDER)
# tb1.set_style(pt.PLAIN_COLUMNS)
# tb1.set_style(pt.MSWORD_FRIENDLY)

tb1.field_names = ["SN#", "System BIOS Version", "Marketing Name", "Feature Byte"]
folder_path = "./BCU_Files"
for filename in os.listdir(folder_path):
    if filename.endswith(".txt"):
        with open(os.path.join(folder_path, filename), "r") as file:
            lines = file.readlines()
            for i in range(len(lines)):
                if "System BIOS Version" in lines[i]:
                    BIOS = str(lines[i+1].strip())
                    BIOS_Ver = BIOS[0:17]
                if "Primary Battery Serial Number" not in lines[i]:
                    if "Serial Number" in lines[i]:
                        SN = str(lines[i+1].strip())
                if "Product Name" in lines[i]:
                    PD = str(lines[i+1].strip())
                if "Feature Byte" in lines[i]:
                    FB = str(lines[i+1].strip())
                    tb1.add_row([SN, BIOS_Ver, PD, FB])
print(tb1)

def Save_to_CSV():
    with open('Summary.csv', 'w', newline='', encoding="utf-8") as w:
        w.write(str(tb1.get_csv_string()))
def Save_to_TXT():
    with open('Summary.txt', 'w', encoding="utf-8") as f:
        f.write(str(tb1))


if len(os.listdir(folder_path)) != 0:
    Save_to_TXT()
input()





