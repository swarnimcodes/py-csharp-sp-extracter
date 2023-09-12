"""
Editing with Helix Text Editor and Ruff LSP for python 
"""
import os
import re
from openpyxl import Workbook


def tokenize(inl_q: str) -> list[str]:
    temp_list = inl_q.split()
    tbl_list = [item for item in temp_list if (item.startswith("tbl"))]
    return tbl_list


def strip_comments(file_content: str) -> str:
    stripped_contents = ""
    return stripped_contents


def main() -> None:
    folder = input("Enter folder path:\t")
    filelist = []
    excel_fn = "test.xlsx"

    sp_func = [
        "ExecuteNonQuerySP",
        "ExecuteNonQueryAsyncSP",
        "ExecuteReaderSP",
        "ExecuteReaderAsyncSP",
        "ExecuteScalarSP",
        "ExecuteScalarAsyncSP",
        "ExecuteDataSetSP",
    ]

    tbl_func = [
        "FillDropDownOnly",
    ]

    wb = Workbook()
    ws = wb.active
    # Headers
    ws["A1"] = "File Path"
    ws["B1"] = "SP Count"
    ws["C1"] = "SP List"
    ws["D1"] = "Table Count"
    ws["E1"] = "Table List"
    # r=root, d=directories, f = files
    for r, d, f in os.walk(folder):
        for file in f:
            if ".cs" in file:
                filelist.append(os.path.join(r, file))

    for file in filelist:
        xl_append = []
        sp_count = 0
        table_count = 0
        sp_list = []
        table_list = []
        print(file)
        with open(file, "r") as f:
            lines = f.readlines()
            for line in lines:
                if not line.startswith("//"):
                    if bool([ele for ele in sp_func if (ele in line)]):
                        sp_list.extend(re.findall(r'"([^"]*)"', line))
                        sp_count = sp_count + 1
                    elif bool([ele for ele in tbl_func if (ele in line)]):
                        match = re.search(r'"([^"]*)"', line)
                        if match:
                            m_tbl = match.group(1)
                            if not bool(re.search(r"\s", m_tbl)):
                                table_list.append(m_tbl)
                                table_count += 1
                            elif bool(re.search(r"\s", m_tbl)):
                                # tokenize should just return a list of tbls
                                tbl_list = tokenize(m_tbl)
                                table_list.extend(tbl_list)
                                table_count += 1
        print(
            f"Filename:\t{file}\nSP Count:\t{sp_count}\nSP List:\t{sp_list}\n"
            + f"Table Count:\t{table_count}\nTable List:\t{table_list}"
        )
        xl_append.append(file)
        xl_append.append(sp_count)
        xl_append.append("\n".join(sp_list))
        xl_append.append(table_count)
        xl_append.append("\n".join(table_list))
        ws.append(xl_append)

    wb.save(excel_fn)


if __name__ == "__main__":
    main()
