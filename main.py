"""
Editing with Helix Text Editor and Ruff LSP for python 
"""
import os
import re
from openpyxl import Workbook


def tokenize_inl_query(inl_q: str) -> list[str]:
    temp_list: list[str] = inl_q.split()
    tbl_list: list[str] = [item for item in temp_list if (item.startswith("tbl"))]
    return tbl_list


def main() -> None:
    folder: str = input("Enter folder path:\t")
    filelist: list[str] = []
    excel_fn: str = "test.xlsx"

    sp_methods: list[str] = [
        "ExecuteNonQuerySP",
        "ExecuteNonQueryAsyncSP",
        "ExecuteReaderSP",
        "ExecuteReaderAsyncSP",
        "ExecuteScalarSP",
        "ExecuteScalarAsyncSP",
        "ExecuteDataSetSP",
    ]

    tbl_methods: list[str] = [
        "FillDropDownOnly",
    ]

    wb = Workbook()
    ws = wb.active
    # Headers
    ws["A1"] = "File Path"
    ws["B1"] = "SP Count"
    ws['C1'] = "SP Line No."
    ws["D1"] = "SP List"
    ws["E1"] = "Table Count"
    ws["F1"] = "Table List"
    ws["G1"] = "Query Line No."
    ws["H1"] = "Table Query"
    # r=root, d=directories, f = files
    for r, d, f in os.walk(folder):
        for file in f:
            if ".cs" in file:
                filelist.append(os.path.join(r, file))

    for file in filelist:
        excel_row = []
        sp_count: int = 0
        table_count: int = 0
        sp_list: list[str] = []
        sp_ln: list[int] = []
        table_list: list[str] = []
        inl_ln: list[int] = []
        inl_query: list[str] = []
        print(file)
        with open(file, "r") as f:
            lines = f.readlines()
            line_num = 0
            for line in lines:
                line_num += 1
                if not line.startswith("//"):
                    if bool([ele for ele in sp_methods if (ele in line)]):
                        sp_list.extend(re.findall(r'"([^"]*)"', line))
                        sp_ln.append(line_num)
                        sp_count = sp_count + 1
                    elif bool([ele for ele in tbl_methods if (ele in line)]):
                        inl_query.append(re.findall(r'"([^"]*)"', line))
                        inl_ln.append(line_num)
                        match = re.search(r'"([^"]*)"', line)
                        if match:
                            m_tbl = match.group(1)
                            if not bool(re.search(r"\s", m_tbl)):
                                table_list.append(m_tbl)
                                # inl_ln.append(line_num)
                                table_count += 1
                            elif bool(re.search(r"\s", m_tbl)):
                                # tokenize should just return a list of tbls
                                tbl_list = tokenize_inl_query(m_tbl)
                                table_list.extend(tbl_list)
                                # inl_ln.append(line_num)
                                table_count += 1
        print(
            f"Filename:\t{file}\nSP Count:\t{sp_count}\nSP List:\t{sp_list}\n"
            + f"Table Count:\t{table_count}\nTable List:\t{table_list}"
        )
        excel_row.append(file)
        excel_row.append(sp_count)
        excel_row.append("\n".join(map(str, sp_ln)))
        excel_row.append("\n".join(sp_list))
        excel_row.append(table_count)
        excel_row.append("\n".join(table_list))
        excel_row.append("\n".join(map(str, inl_ln)))
        excel_row.append("\n".join(map(str, inl_query)))
        # Add line nums just for queries. not for separate tables
        ws.append(excel_row)

    wb.save(excel_fn)


if __name__ == "__main__":
    main()
