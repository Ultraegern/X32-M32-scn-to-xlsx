import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from tkinter import messagebox
import xlsxwriter
import os

inspiration: str = "https://github.com/cabcookie/saddleback-x32-general-scene?tab=readme-ov-file"

def get_file_path() -> str:
    root = Tk()
    Tk().withdraw()  # Hide the root window
    root.wm_attributes("-topmost", True)
    root.withdraw()
    filetypes = [("X32/M32 scene files", "*.scn")]  # Specify allowed file types
    filename = askopenfilename(filetypes=filetypes, parent=root)
    if filename:
        return os.path.abspath(filename)
    else:
        print("Could not get file path")
        return ""

def save_data(data_to_save: pd.DataFrame, all_data: pd.DataFrame, writer: any, start_row: int = 0, start_col: int = 0) -> None:
    data_to_save.to_excel(writer, sheet_name="Kanalplan", index=False, startrow=start_row, startcol=start_col)

    workbook = writer.book
    worksheet = writer.sheets["Kanalplan"]

    # Define formats
    formats: dict = {
        "Black": workbook.add_format({"bg_color": "#9f9f9f", "font_color": "#ffffff", 'border': 1, 'valign': 'vcenter'}),
        "Red": workbook.add_format({"bg_color": "#ff9f9f", "font_color": "#000000", 'border': 1, 'valign': 'vcenter'}),
        "Green": workbook.add_format({"bg_color": "#9fff9f", "font_color": "#000000", 'border': 1, 'valign': 'vcenter'}),
        "Yellow": workbook.add_format({"bg_color": "#ffff9f", "font_color": "#000000", 'border': 1, 'valign': 'vcenter'}),
        "Blue": workbook.add_format({"bg_color": "#879fff", "font_color": "#000000", 'border': 1, 'valign': 'vcenter'}),
        "Magenta": workbook.add_format({"bg_color": "#ff9fff", "font_color": "#000000", 'border': 1, 'valign': 'vcenter'}),
        "Cyan": workbook.add_format({"bg_color": "#9fdaff", "font_color": "#000000", 'border': 1, 'valign': 'vcenter'}),
        "White": workbook.add_format({"bg_color": "#ffffff", "font_color": "#000000", 'border': 1, 'valign': 'vcenter'})
    }

    if "DCA" in data_to_save.columns:
        col_idx = data_to_save.columns.get_loc("DCA")
        start_row_merge = 1
        for row in range(2, len(data_to_save) + 1):
            if data_to_save.iloc[row - 1, col_idx] != data_to_save.iloc[row - 2, col_idx]:
                if start_row_merge != row - 1:
                    worksheet.merge_range(start_row_merge + start_row, col_idx + start_col, row + start_row - 1, col_idx + start_col, data_to_save.iloc[start_row_merge - 1, col_idx])
                start_row_merge = row
        if start_row_merge != len(data_to_save):
            worksheet.merge_range(start_row_merge + start_row, col_idx + start_col, len(data_to_save) + start_row, col_idx + start_col, data_to_save.iloc[start_row_merge - 1, col_idx])
    
    # Colour rows + DCAs
    for row_idx in range(len(data_to_save)):
        color = all_data.at[row_idx, "Colour"]
        if "DCA" in all_data: DCA_color = all_data.at[row_idx, "DCA Colour"] if pd.notna(all_data.at[row_idx, "DCA Colour"]) else "White"
        if color in formats:
            for col_idx, col_name in enumerate(data_to_save.columns):
                if col_name != "DCA":  # Use the general color for other columns
                    worksheet.write(row_idx + 1 + start_row, col_idx + start_col, data_to_save.iloc[row_idx, col_idx], formats[color])
                else:  # Use the DCA color for the "DCA" column
                    if DCA_color in formats:
                        worksheet.write(row_idx + 1 + start_row, col_idx + start_col, data_to_save.iloc[row_idx, col_idx], formats[DCA_color])
                    else:
                        worksheet.write(row_idx + 1 + start_row, col_idx + start_col, data_to_save.iloc[row_idx, col_idx], formats["White"])

    # Auto-adjust column width
    for column in data_to_save.columns:
        column_width = max(all_data[column].astype(str).map(len).max(), len(column)) + 2
        col_idx = data_to_save.columns.get_loc(column) + start_col
        worksheet.set_column(col_idx, col_idx, column_width)

def save_to_excel(input_data: pd.DataFrame = None, output_data: pd.DataFrame = None, aux_input_data: pd.DataFrame = None, output_path: str = "C:/tmp/Kanalplan.xlsx", input_columns_to_save: list = ["Ch", "Pysical Ch", "Name", "DCA"], aux_input_columns_to_save: list = ["Ch", "Pysical Ch", "Name", "DCA"], output_columns_to_save: list = ["Ch", "Mixer Ch", "Name"]) -> None:
    if os.path.exists(output_path):
        if confirm_overwrite(output_path):
            if check_file(output_path):
                pass
            else:
                return
        else:
            return
    
    input_data_to_save: pd.DataFrame
    if input_columns_to_save:
        input_data_to_save = input_data[input_columns_to_save]
    else:
        input_data_to_save = input_data
        
    aux_input_data_to_save: pd.DataFrame
    if aux_input_columns_to_save:
        aux_input_data_to_save = aux_input_data[aux_input_columns_to_save]
    else:
        aux_input_data_to_save = aux_input_data
    
    output_data_to_save: pd.DataFrame
    if output_columns_to_save:
        output_data_to_save = output_data[output_columns_to_save]
    else:
        output_data_to_save = output_data

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer: 
        save_data(input_data_to_save, input_data, writer, start_row= 2, start_col= 1) # Write inputs
        save_data(aux_input_data_to_save, aux_input_data, writer, start_row= 2, start_col= 6) # Write aux inputs
        save_data(output_data_to_save, output_data, writer, start_row= 2, start_col= 11) # Write outputs

        workbook = writer.book
        worksheet = writer.sheets["Kanalplan"]

        # Write headers
        merge_format = workbook.add_format({
            'bold': 1,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#000000',
            "font_color": "#ffffff"
        })
        worksheet.merge_range('B2:E2', 'Inputs', merge_format)
        worksheet.merge_range('G2:J2', 'Aux Inputs', merge_format)
        worksheet.merge_range('L2:N2', 'Outputs', merge_format) 
        
    os.startfile(output_path)

def confirm_overwrite(path: str) -> bool:
    root = Tk()
    root.wm_attributes("-topmost", True)
    root.withdraw()
    result = messagebox.askyesno("Output file already exists", f'Output file already exists.\nAre you sure you want to overwrite "{path}" ?', parent=root)
    # if result:
    #     print("File will be overwritten.")
    #     print("")
    # else:
    #     print("File will not be overwritten.")
    #     print("")
    return result

def check_file(path: str) -> bool:
    root = Tk()
    root.wm_attributes("-topmost", True)
    root.withdraw()
    if is_file_in_use(path):
        retry = messagebox.askretrycancel("File in Use", f'The file "{path}" is currently in use. \nPlease close the file and retry.')
        if retry:
            if check_file(path):
                return True
        else:
            return False
    else:
        return True

def is_file_in_use(path: str) -> bool:
    try:
        # Try to open the file in exclusive mode
        with open(path, 'r+'):
            pass
    except IOError:
        # If an IOError is raised, the file is in use
        return True
    return False

def get_first_DCA_name(lines: list[str], ch: str, aux_ch: bool = False) -> str:
    if not get_grp_line(lines, ch).split(" %")[1].find("1") == -1:
        return get_DCA_names(lines)[DCA_inver_number_lookup_table[get_grp_line(lines, ch, aux_ch).split(" %")[1].find("1")]]
    else:
        return ""

def get_grp_line(lines: list[str], ch: str, aux_ch: bool = False) -> str:
    line_index: int = None
    iteraton: int = -1
    for line in lines:
        iteraton += 1
        if not aux_ch:
            if line.find("/ch/" + ch + "/grp") == 0:
                line_index = iteraton
        else:
            if line.find("/auxin/" + ch + "/grp") == 0:
                line_index = iteraton
    return lines[line_index]

def get_DCA_names(lines: list[str]) -> tuple[str]:
    DCAs: list = []
    for line in lines:
        if line.find("/dca/") == 0 and line.find("/config") == 6:
            DCAs.append(line.split('"')[1])
    return DCAs

def get_DCA_colours(lines: list[str]) -> tuple[str]:
    DCAs: list = []
    for line in lines:
        if line.find("/dca/") == 0 and line.find("/config") == 6:
            DCAs.append(colour_lookup_table[line.split('"')[2].split(" ")[2].split(f"\n")[0]])
    return DCAs

def get_first_DCA_colour(lines: list[str], ch: str, aux_ch: bool = False) -> str:
    if not get_grp_line(lines, ch).split(" %")[1].find("1") == -1:
        return get_DCA_colours(lines)[DCA_inver_number_lookup_table[get_grp_line(lines, ch, aux_ch).split(" %")[1].find("1")]]
    else:
        return ""

def get_user_in_routing_indexes(lines: str) -> list[int]:
    user_in_routing: list[int] = []
    for line in lines:
        if line.find("/config/userrout/in") == 0: #Get user in routing
            for value in line.split("/config/userrout/in")[1].split(" "):
                if not value == "":
                    user_in_routing.append(int(value))
    return user_in_routing

def get_blocks(lines: list[str]) -> list[str]:
    blocks: list[str] = []
    for line in lines:
        if line.find("/config/routing/IN") == 0:
           for item in line.strip(f"\n").split("/config/routing/IN")[1].split(" "):
               if item == "":
                   pass 
               else:
                   blocks.append(item)
    return blocks

def get_block_routing(ch: int, blocks: list[str], user_in_routing_indexes: list[int]) -> str:
    if 1 <= ch <= 8:
        if block_routing_lookup_table[blocks[0]][0] == "":
            return user_in_routing_lookup_tabel[user_in_routing_indexes[ch + block_routing_lookup_table[blocks[0]][1] - 1]]
        else:
            return block_routing_lookup_table[blocks[0]][ch - 1]
    elif 9 <= ch <= 16:
        if block_routing_lookup_table[blocks[1]][0] == "":
            return user_in_routing_lookup_tabel[user_in_routing_indexes[ch + block_routing_lookup_table[blocks[1]][1] - 8 - 1]]
        else:
            return block_routing_lookup_table[blocks[1]][ch - 1 - 8]
    elif 17 <= ch <= 24:
        if block_routing_lookup_table[blocks[2]][0] == "":
            return user_in_routing_lookup_tabel[user_in_routing_indexes[ch + block_routing_lookup_table[blocks[2]][1] - 16 - 1]]
        else:
            return block_routing_lookup_table[blocks[2]][ch - 1 - 16]
    elif 25 <= ch <= 32:
        if block_routing_lookup_table[blocks[3]][0] == "":
            return user_in_routing_lookup_tabel[user_in_routing_indexes[ch + block_routing_lookup_table[blocks[3]][1] - 24 - 1]]
        else:
            return block_routing_lookup_table[blocks[3]][ch - 1 - 24]

def get_override_routing(line: str, blocks: list[str], user_in_routing_indexes: list[int]) -> str:
    override_index: int = int(line.split('"')[2].split(" ")[3])
    if 1 <= override_index <= 32:
        return get_block_routing(override_index, blocks, user_in_routing_indexes) 
    else:
        return override_routing_lookup_table[override_index]

def get_inputs(lines: list[str]) -> pd.DataFrame:
    inputs: pd.DataFrame = pd.DataFrame([], columns = ["In/Out", "Mixer Ch", "Pysical Ch", "Name", "Colour", "Icon", "DCA", "DCA Colour"])
    # user_in_routing: list[int] = get_user_in_routing_indexes(lines)
    for line in lines:
        if line.find("/ch/") == 0 and line.find("/config ") == 6: #Read input ch
            curent_ch: str = line.split("ch/")[1].split("/config")[0]
            new_data: dict = {
                "In/Out": "In",
                "Ch": int(curent_ch),
                "Mixer Ch": "Ch" + curent_ch,
                # "Pysical Ch": routing_lookup_tabel[user_in_routing[int(curent_ch) - 1] - 0],
                "Pysical Ch": get_override_routing(line, get_blocks(lines), get_user_in_routing_indexes(lines)),
                "Name": line.split('"')[1],
                "Colour": colour_lookup_table[line.split('"')[2].split(" ")[2]],
                "Icon": icon_lookup_tabel[int(line.split('"')[2].split(" ")[1]) - 1],
                "DCA": get_first_DCA_name(lines, curent_ch),
                "DCA Colour": get_first_DCA_colour(lines, curent_ch)
            }
            inputs = pd.concat([inputs, pd.DataFrame([new_data])], ignore_index=True)
    return inputs

def find_output_line(lines: str, output_index: int) -> str:
    for line in lines:
        if line.find("/" + output_lookup_table[output_index][0] + "/" + output_lookup_table[output_index][1] + "/config") == 0:
            return line
    return ""

def get_outputs(lines: list[str]) -> pd.DataFrame:
    outputs: pd.DataFrame = pd.DataFrame([], columns = ["In/Out", "Ch", "Mixer Ch", "Name", "Colour"])

    for line in lines:
        if line.find("/outputs/main/") == 0 and line.find(" ") == 16:
            curent_ch: str = line.split("/outputs/main/")[1].split(" ")[0]
            output_index: int = int(line.split("/outputs/main/")[1].split(" ")[1])

            output_line = find_output_line(lines, output_index)
            if output_line == "":
                new_data: dict = {
                    "In/Out": "Out",
                    "Ch": int(curent_ch),
                    "Mixer Ch": output_lookup_table[output_index][2],
                    "Name": "Off",
                    "Colour": "White",
                }
            elif output_lookup_table[output_index][1] == "st" and output_index == 1 or output_index == 2:
                new_data: dict = {
                    "In/Out": "Out",
                    "Ch": int(curent_ch),
                    "Mixer Ch": output_lookup_table[output_index][2],
                    "Name": "LR",
                    "Colour": colour_lookup_table[output_line.split('"')[2].split(" ")[2].strip(f"\n")],
                }
            elif output_lookup_table[output_index][1] == "m":
                new_data: dict = {
                    "In/Out": "Out",
                    "Ch": int(curent_ch),
                    "Mixer Ch": output_lookup_table[output_index][2],
                    "Name": "M/C",
                    "Colour": colour_lookup_table[output_line.split('"')[2].split(" ")[2].strip(f"\n")],
                }
            else:
                new_data: dict = {
                    "In/Out": "Out",
                    "Ch": int(curent_ch),
                    "Mixer Ch": output_lookup_table[output_index][2],
                    "Name": output_line.split('"')[1],
                    "Colour": colour_lookup_table[output_line.split('"')[2].split(" ")[2].strip(f"\n")],
                }
                
            outputs = pd.concat([outputs, pd.DataFrame([new_data])], ignore_index=True)

    return outputs

def get_lines(path: str) -> list[str]:
    with open(path) as file:
        lines: list[str] = file.readlines()
    return lines

def get_aux_ch_routing(lines: list[str]) -> list[str]:
    output: list[str] = []
    aux_remap: str = get_aux_remap(lines)
    user_in_routing_indexes: list[int] = get_user_in_routing_indexes(lines)
    for ch in range(8):
        remap: str = aux_ch_remap[aux_remap][ch]
        if remap == "":
            output.append(user_in_routing_lookup_tabel[user_in_routing_indexes[ch]])
        else:
            output.append(remap)
    return output

def get_aux_remap(lines: list[str]) -> str:
    for line in lines:
        if line.find("/config/routing/IN") == 0:
            return line.split(" ")[5].strip(f"\n")

def get_aux_inputs(lines: list[str]) -> pd.DataFrame:
    inputs: pd.DataFrame = pd.DataFrame([], columns = ["In/Out", "Mixer Ch", "Pysical Ch", "Name", "Colour", "Icon", "DCA", "DCA Colour"])
    for line in lines:
        if line.find("/auxin/") == 0 and line.find("/config ") == 9:
            curent_ch: str = line.split("/auxin/")[1].split("/config")[0]
            new_data: dict = {
                "In/Out": "In",
                "Ch": int(curent_ch),
                "Mixer Ch": "Aux" + curent_ch,
                "Pysical Ch": get_aux_ch_routing(lines)[int(curent_ch) - 1],
                "Name": line.split('"')[1],
                "Colour": colour_lookup_table[line.split('"')[2].split(" ")[2]],
                "Icon": icon_lookup_tabel[int(line.split('"')[2].split(" ")[1]) - 1],
                "DCA": get_first_DCA_name(lines, curent_ch, aux_ch= True),
                "DCA Colour": get_first_DCA_colour(lines, curent_ch, aux_ch= True)
            }
            inputs = pd.concat([inputs, pd.DataFrame([new_data])], ignore_index=True)
    return inputs

aux_ch_remap: dict[tuple[str]] = {
    "AUX1-4": ["Aux in 1", "Aux in 2", "Aux in 3", "Aux in 4", "Aux in 5", "Aux in 6", "USB L", "USB R"],
    "AN1-2": ["Local in 1", "Local in 2", "Aux in 3", "Aux in 4", "Aux in 5", "Aux in 6", "USB L", "USB R"],
    "AN1-4": ["Local in 1", "Local in 2", "Local in 3", "Local in 4", "Aux in 5", "Aux in 6", "USB L", "USB R"],
    "AN1-6": ["Local in 1", "Local in 2", "Local in 3", "Local in 4", "Local in 5", "Local in 6", "USB L", "USB R"],
    "A1-2": ["AES50-A 1", "AES50-A 2", "Aux in 3", "Aux in 4", "Aux in 5", "Aux in 6", "USB L", "USB R"],
    "A1-4": ["AES50-A 1", "AES50-A 2", "AES50-A 3", "AES50-A 4", "Aux in 5", "Aux in 6", "USB L", "USB R"],
    "A1-6": ["AES50-A 1", "AES50-A 2", "AES50-A 3", "AES50-A 4", "AES50-A 5", "AES50-A 6", "USB L", "USB R"],
    "B1-2": ["AES50-B 1", "AES50-B 2", "Aux in 3", "Aux in 4", "Aux in 5", "Aux in 6", "USB L", "USB R"],
    "B1-4": ["AES50-B 1", "AES50-B 2", "AES50-B 3", "AES50-B 4", "Aux in 5", "Aux in 6", "USB L", "USB R"],
    "B1-6": ["AES50-B 1", "AES50-B 2", "AES50-B 3", "AES50-B 4", "AES50-B 5", "AES50-B 6", "USB L", "USB R"],
    "CARD1-2": ["Card 1", "Card 2", "Aux in 3", "Aux in 4", "Aux in 5", "Aux in 6", "USB L", "USB R"],
    "CARD1-4": ["Card 1", "Card 2", "Card 3", "Card 4", "Aux in 5", "Aux in 6", "USB L", "USB R"],
    "CARD1-6": ["Card 1", "Card 2", "Card 3", "Card 4", "Card 5", "Card 6", "USB L", "USB R"],
    "UIN1-2": ["", "", "Aux in 3", "Aux in 4", "Aux in 5", "Aux in 6", "USB L", "USB R"],
    "UIN1-4": ["", "", "", "", "Aux in 5", "Aux in 6", "USB L", "USB R"],
    "UIN1-6": ["", "", "", "", "", "", "USB L", "USB R"]
}

block_routing_lookup_table: dict[tuple] = {
    "AN1-8": ["Local 1", "Local 2", "Local 3", "Local 4", "Local 5", "Local 6", "Local 7", "Local 8"],
    "AN9-16": ["Local 9", "Local 10", "Local 11", "Local 12", "Local 13", "Local 14", "Local 15", "Local 16"],
    "AN17-24": ["Local 17", "Local 18", "Local 19", "Local 20", "Local 21", "Local 22", "Local 23", "Local 24"],
    "AN25-32": ["Local 25", "Local 26", "Local 27", "Local 28", "Local 29", "Local 30", "Local 31", "Local 32"],
    "A1-8": ["AES50-A 1", "AES50-A 2", "AES50-A 3", "AES50-A 4", "AES50-A 5", "AES50-A 6", "AES50-A 7", "AES50-A 8"],
    "A9-16": ["AES50-A 9", "AES50-A 10", "AES50-A 11", "AES50-A 12", "AES50-A 13", "AES50-A 14", "AES50-A 15", "AES50-A 16"],
    "A17-24": ["AES50-A 17", "AES50-A 18", "AES50-A 19", "AES50-A 20", "AES50-A 21", "AES50-A 22", "AES50-A 23", "AES50-A 24"],
    "A25-32": ["AES50-A 25", "AES50-A 26", "AES50-A 27", "AES50-A 28", "AES50-A 29", "AES50-A 30", "AES50-A 31", "AES50-A 32"],
    "A33-40": ["AES50-A 33", "AES50-A 34", "AES50-A 35", "AES50-A 36", "AES50-A 37", "AES50-A 38", "AES50-A 39", "AES50-A 40"],
    "A41-48": ["AES50-A 41", "AES50-A 42", "AES50-A 43", "AES50-A 44", "AES50-A 45", "AES50-A 46", "AES50-A 47", "AES50-A 48"],
    "B1-8": ["AES50-B 1", "AES50-B 2", "AES50-B 3", "AES50-B 4", "AES50-B 5", "AES50-B 6", "AES50-B 7", "AES50-B 8"],
    "B9-16": ["AES50-B 9", "AES50-B 10", "AES50-B 11", "AES50-B 12", "AES50-B 13", "AES50-B 14", "AES50-B 15", "AES50-B 16"],
    "B17-24": ["AES50-B 17", "AES50-B 18", "AES50-B 19", "AES50-B 20", "AES50-B 21", "AES50-B 22", "AES50-B 23", "AES50-B 24"],
    "B25-32": ["AES50-B 25", "AES50-B 26", "AES50-B 27", "AES50-B 28", "AES50-B 29", "AES50-B 30", "AES50-B 31", "AES50-B 32"],
    "B33-40": ["AES50-B 33", "AES50-B 34", "AES50-B 35", "AES50-B 36", "AES50-B 37", "AES50-B 38", "AES50-B 39", "AES50-B 40"],
    "B41-48": ["AES50-B 41", "AES50-B 42", "AES50-B 43", "AES50-B 44", "AES50-B 45", "AES50-B 46", "AES50-B 47", "AES50-B 48"],
    "CARD1-8": ["Card 1", "Card 2", "Card 3", "Card 4", "Card 5", "Card 6", "Card 7", "Card 8"],
    "CARD9-16": ["Card 9", "Card 10", "Card 11", "Card 12", "Card 13", "Card 14", "Card 15", "Card 16"],
    "CARD17-24": ["Card 17", "Card 18", "Card 19", "Card 20", "Card 21", "Card 22", "Card 23", "Card 24"],
    "CARD25-32": ["Card 25", "Card 26", "Card 27", "Card 28", "Card 29", "Card 30", "Card 31", "Card 32"],
    "UIN1-8": ["", 0],
    "UIN9-16": ["", 8],
    "UIN17-24": ["", 16],
    "UIN25-32": ["", 24]
}

override_routing_lookup_table: tuple[str] = [
    "Off",
    "", # 1
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "", # 32
    "Aux 1",
    "Aux 2",
    "Aux 3",
    "Aux 4",
    "Aux 5",
    "Aux 6",
    "USB L",
    "USB R",
    "Fx Rt 1 L",
    "Fx Rt 1 R",
    "Fx Rt 2 L",
    "Fx Rt 2 R",
    "Fx Rt 3 L",
    "Fx Rt 3 R",
    "Fx Rt 4 L",
    "Fx Rt 4 R",
    "Bus 1",
    "Bus 2",
    "Bus 3",
    "Bus 4",
    "Bus 5",
    "Bus 6",
    "Bus 7",
    "Bus 8",
    "Bus 9",
    "Bus 10",
    "Bus 11",
    "Bus 12",
    "Bus 13",
    "Bus 14",
    "Bus 15",
    "Bus 16"
]

DCA_inver_number_lookup_table: tuple[int] = [
    7,
    6,
    5,
    4,
    3,
    2,
    1,
    0
]

user_in_routing_lookup_tabel: tuple[str] = [
    "Off",
    "?",
    "Local 1",
    "Local 2",
    "Local 3",
    "Local 4",
    "Local 5",
    "Local 7",
    "Local 8",
    "Local 9",
    "Local 10",
    "Local 11",
    "Local 12",
    "Local 13",
    "Local 14",
    "Local 15",
    "Local 16",
    "Local 17",
    "Local 18",
    "Local 19",
    "Local 20",
    "Local 21",
    "Local 22",
    "Local 23",
    "Local 24",
    "Local 25",
    "Local 26",
    "Local 27",
    "Local 28",
    "Local 29",
    "Local 30",
    "Local 31",
    "Local 32",
    "AES50-A 1",
    "AES50-A 2",
    "AES50-A 3",
    "AES50-A 4",
    "AES50-A 5",
    "AES50-A 6",
    "AES50-A 7",
    "AES50-A 8",
    "AES50-A 9",
    "AES50-A 10",
    "AES50-A 11",
    "AES50-A 12",
    "AES50-A 13",
    "AES50-A 14",
    "AES50-A 15",
    "AES50-A 16",
    "AES50-A 17",
    "AES50-A 18",
    "AES50-A 19",
    "AES50-A 20",
    "AES50-A 21",
    "AES50-A 22",
    "AES50-A 23",
    "AES50-A 24",
    "AES50-A 25",
    "AES50-A 26",
    "AES50-A 27",
    "AES50-A 28",
    "AES50-A 29",
    "AES50-A 30",
    "AES50-A 31",
    "AES50-A 32",
    "AES50-A 33",
    "AES50-A 34",
    "AES50-A 35",
    "AES50-A 36",
    "AES50-A 37",
    "AES50-A 38",
    "AES50-A 39",
    "AES50-A 40",
    "AES50-A 41",
    "AES50-A 42",
    "AES50-A 43",
    "AES50-A 44",
    "AES50-A 45",
    "AES50-A 46",
    "AES50-A 47",
    "AES50-A 48",
    "AES50-B 1",
    "AES50-B 2",
    "AES50-B 3",
    "AES50-B 4",
    "AES50-B 5",
    "AES50-B 6",
    "AES50-B 7",
    "AES50-B 8",
    "AES50-B 9",
    "AES50-B 10",
    "AES50-B 11",
    "AES50-B 12",
    "AES50-B 13",
    "AES50-B 14",
    "AES50-B 15",
    "AES50-B 16",
    "AES50-B 17",
    "AES50-B 18",
    "AES50-B 19",
    "AES50-B 20",
    "AES50-B 21",
    "AES50-B 22",
    "AES50-B 23",
    "AES50-B 24",
    "AES50-B 25",
    "AES50-B 26",
    "AES50-B 27",
    "AES50-B 28",
    "AES50-B 29",
    "AES50-B 30",
    "AES50-B 31",
    "AES50-B 32",
    "AES50-B 33",
    "AES50-B 34",
    "AES50-B 35",
    "AES50-B 36",
    "AES50-B 37",
    "AES50-B 38",
    "AES50-B 39",
    "AES50-B 40",
    "AES50-B 41",
    "AES50-B 42",
    "AES50-B 43",
    "AES50-B 44",
    "AES50-B 45",
    "AES50-B 46",
    "AES50-B 47",
    "AES50-B 48",
    "Card 1",
    "Card 2",
    "Card 3",
    "Card 4",
    "Card 5",
    "Card 6",
    "Card 7",
    "Card 8",
    "Card 9",
    "Card 10",
    "Card 11",
    "Card 12",
    "Card 13",
    "Card 14",
    "Card 15",
    "Card 16",
    "Card 17",
    "Card 18",
    "Card 19",
    "Card 20",
    "Card 21",
    "Card 22",
    "Card 23",
    "Card 24",
    "Card 25",
    "Card 26",
    "Card 27",
    "Card 28",
    "Card 29",
    "Card 30",
    "Card 31",
    "Card 32",
    "Aux In 1",
    "Aux In 2",
    "Aux In 3",
    "Aux In 4",
    "Aux In 5",
    "Aux In 6",
    "TB",
    "TB"
]

icon_lookup_tabel: tuple[str] = [
    "Igen",
    "Stortromme Forside",
    "Stortromme Bagside",
    "Lilletromme Top",
    "Lilletromme Bund",
    "Høje Tam",
    "Venstre Tam",
    "Gulvtam",
    "Hi-Hat",
    "Bæken",
    "Trommesæt",
    "Ko-Klokke",
    "Bongotrommer 1",
    "Bongotrommer 2",
    "Tamburin",
    "Xylofon",
    "Bas",
    "Guitar 1",
    "Guitar 2",
    "Guitar 3",
    "El Guitar 1",
    "El Guitar 2",
    "Acustisk Guitar",
    "Forstærker 1",
    "Forstærker 2",
    "Forstærker 3",
    "Flygel",
    "Klaver",
    "Keybord 1",
    "Keybord 2",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "Condenser Mikrofon",
    "Lille Condenser Mikrofon L",
    "Lille Condenser Mikrofon R",
    "Dynamisk Mikrofon",
    "Trådløs Mikrofon",
    "Podie Mikrofon",
    "Øresnegl",
    "",
    "",
    "",
    "",
    "",
    "",
    "Kasettebånd",
    "FX",
    "Computer",
    "Monitor",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
    "",
]

colour_lookup_table: dict[str] = {
    "OFF": "Black",
    "RD": "Red",
    "GN": "Green",
    "YE": "Yellow",
    "BL": "Blue",
    "MG": "Magenta",
    "CY": "Cyan",
    "WH": "White",
    "OFFi": "Black",
    "RDi": "Red",
    "GNi": "Green",
    "YEi": "Yellow",
    "BLi": "Blue",
    "MGi": "Magenta",
    "CYi": "Cyan",
    "WHi": "White"
}

output_lookup_table: tuple[tuple[str]] = [
    ["Off", "", "Off"],
    ["main", "st", "L"],
    ["main", "st", "R"],
    ["main", "m", "M/C"],
    ["bus", "01", "Bus 1"],
    ["bus", "02", "Bus 2"],
    ["bus", "03", "Bus 3"],
    ["bus", "04", "Bus 4"],
    ["bus", "05", "Bus 5"],
    ["bus", "06", "Bus 6"],
    ["bus", "07", "Bus 7"],
    ["bus", "08", "Bus 8"],
    ["bus", "09", "Bus 9"],
    ["bus", "10", "Bus 10"],
    ["bus", "11", "Bus 11"],
    ["bus", "12", "Bus 12"],
    ["bus", "13", "Bus 13"],
    ["bus", "14", "Bus 14"],
    ["bus", "15", "Bus 15"],
    ["bus", "16", "Bus 16"],
    ["mtx", "01", "Matrix 1"],
    ["mtx", "02", "Matrix 2"],
    ["mtx", "03", "Matrix 3"],
    ["mtx", "04", "Matrix 4"],
    ["mtx", "05", "Matrix 5"],
    ["mtx", "06", "Matrix 6"],
    
]

# note: str = 'Please Wait'
# print(note)

file_path: str = get_file_path()
if file_path: 
    lines: list[str] = get_lines(file_path)
    get_aux_inputs(lines)
    save_to_excel(get_inputs(lines), get_outputs(lines), get_aux_inputs(lines))