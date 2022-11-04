""" MAC Converter Script """
import tkinter as tk
from tkinter import Button, Entry, filedialog, Label, messagebox
from tkinter.constants import FALSE
import contextlib
import traceback
import pandas as pd
import numpy as np

window = tk.Tk()

# Set global variables. Because I can. Global variables are cool. All the cool kids are using them.
ICON = r"C:\\Users\\nicho\\Desktop\\Dev Projects\\MAC Converter\\logo_TRG.ico"
PAD = '\t\t'
DPAD = '\t\t\t\t'
M_INDEX = 0
ML_LEVEL = []

def open_file() -> str:
    """ Opens a file dialog box and prints path to console """
    open_path = filedialog.askopenfilename(initialdir='/', title='Select MAC', filetypes=(
        ('xlsx files', '*.xlsx'), ('xls files', '*.xls'),
        ('csv files', '*.csv'), ('all files', '*.*')))
    if open_path == '':
        messagebox.showerror('Error', 'No file selected')
    else:
        messagebox.showinfo('File Selected', open_path)
        btn_generate2 =  Button(window, text='Generate XML', \
            command=lambda: generate_xml(open_path), font=('helvetica', 13, 'bold'), \
                width=20, bg='#fcba03', pady=5)
        btn_generate2.grid(column=2, row=2)
        return open_path

def generate_xml(open_path) -> None:
    """ Generates a MAC document in XML format or provides you with an error message """
    wp_id = ent_wp_id.get()
    wp_title = ent_wp_name.get()
    path = filedialog.askdirectory(initialdir="/", title="Select file")
    # print(str(path))
    if wp_id == '':
        # Shows error message
        messagebox.showerror('Missing WP ID', 'Error: WP ID is missing. Please try again!')
    elif wp_title == '':
        # Shows error message
        messagebox.showerror('Missing WP Name', 'Error: WP Name is missing. Please try again!')
    else:
        create_header(wp_id, wp_title, path)
        create_body(open_path, path)
        success_msg = f'Maintenance Allocation Chart (MAC) for the {wp_title} has been converted successfully!\n\nThe XML file is located at:\n\n{path}/{wp_title} MAC.xml'
        messagebox.showinfo('Mission Accomplished!', success_msg)


def get_tools(file_to_grab, sheet) -> str:
    """ Function to grab tools list.
        loads excel file and converts it to a Pandas dataframe. """
    tools_df = pd.read_excel(file_to_grab, sheet, index_col=None, header=1)
    tools_df.insert(1, 'trefcode ID', '')  # Adds the trefcode ID column
    # Adds value for trefcode ID column
    for index, row in tools_df.iterrows():
        teref_id = str(index + 1).rjust(2, '0')
        teref_str = ''
        teref_str += f'MAC_TOOL_{teref_id}'
        tools_df.iloc[index, 1] = teref_str
    return tools_df

def get_rems(rem_xls, sheet) -> str:
    """ Function to grab remarks list
        loads remarks excel file and converts it to a Pandas dataframe. """
    rem_df = pd.read_excel(rem_xls, sheet, index_col=None, header=1)
    rem_df.insert(1, 'remarkcode ID', '')  # adds the remarkcode ID column
    # Adds value for remarkcode ID column
    for index, row in rem_df.iterrows():
        rem_id = str(index + 1).rjust(2, "0")
        rem_str = ''
        rem_str += f'MAC_REM_{rem_id}'
        rem_df.iloc[index, 1] = rem_str
    return rem_df

def create_header(wp_id, wp_title, path) -> None:
    """ Function to create xml header info. """
    header_tmp = '<?xml version="1.0" encoding="UTF-8"?>\n'
    header_tmp += f'<macwp chngno="0" wpno="{wp_id}">\n'
    header_tmp += '\t<wpidinfo>\n'
    header_tmp += f'{PAD}<maintlvl level="maintainer"/>\n'
    header_tmp += f'{PAD}<title>Maintenance Allocation Chart (MAC)</title>\n'
    header_tmp += '\t</wpidinfo>\n'
    header_tmp += '\t<mac>\n'
    header_tmp += f'{PAD}<title>Maintenance Allocation Chart for {wp_title}</title>\n'

    with open(f'{path}/{ent_wp_name.get()} MAC.xml', 'w', encoding="utf-8") as _f:
        _f.write(header_tmp)

def maint_hours(ml_array) -> str:
    """ Function for maintenance hours. """
    # Array, sorry - "Python list" for maintainer levels.
    global ML_LEVEL
    ML_LEVEL = ['c', 'f', 'h', 'd']
    temp_str = ''
    for _m in ml_array:
        if not np.isnan(_m):
            global M_INDEX
            M_INDEX = ml_array.index(_m)  # Gets the index position
            temp_str += f'{DPAD}\t<maintclass-2lvl>\n'
            temp_str += f'{DPAD}{PAD}<{ML_LEVEL[M_INDEX]}>{str(_m)}</{ML_LEVEL[M_INDEX]}>\n'
            temp_str += f'{DPAD}\t</maintclass-2lvl>\n'
    return temp_str

def create_body(open_path, path) -> None:
    """ Creates the body of the MAC XML file. """
    # Loads remarks excel file into a dataframe
    rem_df = get_rems(open_path, 'REMARKS INPUT')
    # Loads tools excel file into a dataframe
    tools_df = get_tools(open_path, 'TOOLS INPUT')
    # Variable that toggles close "compassemgroup" tag opened or closed. "False" means it's open
    close_row = False
    # Loads main excel file and converts it to a dataframe.
    _df = pd.read_excel(open_path, 'MAC INPUT', index_col=None, header=5)
    df_len = len(_df)

    for index, row in _df.iterrows():
        temp_str = ""
        # Create array for maintainer levels - c,f,h,d
        ml_array = [row[3], row[4], row[5], row[6]]
        # Assigns value of rows 7 and 8 to a variable; I had to do this because
        # otherwise it throws up when you try to check for "isnan". Now you can
        # ask it to look for the string 'nan', which is kludgy as hell but it works.
        teref = str(row[7])
        remark_refs = str(row[8])

        if not np.isnan(row[0]) and close_row:
            temp_str += f'{DPAD}</qualify-2lvl>\n'
            temp_str += f'{PAD}\t</compassemgroup-2lvl>\n'
            temp_str += f'{PAD}</mac-group-2lvl>\n'

            # Takes the stupid period out of the cast string. It does this by dropping the row[0]
            # value into a variable and pulling the period out from there.
            group_no = str(int(row[0]))
            group_no = group_no.replace('.', '')

            # Adds a zero in front of the group number, unless the group number is "00"
            # Throws a leading zero in front of the groupno value. Because stupid leading zeros.
            if group_no != "00":
                group_no = f"0{group_no}"
            temp_str += f'{PAD}<mac-group-2lvl>\n'
            temp_str += f'{PAD}\t<groupno>{group_no}</groupno>\n'
            temp_str += f"{PAD}\t<compassemgroup-2lvl>\n"
            temp_str += f"{PAD}{PAD}<compassem>\n"
            temp_str += f'{PAD}{PAD}\t<name>{str(row[1])}</name>\n'
            temp_str += f'{DPAD}</compassem>\n'
            close_row = True

        temp_str += DPAD + '<qualify-2lvl>\n'
        # Added if/else statement to remove the func="nan" results from the xml file.
        if str(row[2]) != 'nan':
            temp_str += f'{DPAD}\t<maintfunc func="{str(row[2]).lower()}"/>\n' # original line
        else:
            temp_str += f'{DPAD}\t<maintfunc func="none"/>\n'
            temp_str += f'{DPAD}\t<maintclass-2lvl>\n'
            temp_str += f'{DPAD}{PAD}<{ML_LEVEL[M_INDEX]}/>\n'
            temp_str += f'{DPAD}\t</maintclass-2lvl>\n'

        # Runs maint_hours function on ml_array list
        main_str = maint_hours(ml_array)
        temp_str += main_str  # Adds returned values to temp_str

        if teref != "nan":
            temp_str += f'{DPAD}\t<terefs>\n'
            tools_list = teref.split(",")

            for _e in tools_list:
                # try:
                _e = int(_e)
                abd = tools_df.loc[tools_df['TOOLS OR TEST\nEQUIPMENT\nREF CODE'] == _e]
                abd = abd.values.flatten()
                temp_str += f'{DPAD}{PAD}<teref refs="{abd[1]}"/>\n'
                # except: # pylint: disable=bare-except
                #     traceback.print_exc()
                #     # pass
            temp_str += f'{DPAD}\t</terefs>\n'

        if remark_refs != "nan":
            # print(f'Number: {row[0]} Name: {row[1]}')
            rem_list = remark_refs.split(",")
            temp_str += f'{DPAD}\t<remarkrefs>\n'
            for _r in rem_list:
                # try:
                # print(_r)
                _rc = rem_df.loc[rem_df['REMARK CODE'] == _r]
                _rc = _rc.values.flatten()
                temp_str += f'{DPAD}{PAD}<remarkref refs="_rc[1]"/>\n'
                # except: # pylint: disable=bare-except
                #     traceback.print_exc()
                #     # pass
            temp_str += f'{DPAD}\t</remarkrefs>\n'

        if index + 1 < df_len:
            next_row = _df.iloc[index + 1]
            if np.isnan(next_row[0]):
                temp_str += f'{DPAD}</qualify-2lvl>\n'

        with open(f'{path}/{ent_wp_name.get()} MAC.xml', 'a', encoding="utf-8") as _f:
            _f.write(temp_str)

    # Creates tool list from existing tools_df
    end_str = f'{DPAD}</qualify-2lvl>\n'
    end_str += f'{PAD}\t</compassemgroup-2lvl>\n'
    end_str += f'{PAD}</mac-group-2lvl>\n'
    end_str += '\t</mac>\n'
    end_str += '\t<?Pub _newpage?>\n'
    end_str += '\t<tereqtab>\n'
    end_str += f'{PAD}<title>Tools and Test Equipment for {ent_wp_name.get()}</title>\n'

    for index, row in tools_df.iterrows():
        nsn = str(row[4])  # Converts the NSN (which is an int) to a string.
        if nsn != 'nan':
            # Grabs the first four characters of the NSN and convert it to a string called "fsc"
            fsc = nsn[:3]
            # Checks to see if python has chopped the leading zero. Which it always does.
            if len(fsc) < 4:
                fsc = f'0{fsc}'  # Puts the damn leading zero back.
            niin = nsn[5:]
        else:
            fsc = ''
            niin = ''
        maint_lvl = str(row[2])
        maint_lvl = maint_lvl.lower()
        end_str += f'{PAD}<teref-group>\n'
        end_str += f'{PAD}\t<terefcode id="{row[1]}">{(str(row[0])).zfill(2)}</terefcode>\n'
        end_str += f'{PAD}\t<maintenance lvl="{maint_lvl.lower()}"/>\n'
        end_str += f'{PAD}\t<name>{row[3]}</name>\n'
        end_str += f'{PAD}\t<nsn>\n'
        end_str += f'{DPAD}<fsc>{fsc}</fsc>\n'
        end_str += f'{DPAD}<niin>{niin}</niin>\n'
        end_str += f'{PAD}\t</nsn>\n'
        end_str += f'{PAD}\t<toolno>{str(row[5])}</toolno>\n'
        end_str += f'{PAD}</teref-group>\n'
    end_str += '\t</tereqtab>\n'

    with open(f'{path}/{ent_wp_name.get()} MAC.xml', 'a', encoding="utf-8") as end:
        end.write(end_str)

    # Uses remarks spreadsheet to create and write xml in for remarks section.
    rem_str = '\n\t<?Pub _newpage?>\n' + '\t<remarktab>\n'
    rem_str += f'{PAD}<title>Remarks for {ent_wp_name.get()}</title>\n'

    for index, row in rem_df.iterrows():
        rem_id = (str(index+1).rjust(2, '0'))
        rem_str += f'{PAD}<remark-group>\n'
        rem_str += f'{PAD}\t<remarkcode id="MAC_REM_{rem_id}">{row[0]}</remarkcode>\n'
        rem_str += f'{PAD}\t<remarks>{row[2]}</remarks>\n'
        rem_str += f'{PAD}</remark-group>\n'
    rem_str += '\t</remarktab>\n'
    rem_str += '</macwp>\n'

    with open(f'{path}/{ent_wp_name.get()} MAC.xml', 'a', encoding="utf-8") as rem:
        rem.write(rem_str)

window.title("Mark's MAC Converter")
window.geometry('560x160')
window.config(bg='#bdcff0')
window.resizable(width=FALSE, height=FALSE)

with contextlib.suppress(tk.TclError):
    window.iconbitmap(ICON)

lbl_wpid = Label(window, text='WP ID: ', font=('helvetica', 13, 'bold',), pady=5, bg='#bdcff0')
lbl_wpid.grid(column=0, row=0)

lbl_wp_name = Label(window, text='WP Name: ', font=('helvetica', 13, 'bold'), pady=5, bg='#bdcff0')
lbl_wp_name.grid(column=0, row=1)

ent_wp_id = Entry(window, width=40)
ent_wp_id.grid(column=1, row=0)

ent_wp_name = Entry(window, width=40)
ent_wp_name.grid(column=1, row=1)

btn_open =  Button(window, text='Select MAC', command=open_file, font=(
    'helvetica', 13, 'bold'), width=20, bg='#fcba03', pady=5)
btn_open.grid(column=1, row=2)

window.mainloop()
