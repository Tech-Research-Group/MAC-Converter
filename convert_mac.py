""" MAC Converter Script """
import tkinter as tk
from tkinter import Button, Entry, filedialog, Label, messagebox
import pandas as pd
import numpy as np

window = tk.Tk()

# Set global variables. Because I can. Global variables are cool. All the cool kids are using them.
PAD = '\t\t'
DPAD = '\t\t\t\t'

def open_file() -> str:
    """ Opens a file dialog box and prints path to console """
    open_path = filedialog.askopenfilename(initialdir='/', title='Select File', filetypes=(
        ('xlsx files', '*.xlsx'), ('xls files', '*.xls'),
        ('csv files', '*.csv'), ('all files', '*.*')))
    messagebox.showinfo('File Selected', open_path)
    btn_generate2 =  Button(window, text='Generate XML', command=lambda: generate_xml(open_path), font=(
    'helvetica', 13, 'bold'), width=20, bg='#fcba03', pady=5)
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
        success_msg = 'Maintenance Allocation Chart (MAC) for the ' + wp_title + \
        ' has been converted successfully!\n\n' + 'The XML file is located at:\n\n' + \
            path + '/' + wp_title + ' MAC.xml'
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
        teref_str += 'MAC_TOOL_' + teref_id
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
        rem_str += 'MAC_REM_' + rem_id
        rem_df.iloc[index, 1] = rem_str
    return rem_df

def create_header(wp_id, wp_title, path) -> None:
    """ Function to create xml header info. """
    header_tmp = '<?xml version="1.0" encoding="UTF-8"?>\n'
    header_tmp += '<macwp chngno="0" wpno="' + wp_id + '">\n'
    header_tmp += '\t<wpidinfo>\n'
    header_tmp += PAD + '<maintlvl level="maintainer"/>\n'
    header_tmp += PAD + '<title>Maintenance Allocation Chart (MAC)</title>\n'
    header_tmp += '\t</wpidinfo>\n'
    header_tmp += '\t<mac>\n'
    header_tmp += PAD + '<title>Maintenance Allocation Chart (MAC) for ' + \
        wp_title + '</title>\n'

    with open(f'{path}/{ent_wp_name.get()} MAC.xml', 'w', encoding="utf-8") as _f:
        _f.write(header_tmp)

def maint_hours(ml_array):
    """ Function for maintenance hours. """
    # Array, sorry - "Python list" for maintainer levels.
    global ml_level
    ml_level = ['c', 'f', 'h', 'd']
    temp_str = ''
    for _m in ml_array:
        if np.isnan(_m) == False:
            global m_index
            m_index = ml_array.index(_m)  # Gets the index position
            temp_str += DPAD + '\t<maintclass-2lvl>\n'
            temp_str += DPAD + PAD + '<' + ml_level[m_index] + '>' + str(_m) + '</' + ml_level[m_index] + '>\n'
            temp_str += DPAD + '\t</maintclass-2lvl>\n'
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

        if np.isnan(row[0]) == False:
            if close_row == True:
                temp_str += DPAD + '</qualify-2lvl>\n'
                temp_str += PAD + '\t</compassemgroup-2lvl>\n'
                temp_str += PAD + '</mac-group-2lvl>\n'

            # Takes the stupid period out of the cast string. It does this by dropping the row[0]
            # value into a variable and pulling the period out from there.
            group_no = str(int(row[0]))
            group_no = group_no.replace('.', '')

            # Adds a zero in front of the group number, unless the group number is "00"
            # Throws a leading zero in front of the groupno value. Because stupid leading zeros.
            if group_no != "00":
                group_no = "0" + group_no
            temp_str += PAD + "<mac-group-2lvl>\n"
            temp_str += PAD + '\t<groupno>' + group_no + '</groupno>\n'
            temp_str += PAD + "\t<compassemgroup-2lvl>\n"
            temp_str += PAD + PAD + "<compassem>\n"
            temp_str += PAD + PAD + "\t<name>" + str(row[1]) + "</name>\n"
            temp_str += DPAD + '</compassem>\n'
            close_row = True

        temp_str += DPAD + '<qualify-2lvl>\n'
        # Added if/else statement to remove the func="nan" results from the xml file.
        if str(row[2]) != 'nan':
            temp_str += DPAD + '\t<maintfunc func="' + str(row[2]).lower() + '"/>\n' # original line
        else:
            temp_str += DPAD + '\t<maintfunc func="none"/>\n'
            temp_str += DPAD + '\t<maintclass-2lvl>\n'
            temp_str += DPAD + PAD + '<' + ml_level[m_index] + '/>\n'
            temp_str += DPAD + '\t</maintclass-2lvl>\n'

        # Runs maint_hours function on ml_array list
        main_str = maint_hours(ml_array)
        temp_str += main_str  # Adds returned values to temp_str

        if teref != "nan":
            temp_str += DPAD + '\t<terefs>\n'
            tools_list = teref.split(",")

            for _e in tools_list:
                try:
                    _e = int(_e)
                    abd = tools_df.loc[tools_df['TOOLS OR TEST\nEQUIPMENT\nREF CODE'] == _e]
                    abd = abd.values.flatten()
                    temp_str += DPAD + PAD + '<teref refs="' + abd[1] + '"/>\n'
                except: # pylint: disable=bare-except
                    #traceback.print_exc()
                    pass
            temp_str += DPAD + '\t</terefs>\n'

        if remark_refs != "nan":
            # print(f'Number: {row[0]} Name: {row[1]}')
            rem_list = remark_refs.split(",")
            temp_str += DPAD + '\t<remarkrefs>\n'
            for _r in rem_list:
                try:
                    # print(_r)
                    _rc = rem_df.loc[rem_df['REMARK CODE'] == _r]
                    _rc = _rc.values.flatten()
                    temp_str += DPAD + PAD + '<remarkref refs="' + _rc[1] + '"/>\n'
                except: # pylint: disable=bare-except
                    # traceback.print_exc()
                    pass
            temp_str += DPAD + '\t</remarkrefs>\n'

        if index + 1 < df_len:
            next_row = _df.iloc[index + 1]
            if np.isnan(next_row[0]) == True:
                temp_str += DPAD + '</qualify-2lvl>\n'

        with open(f'{path}/{ent_wp_name.get()} MAC.xml', 'a', encoding="utf-8") as _f:
            _f.write(temp_str)

    # Creates tool list from existing tools_df
    end_str = DPAD + '</qualify-2lvl>\n'
    end_str += PAD + '\t</compassemgroup-2lvl>\n'
    end_str += PAD + '</mac-group-2lvl>\n'
    end_str += '\t</mac>\n'
    end_str += '\t<?Pub _newpage?>\n'
    end_str += '\t<tereqtab>\n'
    end_str += PAD + '<title>Tools and Test Equipment for ' + ent_wp_name.get() + '</title>\n'

    for index, row in tools_df.iterrows():
        nsn = str(row[4])  # Converts the NSN (which is an int) to a string.
        if nsn != 'nan':
            # Grabs the first four characters of the NSN and convert it to a string called "fsc"
            fsc = nsn[0:3]
            # Checks to see if python has chopped the leading zero. Which it always does.
            if len(fsc) < 4:
                fsc = '0' + fsc  # Puts the damn leading zero back.
            niin = nsn[5:]
        else:
            fsc = ''
            niin = ''
        maint_lvl = str(row[2])
        maint_lvl = maint_lvl.lower()
        end_str += PAD + '<teref-group>\n'
        end_str += PAD + '\t<terefcode id="' + row[1] + '">' + (str(row[0])).zfill(2) + \
            '</terefcode>\n'
        end_str += PAD + '\t<maintenance lvl="' + maint_lvl.lower() + '"/>\n'
        end_str += PAD + '\t<name>' + row[3] + '</name>\n'
        end_str += PAD + '\t<nsn>\n'
        end_str += DPAD + '<fsc>' + fsc + '</fsc>\n'
        end_str += DPAD + '<niin>' + niin + '</niin>\n'
        end_str += PAD + '\t</nsn>\n'
        end_str += PAD + '\t<toolno>' + str(row[5]) + '</toolno>\n'
        end_str += PAD + '</teref-group>\n'
    end_str += '\t</tereqtab>\n'

    with open(f'{path}/{ent_wp_name.get()} MAC.xml', 'a', encoding="utf-8") as end:
        end.write(end_str)

    # Uses remarks spreadsheet to create and write xml in for remarks section.
    rem_str = '\n'
    rem_str += '\t<?Pub _newpage?>\n'
    rem_str += '\t<remarktab>\n'
    rem_str += PAD + '<title>Remarks for ' + ent_wp_name.get() + '</title>\n'

    for index, row in rem_df.iterrows():
        rem_id = (str(index+1).rjust(2, '0'))
        rem_str += PAD + '<remark-group>\n'
        rem_str += PAD + '\t<remarkcode id="MAC_REM_' + rem_id + '">' + row[0] + '</remarkcode>\n'
        rem_str += PAD + '\t<remarks>' + row[2] + '</remarks>\n'
        rem_str += PAD + '</remark-group>\n'
    rem_str += '\t</remarktab>\n'
    rem_str += '</macwp>\n'

    with open(f'{path}/{ent_wp_name.get()} MAC.xml', 'a', encoding="utf-8") as rem:
        rem.write(rem_str)

window.title("Mark's MAC Converter")
window.geometry('560x160')
window.config(bg='#bdcff0')

lbl_wpid = Label(window, text='WP ID: ', font=('helvetica', 13, 'bold',), pady=5, bg='#bdcff0')
lbl_wpid.grid(column=0, row=0)

lbl_wp_name = Label(window, text='WP Name: ', font=('helvetica', 13, 'bold'), pady=5, bg='#bdcff0')
lbl_wp_name.grid(column=0, row=1)

ent_wp_id = Entry(window, width=40)
# ent_wp_id.insert(0, 'S00003-9-4120-434')
ent_wp_id.grid(column=1, row=0)

ent_wp_name = Entry(window, width=40)
# ent_wp_name.insert(0, '36K IECU')
ent_wp_name.grid(column=1, row=1)

btn_open =  Button(window, text='Select File', command=open_file, font=(
    'helvetica', 13, 'bold'), width=20, bg='#fcba03', pady=5)
btn_open.grid(column=1, row=2)

window.mainloop()
