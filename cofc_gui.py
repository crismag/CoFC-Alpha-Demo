import sys
import tkinter as tk
from tkinter import filedialog, ttk, scrolledtext
import lxml.etree as ET
from xml.dom import minidom
import io
import os
import re

script_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.append(script_dir)

from CoFCLib.COFCExcelCriticalDimensionWorksheetReader import COFCExcelCriticalDimensionWorksheetReader
from CoFCLib.COFCExcelRegistrationWorksheetReader import COFCExcelRegistrationWorksheetReader
from CoFCLib.COFCExcelStatisticsWorksheetReader import COFCExcelStatisticsWorksheetReader
from CoFCLib.COFCExcelPhaseShiftTransmissionWorksheetReader import COFCExcelPhaseShiftTransmissionWorksheetReader
from CoFCLib.COFCXMLParser import COFCXMLParser
from CoFCLib.COFCExcelCopyProtect import COFCCopyProtect
from CoFCLib import plmget

stdout_buffer = io.StringIO()
stderr_buffer = io.StringIO()
#sys.stdout = stdout_buffer
#sys.stderr = stderr_buffer

global input_file, output_file, infile, outfile
global file_path, output_entry, uname_entry, passw_entry, docn_entry, filename_entry, submit_button
global input_entry, browse_button, output_entry, runparser_button
global copier
checkboxes = []

# Variables for Excel file extraction
global load_critical_dimension_data
global load_registration_data
global load_statistics_data
global load_pst_data
global selected_worksheets
selected_worksheets = []
load_critical_dimension_data = 1
load_registration_data = 0
load_statistics_data = 0
load_pst_data = 0


def exit_program():
    ui_main.quit()

def browse_input_file():
    global file_path, input_entry, input_entry, input_file, outfile
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls;*.xml")])
    input_entry.delete(0, tk.END)
    input_entry.insert(0, file_path)
    infile = input_entry.get()
    outfile = infile
    if is_excel_file(infile):
        outname = re.sub(r'.xls$','', infile)
        outname = re.sub(r'.xlsx$','', outname)
        outname = outname + '.parsed_excel.xml'
    elif is_xml_file(infile):
        outname = re.sub(r'.xml$','', infile)
        outname = outname + '.cdxml.xml'
    else:
        outname = re.sub(r'.xls$','', infile)
        outname = re.sub(r'.xlsx$','', outname)
        outname = re.sub(r'.xml$','', outname)
        outname = outname + '.parsed_unkn.xml'
    update_entry_value('output_entry', outname)


def browse_excel_file():
    global file_path, in_ent, ou_ent, ou_en2
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls;")])
    in_ent.delete(0, tk.END)
    in_ent.insert(0, file_path)
    in_xls = in_ent.get()
    ou_xls = re.sub(r'.xls$','', in_xls)
    ou_xls = re.sub(r'.xlsx$','', ou_xls)
    ou_xls = ou_xls + '.worksheet_copy.xls'
    update_entry_value('ou_ent', ou_xls)

    ou_xml = re.sub(r'.xls$','', in_xls)
    ou_xml = re.sub(r'.xlsx$','', ou_xml)
    ou_xml = ou_xml + '.parsed.xml'
    update_entry_value('ou_en2', ou_xml)
    get_worksheets()


def process_excel(input_file):
    global output_file
    global load_critical_dimension_data
    global load_registration_data
    global load_statistics_data
    global load_pst_data
    global selected_worksheets
    selected_worksheets = []
    for cb in checkboxes:
        if cb[1].get() == 1:
            selected_worksheets.append(cb[2])
    if len(selected_worksheets) > 0:
        load_statistics_data = 0
        load_registration_data = 0
        load_critical_dimension_data = 0
        load_pst_data = 0

    for ws in selected_worksheets:
        if ws == "Critical Dimension":
            load_critical_dimension_data = 1
        if ws == "Registration":
            load_registration_data = 1
        if ws == "Statistics":
            load_statistics_data = 1
        if ws == "Phase_Shift_Transmission":
            load_pst_data = 1

    output_file = output_entry.get()
    xml_strings = []
    selected_nodes = []
    if load_critical_dimension_data:
        print("processing CD:")
        cofc_cd = COFCExcelCriticalDimensionWorksheetReader(input_file)
        cofc_cd.open_workbook()
        cofc_cd.create_category_tables()
        cofc_cd.generate_json_data()
        cofc_cd.cd_xml = cofc_cd.critical_dimension_json_to_xml('Critical_Dimension')
        log_display(cofc_cd.cd_xml)
        xml_strings.append(cofc_cd.cd_xml)
        selected_nodes.append('//Critical_Dimension')
        print("Done processing CD:")

    if load_registration_data:
        print("Processing REG:")
        cofc_reg = COFCExcelRegistrationWorksheetReader(input_file)
        cofc_reg.open_workbook()
        cofc_reg.generate_registration_data()
        cofc_reg.reg_xml = cofc_reg.regdata_to_xml('Registration')
        log_display(cofc_reg.reg_xml)
        xml_strings.append(cofc_reg.reg_xml)
        selected_nodes.append('//Registration')
        print("Done processing REG:")

    if load_statistics_data:
        print("Processing Stat:", input_file)
        cofc_stat = COFCExcelStatisticsWorksheetReader(input_file)
        cofc_stat.open_workbook()
        cofc_stat.process_table_groups()
        cofc_stat.generate_json_data()
        cofc_stat.cd_xml = cofc_stat.statistics_json_to_xml('Statistics')
        log_display(cofc_stat.cd_xml)
        xml_strings.append(cofc_stat.cd_xml)
        selected_nodes.append('//Statistics')
        print("Done processing Stat:")

    if load_pst_data:
        print("Processing PST:")
        cofc_pst = COFCExcelPhaseShiftTransmissionWorksheetReader(input_file)
        cofc_pst.open_workbook()
        cofc_pst.create_category_tables()
        cofc_pst.generate_json_data()
        cofc_pst.cd_xml = cofc_pst.phase_shift_transmission_json_to_xml('Phase_Shift_Transmission')
        log_display(cofc_pst.cd_xml)
        xml_strings.append(cofc_pst.cd_xml)
        selected_nodes.append('//Phase_Shift_Transmission')
        print("Done processing PST:")

    ## Specify the selected root or child nodes to extract
    #selected_nodes = ['//Critical_Dimension', '//Registration', '//Phase_Shift_Transmission', '//Statistics']

    log_display_buff()
    # Specify the new root name for the combined XML
    new_root_name = 'COFC_Combined'
    combine_selected_nodes(xml_strings, output_file, selected_nodes, new_root_name)

    log_display(f"Success: Excel file processed and XML saved successfully.")
    log_display(f"Outfile={output_file}")
    log_display_buff()


# Function to combine selected root or child nodes from XML strings
def combine_selected_nodes(xml_strings, output_file, xpath_queries, new_root_name):
    # Create the root element for the combined XML
    root = ET.Element(new_root_name)

    # Iterate over each XML string
    for xml_string in xml_strings:
        # Parse the XML string
        file_root = ET.fromstring(xml_string)

        # Iterate over selected nodes
        for xpath_query in xpath_queries:
            # Find the selected nodes in the file's root
            selected_elements = file_root.xpath(xpath_query)
            if selected_elements:
                # Append the selected nodes to the combined XML root
                root.extend(selected_elements)

    # Create the combined XML tree
    combined_tree = ET.ElementTree(root)
    # Write the combined XML to a file
    combined_tree.write(output_file)


def process_xml(infile):
    global output_file, input_entry, output_entry
    input_file = infile
    output_file = output_entry.get()
    log_display('# ---------------------------- Processing CoFC XML --------------------------------')
    log_display(input_file)
    log_display(output_file)
    config_file = "config/OS-006768-01_24.xpaths.json"
    log_path = "./logs/"
    px = COFCXMLParser(input_file,output_file, config_file, log_path)
    px.convert_xml()
    log_display_buff()
    log_display(f"Success: CoFC XML file processed and XML saved successfully.")
    log_display(f"Outfile={output_file}")


def is_excel_file(file_path):
    file_extension = os.path.splitext(file_path)[1].lower()
    if file_extension == ".xlsx" or file_extension == ".xls":
        return True
    else:
        return False


def is_xml_file(file_path):
    file_extension = os.path.splitext(file_path)[1].lower()
    if file_extension == ".xml":
        return True
    else:
        return False


def process_data_file(in_src):
    global infile, input_entry, output_entry, output_file, in_ent, pf_ent
    if in_src == 'input_file':
        infile = input_entry.get()
    if in_src == 'in_ent':
        infile = in_ent.get()

    pf_content = pf_ent.get()
    print("PFC:",pf_content)
    input_pattens = pf_content.split(' ')

    output_file = output_entry.get()
    if os.path.exists(infile):
        pass
    else:
        log_display('Error: input file missing/does not exists.')
        return

    if is_excel_file(infile):
        log_display('# ---------------------------- Parsing Input Excel File --------------------------------')
        process_excel(infile)
    elif is_xml_file(infile):
        process_xml(infile)
    else:
        log_display('# ---------------------------- Process Data File --------------------------------')
        log_display('Error: Unsupported input file.')
    pass


def update_entry_value(wgt,value):
    entry = globals()[wgt]
    entry.delete(0, tk.END)
    entry.insert(0, value)

def get_plm_doc():
    global docn, docn_entry, uname_entry, passw_entry, submit_button
    user = uname_entry.get()
    passwd = passw_entry.get()
    docn = docn_entry.get()
    file = filename_entry.get()
    log_display('# ---------------------------- Downloading --------------------------------')
    submit_button.config(state=tk.DISABLED)
    log_display("Downloading files from PLM... Please wait")
    sc = plmget.SoapClient(
        wsdl_url="http://agileplm.gfoundries.com:7001/Agile/extension/GetXmlFilesFromChanageWS?wsdl",
        username=user,
        password=passwd,
    )
    action = 'getFiles'
    plm_request = sc.plm_build_request(
        action=action,
        docNum=docn,
        fnames=file
    )
    response = sc.send_request(action, plm_request)

    ui_main.after(500, submit_button.config(state=tk.NORMAL))
    log_display_buff()
    if response is not None:
        sc.extract_value(response)
    log_display_buff()

    if sc.downloaded_file:
        update_entry_value('input_entry', sc.downloaded_file)
    else:
        log_display_buff()
        return

    outname = sc.downloaded_file
    if is_excel_file(sc.downloaded_file):
        outname = re.sub(r'.xls$','', outname)
        outname = re.sub(r'.xlsx$','', outname)
        outname = outname + '.parsed_excel.xml'
    elif is_xml_file(sc.downloaded_file):
        outname = re.sub(r'.xml$','', outname)
        outname = outname + '.cdxml.xml'
    else:
        outname = re.sub(r'.xls$','', outname)
        outname = re.sub(r'.xlsx$','', outname)
        outname = re.sub(r'.xml$','', outname)
        outname = outname + '.parsed_unkn.xml'

    update_entry_value('output_entry', outname)
    sc.downloaded_file = None

    pass


def clear_logs():
    global logs_text
    logs_text.delete('1.0', 'end')


def log_display_buff():
    global logs_text
    stdout_contents = stdout_buffer.getvalue()
    stderr_contents = stderr_buffer.getvalue()
    print(stdout_contents)
    print(stderr_contents)

    if stdout_contents != '' and stdout_contents is not None:
        logs_text.insert(tk.END, stdout_contents+"\n")
        stdout_buffer.truncate(0)
    if stderr_contents != '' or stderr_contents is not None:
        logs_text.insert(tk.END, stderr_contents+"\n")
        stderr_buffer.truncate(0)
    logs_text.see(tk.END)


def log_display(text):
    global logs_text
    if text is None or "":
        return
    logs_text.insert(tk.END, text+"\n")
    logs_text.see(tk.END)


def colorized_xml(xml_text):
    dom = minidom.parseString(xml_text)
    colxml = dom.toprettyxml(indent="    ")
    return colxml


def get_worksheets():
    global worksheets, copier

    use_pandas_only = False
    excel_infile: str = in_ent.get()
    excel_outfile = ou_ent.get()
    if excel_infile is None or excel_infile == "":
        return
    sheets_of_interest = ''
    copier = COFCCopyProtect(excel_infile, excel_outfile, sheets_of_interest, use_pandas_only)
    worksheets = copier.get_worksheets()
    ui_main.update_idletasks()
    return worksheets


def select_all_cboxes():
    for checkbox in checkboxes:
        checkbox[1].set(1)

    ui_main.update_idletasks()


def reload_worksheets():
    worksheets = get_worksheets()

    # Remove existing checkboxes
    for checkbox in checkboxes:
        checkbox[0].destroy()
    create_checkboxes(worksheets)
    ui_main.update_idletasks()
    return


def create_checkboxes(worksheets):
    global sheet_cbox_fr
    if worksheets is None or len(worksheets) == 0:
        worksheets = ['Statistics']
    cbox_ctr = 0
    for i, worksheet in enumerate(worksheets):
        checkbox = tk.IntVar()
        if worksheet == "Statistics":
            checkbox.set(1)  # Select "Statistics" by default
        cb = ttk.Checkbutton(sheet_cbox_fr, text=worksheet, variable=checkbox)
        cb.grid(row=i // 6, column=i % 6, sticky='w')
        cbox_ctr +=1
        checkboxes.append((cb, checkbox, worksheet))

    new_height = (cbox_ctr // 6 + 1) * 14
    sheet_cbox_fr.configure(height=new_height)
    ui_main.update_idletasks()
    ui_main.update()

def copy_worksheet():
    global copier, selected_worksheets
    selected_worksheets = []
    for cb in checkboxes:
        if cb[1].get() == 1:
            selected_worksheets.append(cb[2])
    selected_worksheets = list(set(selected_worksheets))
    #print(selected_worksheets)
    copier.sheets_of_interest = selected_worksheets
    log_display(copier.copy_worksheets())
    log_display(f"Copying worksheets: done")
    ui_main.update_idletasks()


def nb_create_page1():
    global uname_entry, passw_entry, docn_entry, filename_entry, file_path
    page1 = ttk.Frame(notebook)
    notebook.add(page1, text="Get PLM Files")

    # -------------------------------------------------
    # PLM Download page
    plm_frame = ttk.LabelFrame(page1, text='PLM Doc Download')
    plm_frame.pack(side='top', padx=8, pady=2, anchor='nw', expand=True, fill='x')

    user_label = ttk.Label(plm_frame, text='Username:')
    pass_label = ttk.Label(plm_frame, text='Password:')
    docnum_label = ttk.Label(plm_frame, text='Document Number:')
    filename_label = ttk.Label(plm_frame, text='Download File:')

    uname_entry = ttk.Entry(plm_frame)
    passw_entry = ttk.Entry(plm_frame, show='*')
    docn_entry = ttk.Entry(plm_frame)
    filename_entry = ttk.Entry(plm_frame, width=100)
    submit_button = ttk.Button(plm_frame, text='Submit', command=get_plm_doc)

    # Grid Layout - PLM Doc Download.
    user_label.grid(row=0, column=0, padx=4, sticky='nw')
    uname_entry.grid(row=0, column=1, padx=4, sticky='nw')
    pass_label.grid(row=1, column=0, padx=4, sticky='nw')
    passw_entry.grid(row=1, column=1, padx=4, sticky='nw')
    docnum_label.grid(row=2, column=0, padx=4, sticky='nw')
    docn_entry.grid(row=2, column=1, padx=4, sticky='nw')
    filename_label.grid(row=3, column=0, padx=4, sticky='nw')
    filename_entry.grid(row=3, column=1, padx=4, sticky='nw')
    submit_button.grid(row=4, column=0, columnspan=2, pady=10, sticky='nw')
    ui_main.update_idletasks()


def nb_create_page2():
    global input_entry, browse_button, output_entry, runparser_button
    page2 = ttk.Frame(notebook)
    notebook.add(page2, text="Process CofC XML")

    # Create the frame for input and output parser widgets
    parser_frame = ttk.LabelFrame(page2, text='Parser XML/Excel')
    parser_frame.pack(side='top', padx=8, pady=2, anchor='nw', expand=True, fill='x')

    # Create the input file widgets
    input_label = ttk.Label(parser_frame, text="Input File:")
    input_entry = ttk.Entry(parser_frame, width=150)
    browse_button = ttk.Button(parser_frame, text="Browse", command=browse_input_file)
    # Create the output file widgets
    output_label = ttk.Label(parser_frame, text="Output XML:")
    output_entry = ttk.Entry(parser_frame, width=150)
    runparser_button = ttk.Button(parser_frame, text="Parse XML/Excel", command=lambda: process_data_file('input_file'))

    input_label.grid(row=0, column=0, padx=4)
    input_entry.grid(row=0, column=1, padx=4)
    browse_button.grid(row=0, column=2, padx=4)
    output_label.grid(row=1, column=0, padx=4)
    output_entry.grid(row=1, column=1, padx=4)
    runparser_button.grid(row=2, column=1, padx=4)
    ui_main.update_idletasks()


def nb_create_page3():
    global in_ent, ou_ent, ou_en2, sheet_fr, sheet_cbox_fr, pf_ent
    page3 = ttk.Frame(notebook)
    notebook.add(page3, text="Process CoFC Excel")

    # Create the frame for input and output parser widgets
    p_fr = ttk.LabelFrame(page3, text='Excel Copy Protected Worksheet')
    p_fr.pack(side='top', padx=8, pady=2, anchor='nw', expand=False, fill='x')

    # Create the input file widgets
    in_lbl = ttk.Label(p_fr, text="Source ExcelFile:")
    in_ent = ttk.Entry(p_fr, width=150)
    brobtn = ttk.Button(p_fr, text="Browse", command=browse_excel_file)
    # Create the output file widgets
    ou_lbl = ttk.Label(p_fr, text="Output Excel:")
    ou_ent = ttk.Entry(p_fr, width=150)
    #ws_lbl = ttk.Label(p_fr, text="WorkSheet Name:")

    ou_lb2 = ttk.Label(p_fr, text="Output XML:")
    ou_en2 = ttk.Entry(p_fr, width=150)

    in_lbl.grid(row=0, column=0, padx=4)
    in_ent.grid(row=0, column=1, padx=4)
    brobtn.grid(row=0, column=2, padx=4)
    ou_lbl.grid(row=1, column=0, padx=4)
    ou_ent.grid(row=1, column=1, padx=4)
    ou_lb2.grid(row=2, column=0, padx=4)
    ou_en2.grid(row=2, column=1, padx=4)
    #ws_lbl.grid(row=2, column=0, padx=4)

    # Create the frame for input and output parser widgets
    b_fr = ttk.LabelFrame(page3, text='WorkSheets')
    b_fr.pack(side='top', padx=8, pady=2, anchor='nw', expand=False, fill='x')

    sheet_fr = ttk.Frame(b_fr)
    sheet_fr.pack(side='left', padx=8, pady=2, anchor='nw', expand=False, fill='x')
    sheet_fr.configure(height=10)
    reload_button = ttk.Button(sheet_fr, text="Get Worksheets", command=reload_worksheets)
    reload_button.grid(row=0, column=0, padx=5)
    select_all_button = ttk.Button(sheet_fr, text="Select All", command=select_all_cboxes)
    select_all_button.grid(row=0, column=1, padx=5)
    parbtn = ttk.Button(sheet_fr, text="Copy Worksheet", command=copy_worksheet)
    parbtn.grid(row=2, column=0, padx=4)
    ou_xml = ttk.Button(sheet_fr, text="Gen XML", command=lambda: process_data_file('in_ent'))
    ou_xml.grid(row=2, column=1, padx=4)

    sheet_cbox_fr = ttk.LabelFrame(b_fr, text='Select WorkSheet')
    sheet_cbox_fr.pack(side='top', padx=8, pady=2, anchor='nw', expand=True, fill='x')
    #sheet_cbox_fr.configure(height=5)
    worksheets = get_worksheets()
    create_checkboxes(worksheets)

    # Selection for subcategory
    f_fr = ttk.LabelFrame(page3, text="Sub-Category")
    f_fr.pack(side='top', padx=8, pady=2, anchor='nw', expand=False, fill='x')
    select_cbox = []
    cbox_names = ['Detail', 'Specification', 'Results']
    for i, cboxes in enumerate(cbox_names):
        select_cbox.append(tk.IntVar)
        check = ttk.Checkbutton(f_fr, text=cboxes, variable=select_cbox[i])
        check.grid(row=0, column=i)
    pf_lbl = ttk.Label(f_fr, text=" | Filter Patterns:")
    pf_ent = ttk.Entry(f_fr, width=100,)
    pf_lbl.grid(row=0, column=3, padx=10)
    pf_ent.grid(row=0, column=4, padx=2)
    #pf_btn = ttk.Button(f_fr, text="Show Table", command=show_filtered_table)
    #pf_btn.grid(row=1, column=4, padx=2)
    #pf_btn = ttk.Button(f_fr, text="Show Table", command=show_filtered_table)
    #pf_btn.grid(row=1, column=4, padx=2)
    ui_main.update_idletasks()

def show_filtered_table():
    pass

def nb2_create_pg1():
    global logs_text
    pg1 = ttk.Frame(notebook2)
    notebook2.add(pg1, text="Log")
    # Create a scrollable frame for logs and output
    scrollable_frame = ttk.Frame(pg1)
    scrollable_frame.pack(side='left', fill='both', expand=True, anchor='nw')

    logs_text = scrolledtext.ScrolledText(scrollable_frame, height=10, width=50)
    logs_text.pack(fill='both', expand=True, padx=8, pady=5)

    logs_scrollbar = ttk.Scrollbar(scrollable_frame, orient='vertical', command=logs_text.yview)
    logs_scrollbar.pack(side='right', fill='y')
    logs_text.configure(yscrollcommand=logs_scrollbar.set, height=10)

    clear_button = ttk.Button(scrollable_frame, text="Clear Logs", command=clear_logs)
    clear_button.pack(pady=10)
    ui_main.update_idletasks()


def nb2_create_pg2():
    pg2 = ttk.Frame(notebook2)
    notebook2.add(pg2, text="Tables")
    ui_main.update_idletasks()

# Create the main window
ui_main = tk.Tk()
ui_main.title('Demo CoFC Parser')

menu_bar = tk.Menu(ui_main)
ui_main.config(menu=menu_bar)

file_menu = tk.Menu(menu_bar,tearoff=0)
menu_bar.add_cascade(label="File",menu=file_menu)
file_menu.add_command(label="Exit", command=exit_program)

# -------------------------------------------------
##PANE SPLIT:
## Create the top pane (horizontal split)
#paned_window = ttk.PanedWindow(ui_main, orient='vertical')
#paned_window.pack(fill='both', expand=True)
## Create top pane
#top_pane = ttk.Frame(paned_window)
#paned_window.add(top_pane)
## Create the bottom pane
#bottom_pane = ttk.Frame(paned_window)
#paned_window.add(bottom_pane)

top_pane = ttk.Frame(ui_main)
top_pane.pack(side='top', fill='x', anchor='nw')
bottom_pane = ttk.Frame(ui_main)
bottom_pane.pack(side='top', fill='both', expand=True, anchor='nw')

# Creating Notebook Pages
notebook = ttk.Notebook(top_pane)
notebook.pack(fill='x')
nb_create_page1()
nb_create_page2()
nb_create_page3()

notebook2 = ttk.Notebook(bottom_pane)
notebook2.pack(fill='both', expand=True)
nb2_create_pg1()
nb2_create_pg2()
# -------------------------------------------------

ui_main.minsize(height=400, width=400)
ui_main.update()
# Start the Tkinter event loop
ui_main.mainloop()

