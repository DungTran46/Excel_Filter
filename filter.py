import csv
import sys
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

################################################
############### Define Constants ###############
################################################

HEADER_SALE_NUMER_RECORD            = "Sales Record Number" 
HEADER_SHIP_TO_NAME                 = "Ship To Name"
HEADER_ITEM_TITLE                   = "Item Title"
HEADER_QUANTITY                     = "Quantity"

HEADER_LIST = [
    HEADER_SALE_NUMER_RECORD,
    HEADER_SHIP_TO_NAME,
    HEADER_ITEM_TITLE,
    HEADER_QUANTITY
]

PICKING_LIST__HEADER_REF = [
    HEADER_QUANTITY,
    HEADER_ITEM_TITLE,
]

PACKING_LIST__HEADER_REF = [
    HEADER_SHIP_TO_NAME,
    HEADER_QUANTITY,
    HEADER_ITEM_TITLE
]

CSV_FILE_EXT                = ".csv"


#################### Global ####################
Input_fnames = []
################################################

def format_csv_file(csv_list):
    formated_list               = []
    header_row_len              = 0
    header_keywords_pos_dict    = {}

    print("******************Remove the Following Row*********************")
    for current_csv_row in csv_list:
        if (len(current_csv_row) != 0):
            #Find the header row and validate with the exepected HEADER_LIST
            if(len(header_keywords_pos_dict) != len(HEADER_LIST)):
                for col_idx, item in enumerate(current_csv_row):
                    # Sometime the header row could have extra charater such as newline or tab
                    # Remove them here to compare the string easier
                    current_csv_row[col_idx] = item.strip()
                    for header_list_idx, header_name in enumerate(HEADER_LIST):
                        #Found the matched header in the current row and save to bit map
                        if (current_csv_row[col_idx] == header_name):
                            header_keywords_pos_dict[HEADER_LIST[header_list_idx]] = col_idx

                if (len(header_keywords_pos_dict) == len(HEADER_LIST)):
                    header_row_len = len(current_csv_row)

                continue


            if (len(header_keywords_pos_dict) == len(HEADER_LIST) and len(current_csv_row) == header_row_len):
                for item in current_csv_row:
                    if(item != ""):
                        # Make sure we don't add any empty rows
                        formated_list.append(current_csv_row)
                        break
            elif(len(header_keywords_pos_dict) == len(HEADER_LIST)):
                # the row is illegal format
                print(current_csv_row)
    print("***************************************************************")
    return formated_list, header_keywords_pos_dict

def get_pick_list_item_name(item):
    return item[PICKING_LIST__HEADER_REF.index(HEADER_ITEM_TITLE)]

# Extract Items for picking list
def extract_picking_list(item_list, header_keywords_pos_dict):
    print("Extract picking list")
    picking_list = []

    #Construct the List
    for row in item_list:
        # If item title is empty, the buyer buys multiple item. Ignore that row
        if(row[header_keywords_pos_dict[HEADER_ITEM_TITLE]] != ''):
            saved_col = []
            #construct a list of row need to be saved.
            for saved_item in PICKING_LIST__HEADER_REF:
                saved_col.append(row[header_keywords_pos_dict[saved_item]].strip())
            picking_list.append(saved_col)

    #sort by item name
    picking_list.sort(key=get_pick_list_item_name)
    # We don't want the header to be sorted with data, so insert header row last
    picking_list.insert(0, PICKING_LIST__HEADER_REF)
    # Debug
    # for item in picking_list:
    #     print(item)
    return picking_list

def get_packing_list_name(item):
    return item[PACKING_LIST__HEADER_REF.index(HEADER_SHIP_TO_NAME)]

# Extract Items for packing list
def extract_packing_list(item_list, header_keywords_pos_dict):
    print("Extract packing list")
    packing_list = []
    previous_row = item_list[0]
    # -2 due to footer rows
    for row in item_list[0:len(item_list) - 2]:
        if(row[header_keywords_pos_dict[HEADER_SHIP_TO_NAME]] == ''):
            row[header_keywords_pos_dict[HEADER_SHIP_TO_NAME]] = previous_row[header_keywords_pos_dict[HEADER_SHIP_TO_NAME]]
        previous_row = row
        saved_col = []
        for saved_item in PACKING_LIST__HEADER_REF:
            # Ship to name need to be upper case to be able to sort 
            if(saved_item == HEADER_SHIP_TO_NAME):
                saved_col.append(row[header_keywords_pos_dict[saved_item]].strip().upper())
            else:
                saved_col.append(row[header_keywords_pos_dict[saved_item]].strip())
        packing_list.append(saved_col)
        #print([row[ship_to_name_idx], row[quantity_idx], row[item_title_idx]])

    #sort by packing list name
    packing_list.sort(key=get_packing_list_name)
    packing_list.insert(0, PACKING_LIST__HEADER_REF)
    # Debug
    # for item in packing_list:
    #     print(item)
    return packing_list
    

def parse_csv(in_file_name, picking_list_fname, packing_list_fname):
    clean_list                  = []
    picking_list                = []
    packing_list                = []
    header_keywords_pos_dict    = {}

    with open(in_file_name, newline='') as csv_in_file:
        reader = csv.reader(csv_in_file)

        # Find header row index and remove empty row and invalid row
        # Then put into a good list
        clean_list, header_keywords_pos_dict = format_csv_file(list(reader))

    if (len(header_keywords_pos_dict) != len(HEADER_LIST)):
        # Cannot find the matching header, report to user
        header_error_report(in_file_name, header_keywords_pos_dict)
    else:
        # Extract pick list
        picking_list = extract_picking_list(clean_list, header_keywords_pos_dict)
        with open(picking_list_fname, 'w', newline='') as write_csv:
            writer = csv.writer(write_csv)
            for row in picking_list:
                writer.writerow(row)

        # Extract packing list
        packing_list = extract_packing_list(clean_list, header_keywords_pos_dict)
        with open(packing_list_fname, 'w', newline='') as write_csv:
            writer = csv.writer(write_csv)
            for row in packing_list:
                writer.writerow(row)

def header_error_report(in_file_name, header_keywords_pos_dict):
        #construct the file messages
        out_message = []
        out_message.append('Opps. *_* Header names not found in')
        out_message.append(in_file_name + '\n')
        out_message.append('Sofware is looking for the following key words:')
        for item in HEADER_LIST:
            if (item in header_keywords_pos_dict):
                out_message.append(item + ' - ' + 'FOUND')
            else:
                out_message.append(item+ ' - ' + 'NOT FOUND')
        messagebox.showerror( 'Bad Excel File', "\n".join(out_message))
        fname_text_box.delete('1.0', END)
        fname_text_box.update()

def do_work():
    global Input_fnames
    for in_file_name in Input_fnames:
        #Construct output file name [in_file]_pick_list.csv
        print("Parsing File {0}".format(in_file_name))
        file_ext_pos = in_file_name.find(CSV_FILE_EXT,0,len(in_file_name))
        pick_list_file_name = in_file_name[0:file_ext_pos] + "_pick_list" + CSV_FILE_EXT
        packing_list_file_name = in_file_name[0:file_ext_pos] + "_packing_list" + CSV_FILE_EXT
        print('Output file name {0} {1}'.format(pick_list_file_name, packing_list_file_name))
        #Parse the output
        parse_csv(in_file_name, pick_list_file_name, packing_list_file_name)
        Input_fnames = []

def parsing_start_call_back():
    global Input_fnames
    print(Input_fnames)
    if (len(Input_fnames) <= 0):
        messagebox.showinfo( "ERROR", "Please Select File Name")
    else:
        do_work()
        messagebox.showinfo( "Info", "Done")
        fname_text_box.delete('1.0', END)
        fname_text_box.update()

# Open File call back
def open_file_callback():
    open_file()

# Open pop up File select menu
def open_file():
    global Input_fnames 
    global fname_text_box
    Input_fnames = list(filedialog.askopenfilenames(title="Openfile"))
    output = ""
    for file in Input_fnames:
        output = output + file + '\n'
    print(output)
    fname_text_box.insert(END, output)

if __name__ == "__main__":

    top = Tk()
    top.title('Ebay Excel Filter')
    top.geometry("900x500")

    # Create Button
    select_file_button = Button(top, text = "Select Files", command = open_file_callback)
    #select_file_button.grid(row = 0, column = 100)
    select_file_button.place(x = 450,y = 0)

    fname_label = Label(top, text="Selected Files Names:")
    #fname_label.grid(row = 1, column = 100)
    fname_label.place(x = 425, y = 50)

    fname_text_box = Text(top, height=20, width=100)
    fname_text_box.insert(END, Input_fnames)
    #fname_text_box.grid(row = 3, column = 100)
    fname_text_box.place(x = 50,y = 75)

    confirm_button = Button(top, text = "Confirm Button", command = parsing_start_call_back)
    #confirm_button.grid(row = 4, column = 100)
    confirm_button.place(x = 440,y = 400)


    top.mainloop()
