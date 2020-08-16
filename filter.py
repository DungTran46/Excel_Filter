import csv
import sys
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox

##################Define Constant###############
SAMPLE_FILE                 = "before.csv"
USER_ID_TXT                 = "Buyer Username"
SALE_NUMER_RECORD_TXT       = "Sales Record Number" 
CSV_FILE_EXT                = ".csv"
ITEM_TITLE_TXT              = "Item Title"
QUANTITY_TXT                = "Quantity"
SHIP_TO_NAME_TXT            = "Ship To Name"
PICKING_LIST_ITEM_NAME_POS  = 1
PACKING_LIST_BUY_NAME_POS   = 0
MAX_ARGS                    = 2
HEAD_DER_POS                = 0
################################################

#################### Global ####################
Input_fnames = []
################################################

def format_csv_file(csv_list):
    temp_list = []
    header_row_len = 0
    found_header_row = False
    print("******************Remove the Following Row*********************")
    for csv_row in csv_list:
        if (len(csv_row) != 0):
            # Remove any tab or new line in each cell
            for idx, item in enumerate(csv_row):
                csv_row[idx] = item.strip()

            if(SALE_NUMER_RECORD_TXT in csv_row):
                header_row_len = len(csv_row)
                found_header_row = True

            if (found_header_row and header_row_len == len(csv_row)):
                for row in csv_row:
                    if(row != ""):
                        temp_list.append(csv_row)
                        break
            else:
                print(csv_row)
    print("***************************************************************")
    return temp_list

def get_pick_list_item_name(item):
    return item[PICKING_LIST_ITEM_NAME_POS]

# Extract Items for picking list
def extract_pick_list(item_list):
    print("Extract picking list")
    pick_list = []
    #Find collumn index of the `Item Tile` and `Quantity`
    item_title_idx = item_list[HEAD_DER_POS].index(ITEM_TITLE_TXT, 0, len(item_list[HEAD_DER_POS]))
    quantity_idx = item_list[HEAD_DER_POS].index(QUANTITY_TXT, 0, len(item_list[HEAD_DER_POS]))
    print("Item Tile Index: {0}, Quantity Index {1}".format(item_title_idx, quantity_idx))

    #Construct the List
    for row in item_list[1:len(item_list)]:
        if(row[item_title_idx] != ''):
            pick_list.append([row[quantity_idx], row[item_title_idx]])
        #print("{0} {1}".format(row[item_title_idx],row[quantity_idx]))

    #sort by item name
    pick_list.sort(key=get_pick_list_item_name)
    pick_list.insert(0, [QUANTITY_TXT,ITEM_TITLE_TXT])
    return pick_list

def get_packing_list_name(item):
    return item[PACKING_LIST_BUY_NAME_POS]

# Extract Items for packing list
def extract_packing_list(item_list):
    print("Extract packing list")
    packing_list = []
    #Find collumn index of the `Ship To Name`, `Item Tile` and `Quantity`
    item_title_idx      = item_list[HEAD_DER_POS].index(ITEM_TITLE_TXT, 0, len(item_list[HEAD_DER_POS]))
    quantity_idx        = item_list[HEAD_DER_POS].index(QUANTITY_TXT, 0, len(item_list[HEAD_DER_POS]))
    ship_to_name_idx    = item_list[HEAD_DER_POS].index(SHIP_TO_NAME_TXT, 0, len(item_list[HEAD_DER_POS]))
    sale_number_idx     = item_list[HEAD_DER_POS].index(SALE_NUMER_RECORD_TXT, 0, len(item_list[HEAD_DER_POS]))
    user_id_idx         = item_list[HEAD_DER_POS].index(USER_ID_TXT, 0, len(item_list[HEAD_DER_POS]))
    print("item_title_idx: {0}, quantity_idx {1}, ship_to_name_idx {2}, sale_number_idx{3}, user_id_idx{4}"
            .format(item_title_idx, quantity_idx, ship_to_name_idx, sale_number_idx, user_id_idx))
    previous_row = item_list[1]

    # -2 due to footer rows
    for row in item_list[1:len(item_list) - 2]:
        if(row[ship_to_name_idx] == ''):
            row[ship_to_name_idx] = previous_row[ship_to_name_idx]
        previous_row = row
        packing_list.append([row[ship_to_name_idx].lower(), row[quantity_idx], row[item_title_idx]])
        #print([row[ship_to_name_idx], row[quantity_idx], row[item_title_idx]])

    #sort by packing list name
    packing_list.sort(key=get_packing_list_name)
    packing_list.insert(0, [SHIP_TO_NAME_TXT, QUANTITY_TXT, ITEM_TITLE_TXT])
    return packing_list
    

def parse_csv(in_file_name, pick_list_fname, packing_list_fname):
    temp_list = []
    pick_list = []
    packing_list = []
    with open(in_file_name, newline='') as csv_in_file:
        reader = csv.reader(csv_in_file)

        #convert reader to a list for row random access
        csv_list = list(reader)
        # Find header row index and remove empty row and invalid row
        # Then put into a good list
        temp_list = format_csv_file(csv_list)


        if len(temp_list) <= 1:
            raise Exception("Bad Excel files")

        #Santity check the temp list again
        for row in temp_list:
            if(len(row) != len(temp_list[0])):
                print(len(row))
                print(row)
                raise Exception("Unexpected Row size")
    
    #Extract pick list
    pick_list = extract_pick_list(temp_list)
    with open(pick_list_fname, 'w', newline='') as write_csv:
        writer = csv.writer(write_csv)
        for row in pick_list:
            writer.writerow(row)

    #Extract packing list
    packing_list = extract_packing_list(temp_list)
    with open(packing_list_fname, 'w', newline='') as write_csv:
        writer = csv.writer(write_csv)
        for row in packing_list:
            writer.writerow(row)

def parsing_call_back():
    global Input_fnames
    print(Input_fnames)
    if (len(Input_fnames) <= 0):
        messagebox.showinfo( "ERROR", "Please Select File Name")
    else:
        do_work()
        messagebox.showinfo( "Info", "Done")
        fname_text_box.delete('1.0', END)
        fname_text_box.update()


def do_work():
    global Input_fnames
    for in_file_name in Input_fnames:
        #Construct output file name [in_file]_pick_list.csv
        file_ext_pos = in_file_name.find(CSV_FILE_EXT,0,len(in_file_name))
        pick_list_file_name = in_file_name[0:file_ext_pos] + "_pick_list" + CSV_FILE_EXT
        packing_list_file_name = in_file_name[0:file_ext_pos] + "_packing_list" + CSV_FILE_EXT
        print('Output file name {0} {1}'.format(pick_list_file_name, packing_list_file_name))
        #Parse the output
        parse_csv(in_file_name, pick_list_file_name, packing_list_file_name)
        Input_fnames = []
        

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
    top.geometry("500x800")

    # Create Button
    select_file_button = Button(top, text = "Select Files", command = open_file)
    select_file_button.grid(row=0, column=0)

    #select_file_button.place(x = 50,y = 50)
    fname_label = Label(top, text="Selected Files Names")
    fname_label.grid(row=1, column=0)

    fname_text_box = Text(top, height=20, width=50)
    fname_text_box.grid(row = 3, column = 0)
    fname_text_box.insert(END, Input_fnames)

    confirm_button = Button(top, text = "Confirm Button", command = parsing_call_back)
    confirm_button.grid(row = 4, column = 0)


    top.mainloop()
