import tkinter as tk
import tkinter.filedialog
import os

from merge_files import merge_files_into_xlsx, merge_files_into_csv

def create_window():

    root = tk.Tk()
    root.title("Merge xlsx files")
    root.geometry("550x350")

    
    # Function to select a folder
    def select_folder( ):
        
        folder_path = tk.filedialog.askdirectory()
        if folder_path == "":
            return
        print("Selected Folder:", folder_path)
        
        in_folder_entry.delete(0, tk.END)
        in_folder_entry.insert(0, folder_path)
        
        # Clear the listbox before adding new items
        file_listbox.delete(0, tk.END)

        # Get all files in the selected folder
        files = [ file for file in os.listdir(folder_path) if file.endswith(".xlsx") ]
    
        # Add each file to the listbox
        for file in files:
            file_listbox.insert(tk.END, file)
        
        file_listbox.focus()
        file_listbox.selection_set(0, tk.END)
            
    def select_outfile():
        filetypes= [
            ("CSV files", "*.csv"),
            ("Excel files", "*.xlsx")
        ]
        
        out_path = tk.filedialog.asksaveasfile(filetypes=filetypes).name
        print("Selected File:", out_path)

        out_file_entry.delete(0, tk.END)
        out_file_entry.insert(0, out_path)

    def select_all():
        file_listbox.select_set(0, tk.END)

    def deselect_all():
        file_listbox.selection_clear(0, tk.END)
    
    def select_invert():
        selection = file_listbox.curselection()
        file_listbox.selection_set(0, tk.END) 
        for item in selection:
            file_listbox.selection_clear(item)

    # Function to move a file up in the list
    def move_file_up(event=None, to_edge=False):
        print("Move up")
        # Get the selected item(s)
        selected_indices = file_listbox.curselection()

        # Move each selected item up in the list
        for index in selected_indices:
            if index > 0:
                where = 0 if to_edge else index - 1
                file_listbox.insert(where, file_listbox.get(index))
                file_listbox.selection_set(where)
                file_listbox.delete(index + 1)

        return "break"  # prevent default bindings from firing
    
    def move_file_down(event=None, to_edge=False):
        # Get the selected item(s)
        print("Move down")

        selected_indices = file_listbox.curselection()

        # Move each selected item down in the list
        for index in reversed(selected_indices):
            if index < file_listbox.size() - 1:
                where = tk.END if to_edge else index + 2
                file_listbox.insert(where, file_listbox.get(index))
                file_listbox.selection_set(where)
                file_listbox.delete(index)

        return "break"  # prevent default bindings from firing
    
    def save(excel = False): 
        
        
        
        result_label.config(text = f"Saving..." )
        
        try: 
            files = [ os.path.join(in_folder_entry.get(), file) for file in file_listbox.get(0, tk.END) ]
            out = out_file_entry.get()
            if out == "":
                out = "output"
            lines = int(lines_entry.get())
            
            expected = ".csv" if not excel else ".xlsx"
            if(not out.endswith(expected)):
                out = out.replace( ".csv" if excel else ".xlsx", expected)
                if(not out.endswith(expected)):
                    out += expected
                out_file_entry.delete(0, tk.END)
                out_file_entry.insert(0, out)
            
            func = merge_files_into_xlsx if excel else merge_files_into_csv
            result = func(files, out, lines)        
            result_label.config(text = f"Successfully saved {result} lines" )
        except Exception as e:
            result_label.config(text = f"Failed: " + str(e) )
            
            
    
    save_xlsx = lambda: save(True)
    save_csv = lambda: save(False)

    top_panel = tk.Frame(root, padx=5, pady=5)
    top_panel.pack(side=tk.TOP, fill=tk.X, expand=True)
    top_panel.columnconfigure(1, weight=1)

    # Create a label and entry widget for a number
    lines_label = tk.Label(top_panel, text="Lines to skip", )
    lines_label.grid(column=0, row=0, padx=5, pady=5, sticky=tk.W)
    lines_entry = tk.Entry(top_panel)
    lines_entry.grid(column=1, row=0, padx=5, pady=5, sticky=tk.W)

    # Create buttons to select a folder and file
    in_folder_button = tk.Button(
        top_panel, text="Select Input Folder (ctrl+I)", command=select_folder)
    in_folder_button.grid(column=0, row=1, sticky=tk.EW)
    in_folder_entry = tk.Entry(top_panel)
    in_folder_entry.grid(column=1, row=1, padx=5, pady=5,  sticky=tk.EW)

    out_file_button = tk.Button(
        top_panel, text="Select Output location (ctrl+O)", command=select_outfile)
    out_file_button.grid(column=0, row=2, sticky=tk.EW)
    out_file_entry = tk.Entry(top_panel)
    out_file_entry.grid(column=1, row=2, padx=5, pady=5,  sticky=tk.EW)
    
    list_grid = tk.Frame(root, padx=5, pady=5)
    list_grid.pack(expand=True, fill=tk.BOTH)

    button_grid = tk.Frame(list_grid, padx=5, pady=5)
    button_grid.pack(side=tk.LEFT, anchor=tk.N)

    # Create a listbox to display the files
    file_listbox = tk.Listbox(list_grid, selectmode=tk.EXTENDED)
    file_listbox.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)

    # Create buttons to move files up and down in the list
    up_button = tk.Button(
        button_grid, text="Move Up (Ctrl+↑)", command=move_file_up)
    up_button.grid(column=0, row=0, sticky=tk.NSEW)
    down_button = tk.Button(
        button_grid, text="Move Down (Ctrl+↓)", command=move_file_down)
    down_button.grid(column=0, row=1, sticky=tk.NSEW)

    top_button = tk.Button(button_grid, text="Move To Top (Ctrl+Shift+↑)",
                           command=lambda: move_file_up(to_edge=True))
    top_button.grid(column=0, row=2, sticky=tk.NSEW)
    bottom_button = tk.Button(button_grid, text="Move To Bottom (Ctrl+Shift+↓)",
                              command=lambda: move_file_down(to_edge=True))
    bottom_button.grid(column=0, row=3, sticky=tk.NSEW)

    deselect_button = tk.Button(
        button_grid, text="Deselect", command=deselect_all)
    deselect_button.grid(column=0, row=4, sticky=tk.NSEW)
    select_all_button = tk.Button(
        button_grid, text="Select All", command=select_all)
    select_all_button.grid(column=0, row=5, sticky=tk.NSEW)
    select_invert_button = tk.Button(
        button_grid, text="Invert Selection", command=select_invert)
    select_invert_button.grid(column=0, row=6, sticky=tk.NSEW)

    convert_csv_button = tk.Button(root, text="To .csv (ctrl-K)", command=save_csv)
    convert_csv_button.pack(side=tk.LEFT, pady=5, padx=5)

    convert_xlsx_button = tk.Button(root, text="To .xlsx (ctrl-L)", command=save_xlsx)
    convert_xlsx_button.pack(side=tk.LEFT, pady=5, padx=5)
    
    result_label = tk.Label(root )
    result_label.pack(side=tk.LEFT, pady=5, padx=5)

    # hotkeys
    file_listbox.bind("<Control-Up>", move_file_up)
    file_listbox.bind("<Control-Down>", move_file_down)

    file_listbox.bind("<Control-Shift-Up>",
                      lambda event: move_file_up(event, to_edge=True))
    file_listbox.bind("<Control-Shift-Down>",
                      lambda event: move_file_down(event, to_edge=True))
    
    root.bind("<Control-k>", lambda event: save_csv)
    root.bind("<Control-l>",  lambda event: save_xlsx)
    root.bind("<Control-i>", lambda event: select_folder())
    root.bind("<Control-o>", lambda event: select_outfile())
    

    return root


def run_gui():
    root = create_window()
    root.mainloop()
