import logging
import tkinter as tk
from tkinter import filedialog
import os
from logs import CustomHandler


from tasks import run_task
from merge import Merger

# This code is ugly, but it is not the point of this project

def guess_output_name(input_filenames):
    for name in input_filenames:
        if "20" in name:
            return "sum_" + name.split("20")[0].strip()

def create_window():
    
    logger = logging.getLogger("console_logger")
    logger.setLevel(logging.DEBUG)
    

    root = tk.Tk()
    root.title("Merge xlsx files")
    root.geometry("800x600")

    # Function to select a folder
    def select_folder():

        folder_path = filedialog.askdirectory()
        if folder_path == "":
            return
        print("Selected Folder:", folder_path)
        
        in_folder_entry.delete(0, tk.END)
        in_folder_entry.insert(0, folder_path)
        

        # Clear the listbox before adding new items
        file_listbox.delete(0, tk.END)

        # Get all files in the selected folder
        files = [file for file in os.listdir(
            folder_path) if file.endswith(".xlsx")]

        # Add each file to the listbox
        for file in files:
            file_listbox.insert(tk.END, file)

        file_listbox.focus()
        file_listbox.selection_set(0, tk.END)
        
        if out_file_entry.get() == "":
            possible_output_name = guess_output_name(files)
            possible_output_dir = os.path.join(folder_path, possible_output_name)
            
            out_file_entry.delete(0, tk.END)
            out_file_entry.insert(0, possible_output_dir)
            
        root.after(1, lambda: root.focus_force())
        

    def select_outfile():
        filetypes = [
            ("CSV files", ".csv"),
            ("Excel files", ".xlsx")
        ]

        out = filedialog.asksaveasfilename(
            filetypes=filetypes,  defaultextension=".csv", initialfile="output.csv")
        out_path = out
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
    def move_file_up(event=None):
        selected_indices = file_listbox.curselection()
        if 0 in selected_indices:
            return "break"
        for index in selected_indices:
            if index > 0:
                file_listbox.insert(index - 1, file_listbox.get(index))
                file_listbox.selection_set(index - 1)
                file_listbox.delete(index + 1)

        return "break"  # prevent default bindings from firing

    def move_file_to_top(event=None):
        indices = file_listbox.curselection()
        files = [file_listbox.get(index) for index in indices]
        for index in reversed(indices):
            file_listbox.delete(index)
        for file in reversed(files):
            file_listbox.insert(0, file)
            file_listbox.selection_set(0)

        return "break"  # prevent default bindings from firing

    def move_file_down(event=None):
        selected_indices = file_listbox.curselection()
        if file_listbox.size() - 1 in selected_indices:
            return "break"
        for index in reversed(selected_indices):
            if index < file_listbox.size() - 1:
                file_listbox.insert(index + 2, file_listbox.get(index))
                file_listbox.selection_set(index + 2)
                file_listbox.delete(index)

        return "break"  # prevent default bindings from firing

    def move_file_to_bottom(event=None):
        indices = file_listbox.curselection()
        
        files = [file_listbox.get(index) for index in indices]
        for index in reversed(indices):
            file_listbox.delete(index)

        for file in files:
            file_listbox.insert(tk.END, file)
            file_listbox.selection_set(tk.END)

        return "break"  # prevent default bindings from firing

    def remove(event=None):
        indices = file_listbox.curselection()
        for index in reversed(indices):
            file_listbox.delete(index)


    def save(excel=False):

        logger.info(f"Started task")

        try:
            files = [os.path.join(in_folder_entry.get(), file_listbox.get(
                file)) for file in file_listbox.curselection()]
            out = out_file_entry.get()
            if out == "":
                out = "output"
            lines = int(lines_entry.get()) if lines_entry.get() != "" else 0

            expected = ".csv" if not excel else ".xlsx"
            if (not out.endswith(expected)):
                out = out.replace(".csv" if excel else ".xlsx", expected)
                if (not out.endswith(expected)):
                    out += expected
                out_file_entry.delete(0, tk.END)
                out_file_entry.insert(0, out)
                
            save_extra_cells = list(map(lambda x : x.strip(), extra_cells_entry.get().split(",")))
            save_extra_cells = list(filter(lambda x : x != "", save_extra_cells))
            merger = Merger(logger, save_extra_cells)

            func = merger.merge_files_into_xlsx if excel else merger.merge_files_into_csv
            logger.info(f"Saving {len(files)} files to {out} with {lines} lines skipped and following extra cells: {save_extra_cells}")
            result = func(files, out, lines)
            logger.info(f"Successfully saved {result} lines into {out}")
        except Exception as e:
            logger.info(f"Failed saving to {out}: " + str(e))

    def save_task(excel=False):
        run_task(lambda: save(excel))

    def save_xlsx(): return save_task(True)
    def save_csv(): return save_task(False)

    top_panel = tk.Frame(root, padx=5, pady=5)
    top_panel.pack(side=tk.TOP, fill=tk.X)
    top_panel.columnconfigure(1, weight=1)
    
    row = 0

    # Create a label and entry widget for a number
    lines_label = tk.Label(top_panel, text="Lines to skip")
    lines_label.grid(column=0, row=row, padx=5, pady=5, sticky=tk.W)
    lines_entry = tk.Entry(top_panel)
    lines_entry.insert(0, "10")
    lines_entry.grid(column=1, row=row, padx=5, pady=5, sticky=tk.W)
    
    
    # Create a label and entry widget for extra celss
    extra_cells_label = tk.Label(top_panel, text="Extra cells to save (comma separated)")
    extra_cells_label.grid(column=2, row=row, padx=5, pady=5, sticky=tk.W)
    extra_cells_entry = tk.Entry(top_panel)
    extra_cells_entry.insert(0, "K6")
    extra_cells_entry.grid(column=3, row=row, padx=5, pady=5, sticky=tk.W)
    
    row += 1

    # Create buttons to select a folder and file
    in_folder_button = tk.Button(
        top_panel, text="Select Input Folder (ctrl+I)", command=select_folder)
    in_folder_button.grid(column=0, row=row, sticky=tk.EW)
    in_folder_entry = tk.Entry(top_panel)
    in_folder_entry.grid(column=1, columnspan=3, row=row, padx=5, pady=5,  sticky=tk.EW)
    
    row += 1

    out_file_button = tk.Button(
        top_panel, text="Select Output location (ctrl+O)", command=select_outfile)
    out_file_button.grid(column=0, row=row, sticky=tk.EW)
    out_file_entry = tk.Entry(top_panel)
    out_file_entry.grid(column=1, columnspan=3, row=row, padx=5, pady=5,  sticky=tk.EW)

    list_grid = tk.Frame(root, padx=5, pady=5)
    list_grid.pack(expand=True, fill=tk.BOTH)

    list_grid.rowconfigure(1, weight=1)

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
                           command= move_file_to_top)
    top_button.grid(column=0, row=2, sticky=tk.NSEW)
    bottom_button = tk.Button(button_grid, text="Move To Bottom (Ctrl+Shift+↓)",
                              command=move_file_to_bottom)
    bottom_button.grid(column=0, row=3, sticky=tk.NSEW)

    deselect_button = tk.Button(
        button_grid, text="Deselect All", command=deselect_all)
    deselect_button.grid(column=0, row=4, sticky=tk.NSEW)
    select_all_button = tk.Button(
        button_grid, text="Select All (ctrl A)", command=select_all)
    select_all_button.grid(column=0, row=5, sticky=tk.NSEW)
    select_invert_button = tk.Button(
        button_grid, text="Invert Selection", command=select_invert)
    select_invert_button.grid(column=0, row=6, sticky=tk.NSEW)
    delete_button = tk.Button(
        button_grid, text="Discard selected (delete)", command=remove)
    delete_button.grid(column=0, row=7, sticky=tk.NSEW)
    
    
    button_pane = tk.Frame(root)
    button_pane.pack(side=tk.TOP, anchor=tk.N, fill=tk.BOTH)
    

    convert_csv_button = tk.Button(
        button_pane, text="To .csv (ctrl-K)", command=save_csv)
    convert_csv_button.pack(side=tk.LEFT, pady=5, padx=5)

    convert_xlsx_button = tk.Button(
        button_pane, text="To .xlsx (ctrl-L)", command=save_xlsx)
    convert_xlsx_button.pack(side=tk.LEFT, pady=5, padx=5)
        
    console = tk.Text(root, pady=5, padx=5)
    console.configure(font=("Helvetica", 8))
    
    console.pack(side=tk.BOTTOM, anchor=tk.W, fill=tk.BOTH, pady=5, padx=5)
    
    
    def _log(line):
        scroll = console.yview()[1] == 1.0
        console.insert(tk.END, line + "\n")
        if scroll: 
            console.yview_pickplace("end")
    
    handler = CustomHandler(_log )
    logger.addHandler(handler)

    # hotkeys
    file_listbox.bind("<Control-Up>", move_file_up)
    file_listbox.bind("<Control-Down>", move_file_down)

    file_listbox.bind("<Control-Shift-Up>", move_file_to_top)
    file_listbox.bind("<Control-Shift-Down>", move_file_to_bottom)

    root.bind("<Control-k>", lambda event: save_csv())
    root.bind("<Control-l>", lambda event: save_xlsx())
    root.bind("<Control-i>", lambda event: select_folder())
    root.bind("<Control-o>", lambda event: select_outfile())
    root.bind("<Delete>", lambda event: remove())

    return root


def run_gui():
    root = create_window()
    root.mainloop()