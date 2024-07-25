from pandastable.headers import ColumnHeader, RowHeader, IndexHeader
from pandastable import Table, TableModel
from tkinter import filedialog, messagebox
from tkinter import ttk
from CTkXYFrame import *

import customtkinter as ctk
import json
import pandas as pd
import threading
import logic
import random
import copy
import inspect
import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)


class Routine_Generator(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("SXC Routine Management System")
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - 1280) // 2
        y = (screen_height - 720) // 2
        self.geometry(f"{1280}x{720}+{x}+{y}")
        self.minsize(1280, 720)

        self.tab_view = ctk.CTkTabview(self)
        self.tab_view.pack(fill='both', expand=True)
        self.tab_view._segmented_button.configure(font=ctk.CTkFont(size=25, weight="bold"))

        self.tab_view.add("Departments")
        self.tab_view.add("Rooms")
        self.tab_view.add("Routine")
        self.tab_view.add("About")

        self.departments_tab = self.tab_view.tab("Departments")
        self.rooms_tab = self.tab_view.tab("Rooms")
        self.routine_tab = self.tab_view.tab("Routine")
        self.about_tab = self.tab_view.tab("About")

        # Departments tab
        self.controls_frame_load_depts = ctk.CTkFrame(self.departments_tab, fg_color="transparent")
        self.controls_frame_load_depts.pack(pady=10, padx=10)

        self.controls_frame_add_remove_depts = ctk.CTkFrame(self.departments_tab, fg_color="transparent")
        self.controls_frame_add_remove_depts.pack(pady=10, padx=10)

        self.main_frame_depts = CTkXYFrame(self.departments_tab)
        self.main_frame_depts.pack(pady=10, padx=10, fill="both", expand=True)
        
        self.load_button_depts = ctk.CTkButton(self.controls_frame_load_depts, text="Load JSON", command=self.load_depts_json)
        self.load_button_depts.pack(side="left", padx=(0, 10))

        self.save_button_depts = ctk.CTkButton(self.controls_frame_load_depts, text="Save JSON", command=self.save_depts_json)
        self.save_button_depts.pack(side="left", padx=(10, 0))
        
        self.add_depts_button = ctk.CTkButton(self.controls_frame_add_remove_depts, text="Add Department", command=self.add_department)
        self.add_depts_button.pack(side="left", padx=(0, 10))
        
        self.remove_depts_button = ctk.CTkButton(self.controls_frame_add_remove_depts, text="Remove Department", command=self.remove_department)
        self.remove_depts_button.pack(side="left", padx=(10, 0))

        self.departments_frame = ctk.CTkFrame(self.main_frame_depts)
        self.departments = []

        #Rooms tab
        self.controls_frame_load_rooms = ctk.CTkFrame(self.rooms_tab, fg_color="transparent")
        self.controls_frame_load_rooms.pack(pady=10, padx=10)

        self.controls_frame_add_remove_rooms = ctk.CTkFrame(self.rooms_tab, fg_color="transparent")
        self.controls_frame_add_remove_rooms.pack(pady=10, padx=10)

        self.main_frame_rooms = ctk.CTkFrame(self.rooms_tab)
        self.main_frame_rooms.pack(pady=10, padx=10, fill="y", expand=True)
        
        self.load_button_rooms = ctk.CTkButton(self.controls_frame_load_rooms, text="Load JSON", command=self.load_rooms_json)
        self.load_button_rooms.pack(side="left", padx=(0, 10))

        self.save_button_rooms = ctk.CTkButton(self.controls_frame_load_rooms, text="Save JSON", command=self.save_rooms_json)
        self.save_button_rooms.pack(side="left", padx=(10, 0))
        
        self.add_rooms_button = ctk.CTkButton(self.controls_frame_add_remove_rooms, text="Add Room", command=self.add_room)
        self.add_rooms_button.pack(side="left", padx=(0, 10))
        
        self.remove_rooms_button = ctk.CTkButton(self.controls_frame_add_remove_rooms, text="Remove Room", command=self.remove_room)
        self.remove_rooms_button.pack(side="left", padx=(10, 0))

        headers = ["Room ID", "Capacity", "AC", "AV", "Floor"]
        data = []
        self.room_table = EditableTreeview(self.main_frame_rooms, headers, data)
        scrollbar = ttk.Scrollbar(self.main_frame_rooms, orient="vertical", command=self.room_table.yview)
        self.room_table.configure(yscroll=scrollbar.set)
        self.room_table.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        self.rooms_tab.bind("<Button-1>", self.on_click_outside)
        
        # Routine tab
        self.controls_frame_load = ctk.CTkFrame(self.routine_tab, fg_color="transparent")
        self.controls_frame_load.pack(pady=10, padx=10)

        self.controls_frame_routine = ctk.CTkFrame(self.routine_tab, fg_color="transparent")
        self.controls_frame_routine.pack(pady=10, padx=10)

        self.controls_frame_export_routine = ctk.CTkFrame(self.routine_tab, fg_color="transparent")

        self.main_frame_routine = ctk.CTkFrame(self.routine_tab)
        self.main_frame_routine.pack(pady=10, padx=100, fill="both", expand=True)

        self.load_button_depts = ctk.CTkButton(self.controls_frame_load, text="Load Departments JSON", command=self.load_depts_json_file)
        self.load_button_depts.pack(side="left", padx=10)

        self.load_depts_filename = ctk.CTkLabel(self.controls_frame_load, text="")
        self.load_depts_filename.pack(side="left", padx=(0, 30))

        self.load_button_rooms = ctk.CTkButton(self.controls_frame_load, text="Load Rooms JSON", command=self.load_rooms_json_file)
        self.load_button_rooms.pack(side="left", padx=5)

        self.load_rooms_filename = ctk.CTkLabel(self.controls_frame_load, text="")
        self.load_rooms_filename.pack(side="left", padx=5)
        
        self.generate_routine_button = ctk.CTkButton(self.controls_frame_routine, text="Generate Routine", command=self.generate_routine)
        self.generate_routine_button.pack(side="left", padx=10)
        self.generate_routine_button.configure(state="disabled")

        self.export_routine_button = ctk.CTkButton(self.controls_frame_export_routine, text="Export Routine", command=self.export_routine)
        self.export_routine_button.pack(side="left", padx=10)

        self.seed = None
        self.loading_frame = None  # Placeholder for the loading frame
        self.is_running = False

        # About tab
        self.create_about_tab()


    def add_department(self, dept_data=None):
        self.departments_frame.pack(padx=10, pady=10, fill="both", expand=True)
        dept_frame = DepartmentFrame(self.departments_frame, dept_data)
        dept_frame.pack(padx=20, pady=20, fill="both", expand=True)
        self.departments.append(dept_frame)
        
    def remove_department(self):
        if self.departments:
            dept_frame = self.departments.pop()
            dept_frame.destroy()
            if len(self.departments) == 0:
                self.departments_frame.pack_forget()


    def validate_depts(self):
        if "depts" not in self.depts_data:
            return False
        for dept in self.depts_data["depts"]:
            if not isinstance(dept, dict):
                return False
            if "name" not in dept or "type" not in dept or "homes" not in dept or "prgms" not in dept:
                return False
            if dept["type"] not in ("arts", "science"):
                return False
            for home in dept["homes"]:
                if "room_id" not in home:
                    return False
            for prgm in dept["prgms"]:
                if "type" not in prgm or "sems" not in prgm:
                    return False
                if prgm["type"] not in ("U.G.", "P.G."):
                    return False
                for sem in prgm["sems"]:
                    if "no" not in sem or "hasoffday" not in sem or "strength" not in sem:
                        return False
                    if not isinstance(sem["no"], int) or not isinstance(sem["strength"], int):
                        return False
                    if not isinstance(sem["hasoffday"], bool):
                        return False
                    if "ths" not in sem or "prs" not in sem:
                        return False
                    ths = sem["ths"]
                    prs = sem["prs"]
                    if not all(k in ths for k in ("major", "minor", "mds", "envs")):
                        return False
                    if not isinstance(ths["major"], int) or not isinstance(ths["minor"], int) or not isinstance(ths["mds"], int) or not isinstance(ths["envs"], int):
                        return False
                    if not all(k in prs for k in ("major", "minor", "mds")):
                        return False
                    for key in ("major", "minor", "mds"):
                        for item in prs[key]:
                            if not all(subkey in item for subkey in ("cons", "freq")):
                                return False
                            if not isinstance(item["cons"], int) or not isinstance(item["freq"], int):
                                return False
        return True

    def validate_rooms(self):
        if "rooms" not in self.rooms_data:
            return False
        for room in self.rooms_data["rooms"]:
            if not isinstance(room, dict):
                return False
            if "room_id" not in room or "capacity" not in room or "hasAC" not in room or "hasAV" not in room or "floor" not in room:
                return False
            if not isinstance(room["capacity"], int) or not isinstance(room["floor"], int):
                return False
            if not isinstance(room["hasAC"], bool) or not isinstance(room["hasAV"], bool):
                return False
        return True

    def load_depts_json(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if file_path:
            if not self.is_running:
                self.is_running = True
                threading.Thread(target=self.load_depts_data, args=(file_path,)).start()
            else:
                messagebox.showerror("Error", "The previous operation is still running.")

    def load_depts_data(self, file_path):
        with open(file_path, "r") as f:
            self.depts_data = json.load(f)
            if not self.validate_depts():
                messagebox.showerror("Failed", "Invalid Departments JSON format")
                if hasattr(self, "depts_data"):
                    delattr(self, "depts_data")
                self.is_running = False
                return

        num_depts = len(self.depts_data.get("depts", []))
        if not self.loading_frame:
            self.loading_frame = LoadingFrame(self)
            self.loading_frame.pack(side="bottom", pady=20, fill="x")
        self.loading_frame.set_progress_max(num_depts)
            
        self.departments = []
        for i, dept_data in enumerate(self.depts_data.get("depts", [])):
            self.add_department(dept_data)
            self.loading_frame.update_progress(i + 1)
            self.update()
        self.loading_frame.destroy()
        self.loading_frame = None
        messagebox.showinfo("Success", "Departments data loaded successfully!")
        self.is_running = False
       
        
    def save_depts_json(self):
        depts_json_data = {"depts": [dept.get_data() for dept in self.departments]}

        json_data = json.dumps(depts_json_data, indent=2)
        
        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if file_path:
            with open(file_path, "w") as f:
                f.write(json_data)
            messagebox.showinfo("Success", "Departments data saved successfully!")


    def add_room(self, room_data=["New Room", 0, "No", "No", 0]):
        new_room = room_data
        selected_item = self.room_table.selection()
        
        if selected_item:
            index = self.room_table.index(selected_item[0])
            new_item = self.room_table.insert('', index + 1, values=new_room)
        else:
            new_item = self.room_table.insert('', "end", values=new_room)
        
        self.room_table.selection_set(new_item)
    
    def remove_room(self):
        selected_item = self.room_table.selection()
        
        if selected_item:
            next_item = self.room_table.next(selected_item[0])
            if not next_item:
                next_item = self.room_table.prev(selected_item[0])
            self.room_table.delete(selected_item)
            if next_item:
                self.room_table.selection_set(next_item)
        else:
            last_item = self.room_table.get_children()
            if last_item:
                self.room_table.delete(last_item[-1])
    
    def on_click_outside(self, event):
        self.room_table.selection_remove(self.room_table.selection())
    
    def load_rooms_json(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if file_path:
            if not self.is_running:
                self.is_running = True
                threading.Thread(target=self.load_rooms_data, args=(file_path,)).start()
            else:
                messagebox.showerror("Error", "The previous operation is still running.")            

    def load_rooms_data(self, file_path):
        with open(file_path, "r") as f:
            self.rooms_data = json.load(f)
            if not self.validate_rooms():
                messagebox.showerror("Failed", "Invalid Rooms JSON format")
                if hasattr(self, "rooms_data"):
                    delattr(self, "rooms_data")
                self.is_running = False
                return

        table_data = [
            [room["room_id"], room["capacity"], room["hasAC"], room["hasAV"], room["floor"]]
            for room in self.rooms_data['rooms']
        ]
        num_rooms = len(table_data)
        if not self.loading_frame:
            self.loading_frame = LoadingFrame(self)
            self.loading_frame.pack(side="bottom", pady=20, fill="x")
        self.loading_frame.set_progress_max(num_rooms)
        
        self.rooms = []
        for i, room_data in enumerate(table_data):
            room_data[2] = 'Yes' if room_data[2] else 'No'
            room_data[3] = 'Yes' if room_data[3] else 'No'
            self.add_room(room_data)
            self.loading_frame.update_progress(i + 1)
            self.update()
        self.loading_frame.destroy()
        self.loading_frame = None
        messagebox.showinfo("Success", "Rooms data loaded successfully!")
        self.is_running = False


    def save_rooms_json(self):
        rooms_data = self.room_table.get_data()
        rooms_json_data = {
            "rooms": [
                {
                    "room_id": str(row[0]),
                    "capacity": row[1],
                    "hasAC": row[2] == "Yes",
                    "hasAV": row[3] == "Yes",
                    "floor": row[4]
                }
                for row in rooms_data
            ]
        }

        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if file_path:
            with open(file_path, "w") as f:
                json.dump(rooms_json_data, f, indent=2)
            messagebox.showinfo("Success", "Rooms data saved successfully!")


    def load_depts_json_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if file_path:
            file_name = file_path.split("/")[-1]
            with open(file_path, "r") as f:
                try:
                    self.depts_data = json.load(f)
                    if not self.validate_depts():
                        messagebox.showerror("Failed", "Invalid Departments JSON format")
                        if hasattr(self, "depts_data"):
                            delattr(self, "depts_data")
                        self.load_depts_filename.configure(text="")
                        self.check_for_data()
                        return
                except json.JSONDecodeError:
                    messagebox.showerror("Error", "Error decoding JSON file.")
                
            self.load_depts_filename.configure(text=file_name)
        else:
            self.load_depts_filename.configure(text="")

            if hasattr(self, "depts_data"):
                delattr(self, "depts_data")

        self.check_for_data()

    def load_rooms_json_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if file_path:
            file_name = file_path.split("/")[-1]
            with open(file_path, "r") as f:
                try:
                    self.rooms_data = json.load(f)
                    if not self.validate_rooms():
                        messagebox.showerror("Failed", "Invalid Rooms JSON format")
                        if hasattr(self, "rooms_data"):
                            delattr(self, "rooms_data")
                        self.check_for_data()
                        self.load_rooms_filename.configure(text="")
                        return
                except json.JSONDecodeError:
                    messagebox.showerror("Error", "Error decoding JSON file.")
                
            self.load_rooms_filename.configure(text=file_name)
        else:
            self.load_rooms_filename.configure(text="")

            if hasattr(self, "rooms_data"):
                delattr(self, "rooms_data")

        self.check_for_data()

    def check_for_data(self):
        if hasattr(self, "depts_data") and hasattr(self, "rooms_data"):
            self.generate_routine_button.configure(state="normal")
        else:
            self.generate_routine_button.configure(state="disabled")
            if hasattr(self, "routine"):
                self.main_frame_routine.pack_forget()
                self.main_frame_routine = ctk.CTkFrame(self.routine_tab)
                self.main_frame_routine.pack(pady=10, padx=100, fill="both", expand=True)
            self.controls_frame_export_routine.pack_forget()


    def randomize(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("Generate Routine")
        screen_width = dialog.winfo_screenwidth()
        screen_height = dialog.winfo_screenheight()
        x = (screen_width - 380) // 2
        y = (screen_height - 140) // 2
        dialog.geometry(f"{380}x{140}+{x}+{y}")
        dialog.resizable(False, False)
        generate_label = ctk.CTkLabel(dialog, text="Generate Routine w.r.t.:")
        generate_label.grid(row=0, column=0, padx=8, pady=10)
        routine_wrt_field = ctk.CTkOptionMenu(dialog, width=215, values=["Semesters", "Departments", "Rooms"])
        routine_wrt_field.grid(row=0, column=1, columnspan=2, pady=10)
        seed_label = ctk.CTkLabel(dialog, text="Enter Seed value:")
        seed_label.grid(row=1, column=0, padx=8, pady=10)
        seed_entry = ctk.CTkEntry(dialog, width=215, validate="key", validatecommand=(self.register(self.validate_numeric), '%P'), textvariable=IntegerVar(value=self.seed))
        seed_entry.grid(row=1, column=1, columnspan=2, pady=10)
        randomize_button = ctk.CTkButton(dialog, text="Randomize", width=100, command=lambda: self.random_seed(seed_entry))
        randomize_button.grid(row=2, column=0, padx=8, pady=5)
        ok_button = ctk.CTkButton(dialog, text="Ok", width=100, command=lambda: self.on_generate_clicked(seed_entry.get(), routine_wrt_field.get(), dialog))
        ok_button.grid(row=2, column=1, padx=(0, 15), pady=5)
        cancel_button = ctk.CTkButton(dialog, text="Cancel", width=100, command=dialog.destroy)
        cancel_button.grid(row=2, column=2, pady=5)
        dialog.grab_set()
        dialog.focus_force()
        dialog.bind('<Return>', lambda event: self.on_generate_clicked(seed_entry.get(), routine_wrt_field.get(), dialog))

    def random_seed(self, seed_entry):
        seed = random.randint(0, 10000)
        seed_entry.delete(0, "end")
        seed_entry.insert(0, seed)

    def validate_numeric(self, seed):
        if seed == "" or seed.isdigit():
            return True
        else:
            return False
        

    def on_generate_clicked(self, seed, view_field, dialog):
        if seed is not None and len(seed) > 0 and seed.isdigit():
            self.seed = seed
            dialog.destroy()
            if hasattr(self, "routine_tables_frame") and self.routine_tables_frame:
                self.main_frame_routine.pack_forget()
                self.main_frame_routine = ctk.CTkFrame(self.routine_tab)
                self.main_frame_routine.pack(pady=10, padx=100, fill="both", expand=True)
            self.controls_frame_export_routine.pack_forget()

            if not self.is_running:
                self.is_running = True
                self.routine_df = logic.generate(copy.deepcopy(self.depts_data), copy.deepcopy(self.rooms_data), int(seed))

                scrollable_frame = ctk.CTkScrollableFrame(self.main_frame_routine, orientation="horizontal", height=50)
                scrollable_frame.pack(fill="x")
                self.selected_routine = ctk.StringVar()


                self.loading_frame = LoadingFrame(self)
                self.loading_frame.pack(side="bottom", pady=20, fill="x")

                if view_field == "Semesters":
                    def classify_departments(departments):
                        arts_depts = [dept['name'] for dept in departments['depts'] if dept['type'] == 'arts']
                        science_depts = [dept['name'] for dept in departments['depts'] if dept['type'] == 'science']
                        return arts_depts, science_depts

                    arts_depts, science_depts = classify_departments(self.depts_data)

                    # Fetch all unique semesters dynamically from the JSON file
                    def fetch_semesters(departments):
                        semesters = set()
                        for dept in departments['depts']:
                            for prgm in dept['prgms']:
                                for sem in prgm['sems']:
                                    semesters.add(f"{prgm['type']} Sem {sem['no']}")
                        return sorted(semesters)

                    semesters = fetch_semesters(self.depts_data)

                    self.loading_frame.set_progress_max(len(semesters) * 2)

                    self.routine_tables_frame = {}

                    default_routine = None

                    for stream, depts in [("Arts", arts_depts), ("Science", science_depts)]:
                        for sem in semesters:
                            routine_name = f"{stream} {sem}"
                            if default_routine is None:
                                default_routine = routine_name

                            radio_button = ctk.CTkRadioButton(scrollable_frame, text=routine_name,
                                                            variable=self.selected_routine,
                                                            value=routine_name,
                                                            command=self.display_selected_routine)
                            radio_button.pack(side="left", padx=20)

                            columns_to_include = ['Day', 'Time'] + [col for col in self.routine_df.columns if sem in col and any(dept in col for dept in depts)]
                            filtered_routine = self.routine_df[columns_to_include]
                            routine_table_frame = self.create_routine_table(view_field, filtered_routine)
                            routine_table_frame.pack_forget()
                            self.routine_tables_frame[routine_name] = routine_table_frame

                            progress_index = semesters.index(sem) * 2 if stream == "Arts" else (semesters.index(sem) * 2) + 1
                            self.loading_frame.update_progress(progress_index)
                            self.update()

                    # Set default selection and display
                    self.selected_routine.set(default_routine)
                    self.display_selected_routine()


                elif view_field == "Departments":
                    distinct_depts = [dept['name'] for dept in self.depts_data['depts']]
                    self.loading_frame.set_progress_max(len(distinct_depts))

                    self.routine_tables_frame = {}

                    default_routine = None

                    for i, dept in enumerate(distinct_depts):
                        if default_routine is None:
                            default_routine = dept

                        radio_button = ctk.CTkRadioButton(scrollable_frame, text=dept,
                                                        variable=self.selected_routine,
                                                        value=dept,
                                                        command=self.display_selected_routine)
                        radio_button.pack(side="left", padx=20)

                        columns_to_include = ['Day', 'Time'] + [col for col in self.routine_df.columns if col.startswith(dept)]
                        filtered_routine = self.routine_df[columns_to_include]
                        routine_table_frame = self.create_routine_table(view_field, filtered_routine)
                        routine_table_frame.pack_forget()
                        self.routine_tables_frame[dept] = routine_table_frame

                        self.loading_frame.update_progress(i)
                        self.update()

                    # Set default selection and display
                    self.selected_routine.set(default_routine)
                    self.display_selected_routine()

                elif view_field == "Rooms":
                    def create_floor_schedule(floor_rooms):
                        floor_schedule = pd.DataFrame(columns=['Day', 'Time'] + [f"Room {room}" for room in floor_rooms])
                        floor_schedule['Day'] = self.routine_df['Day']
                        floor_schedule['Time'] = self.routine_df['Time']
                        
                        for _, row in self.routine_df.iterrows():
                            for col in self.routine_df.columns[2:]:
                                if pd.isna(row[col]) or row[col] == '':
                                    continue
                                room_number = " ".join(row[col].__str__().split()[1:])
                                if room_number in floor_rooms:
                                    floor_schedule.loc[(floor_schedule['Day'] == row['Day']) & (floor_schedule['Time'] == row['Time']), f"Room {room_number}"] = col
                        return floor_schedule

                    floors = sorted(set(room['floor'] for room in self.rooms_data['rooms']))
                    self.loading_frame.set_progress_max(len(floors))

                    self.routine_tables_frame = {}

                    default_routine = None

                    for i, floor in enumerate(floors):
                        routine_name = f"Floor {floor}"
                        if default_routine is None:
                            default_routine = routine_name

                        radio_button = ctk.CTkRadioButton(scrollable_frame, text=routine_name,
                                                        variable=self.selected_routine,
                                                        value=routine_name,
                                                        command=self.display_selected_routine)
                        radio_button.pack(side="left", padx=20)

                        floor_rooms = [room['room_id'] for room in self.rooms_data['rooms'] if room['floor'] == floor]
                        floor_schedule = create_floor_schedule(floor_rooms)
                        routine_table_frame = self.create_routine_table(view_field, floor_schedule)
                        routine_table_frame.pack_forget()
                        self.routine_tables_frame[routine_name] = routine_table_frame

                        self.loading_frame.update_progress(i)
                        self.update()

                    # Set default selection and display
                    self.selected_routine.set(default_routine)
                    self.display_selected_routine()

                self.loading_frame.destroy()
                self.loading_frame = None

                messagebox.showinfo("Success", "Routine generated successfully!")
                self.is_running = False
                self.controls_frame_export_routine.pack(pady=10, padx=10)

                # # Bind the event for tab change to apply the custom style
                # tab_control.bind("<<NotebookTabChanged>>", lambda event: self.on_tab_changed(tab_control, style))

            else:
                messagebox.showerror("Error", "The previous operation is still running.")



    def display_selected_routine(self):
        selected_value = self.selected_routine.get()
        for routine_name, table_frame in self.routine_tables_frame.items():
            table_frame.pack_forget()
        if selected_value in self.routine_tables_frame:
            self.routine_tables_frame[selected_value].pack(expand=True, fill="both")

    def create_routine_table(self, view_field, filtered_routine):
        routine_table_frame = ctk.CTkFrame(self.main_frame_routine)
        routine_table = RoutineTable(routine_table_frame, showstatusbar=True, model=RoutineTableModel(filtered_routine))
        routine_table.model.df.set_index(['Day', 'Time'], inplace=True)
        routine_table.showIndex()

        displayed_columns = []
        for i, col in enumerate(filtered_routine.columns):
            if col not in displayed_columns:
                displayed_columns.append(col)
                display_df = filtered_routine[displayed_columns]
                routine_table.model.df = display_df
                routine_table.handleCellEntry(i, col)

                max_len = max(filtered_routine[col].astype(str).map(len).max(), len(col))
                routine_table.columnwidths[col] = max_len * 11.2

        categories = {
            'MAJOR(T)': (0, 8),
            'MAJOR(P)': (0, 8),
            'MINOR(T)': (0, 8),
            'MINOR(P)': (0, 8),
            'MDS(T)': (0, 6),
            'MDS(P)': (0, 6),
            'ENVS(T)': (0, 7)
        }

        def category_check(x):
            for category, (start, end) in categories.items():
                if x.__str__()[start:end] == category[start:end]:
                    return x.__str__()
            return ''

        if view_field != "Rooms":
            routine_table.model.df = routine_table.model.df.map(category_check)

        routine_table.show()

        return routine_table_frame




    def generate_routine(self):
        self.randomize()

    def export_routine(self):
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.save_routine_as_excel(file_path)
            messagebox.showinfo("Success", "Routine saved successfully!")

    def save_routine_as_excel(self, file_path):
        writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
        workbook  = writer.book

        for tab_name, routine_table_frame in self.routine_tables_frame.items():
            df = routine_table_frame.winfo_children()[0].model.df
            df = df.applymap(lambda x: '' if pd.isna(x) else x)

            df.to_excel(writer, sheet_name=tab_name, index=True)
            worksheet = writer.sheets[tab_name]

            formats = {
                'MAJOR(T)': workbook.add_format({'bg_color': '#FFCC99', 'border': 1, 'align': 'center', 'valign': 'vcenter'}),
                'MAJOR(P)': workbook.add_format({'bg_color': '#FF9900', 'border': 1, 'align': 'center', 'valign': 'vcenter'}),
                'MINOR(T)': workbook.add_format({'bg_color': '#99CCFF', 'border': 1, 'align': 'center', 'valign': 'vcenter'}),
                'MINOR(P)': workbook.add_format({'bg_color': '#3399FF', 'border': 1, 'align': 'center', 'valign': 'vcenter'}),
                'MDS(T)': workbook.add_format({'bg_color': '#CC99FF', 'border': 1, 'align': 'center', 'valign': 'vcenter'}),
                'MDS(P)': workbook.add_format({'bg_color': '#9933FF', 'border': 1, 'align': 'center', 'valign': 'vcenter'}),
                'ENVS(T)': workbook.add_format({'bg_color': '#99FF99', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            }

            header_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            index_format = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            
            for col_num, column in enumerate(df.columns):
                worksheet.write(0, col_num + df.index.nlevels, column, header_format)

            for level in range(df.index.nlevels):
                worksheet.write(0, level, df.index.names[level], header_format)
            
            for row_num, (index_values, row_data) in enumerate(zip(df.index, df.values), start=1):
                for col_num, index_value in enumerate(index_values):
                    worksheet.write(row_num, col_num, index_value, index_format)
                for col_num, cell_value in enumerate(row_data, start=df.index.nlevels):
                    applied_format = None
                    for key in formats:
                        if str(cell_value).startswith(key):
                            applied_format = formats[key]
                            break
                    
                    if applied_format:
                        worksheet.write(row_num, col_num, cell_value, applied_format)
                    else:
                        worksheet.write(row_num, col_num, cell_value, workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'}))

            for column in df.columns:
                column_length = max(df[column].astype(str).map(len).max(), len(column))
                col_idx = df.columns.get_loc(column) + df.index.nlevels
                worksheet.set_column(col_idx, col_idx, column_length + 1)
            
            for level in range(df.index.nlevels):
                index_column_length = max(df.index.get_level_values(level).astype(str).map(len).max(), len(df.index.names[level] if df.index.names[level] else ''))
                worksheet.set_column(level, level, index_column_length + 2)
            
            worksheet.freeze_panes(1, df.index.nlevels)

        writer.close()



    
    def create_about_tab(self):
        about_frame = ctk.CTkScrollableFrame(self.about_tab)
        about_frame.pack(pady=10, padx=10, fill="both", expand=True)

        header_label = ctk.CTkLabel(about_frame, text="About SXC Routine Generator", font=ctk.CTkFont(size=30, weight="bold"))
        header_label.pack(pady=20)

        # Introduction Section
        intro_frame = ctk.CTkFrame(about_frame)
        intro_frame.pack(pady=10, padx=10, fill="x")

        intro_label = ctk.CTkLabel(intro_frame, text="SXC Routine Generator is an advanced tool designed to streamline the process of routine generation at St. Xavier's College (Autonomous), Kolkata.\n"
                                                    "It includes various functionalities to load, edit, save, and export data, making routine management efficient and user-friendly. "
                                                    "Our goal is to make routine management efficient and user-friendly, reducing the workload for administrators and ensuring optimal use of resources.\n\n"
                                                    "Key Highlights:\n"
                                                    "- Intuitive Interface: Designed with user experience in mind, providing a clean and easy-to-navigate layout.\n"
                                                    "- Application Appearance: Adapts to system settings to provide a seamless user experience.\n"
                                                    "- Robust Functionality: Handles complex scheduling needs with advanced algorithms to ensure conflict-free schedules.\n"
                                                    "- Flexibility: Easily adaptable to changes and updates, allowing for dynamic adjustments as needed.\n"
                                                    "- Comprehensive Data Management: Supports importing and exporting data, making it easy to integrate with other systems.\n"
                                                    "- User Support: Comprehensive support and resources to help users get the most out of the software.",
                                font=ctk.CTkFont(size=16), justify="left", wraplength=800)
        intro_label.pack(pady=10, padx=10)

        # Features Section
        features_frame = ctk.CTkFrame(about_frame)
        features_frame.pack(pady=10, padx=10, fill="x")

        features_header = ctk.CTkLabel(features_frame, text="Features", font=ctk.CTkFont(size=20, weight="bold"))
        features_header.pack(pady=10)

        features_text = (
            "1. Load Data:\n"
            "   - Load Departments Data: Easily import department details from JSON files to keep information up to date.\n"
            "   - Load Rooms Data: Seamlessly import room details from JSON files for accurate scheduling.\n\n"
            "2. Edit Data:\n"
            "   - Add/Edit Departments: Add new departments or edit existing ones to ensure all data is current.\n"
            "   - Add/Edit Rooms: Modify room details or add new rooms to accommodate scheduling needs.\n"
            "   - Inline Editing: Directly edit cells within the table for quick adjustments.\n\n"
            "3. Generate Routine:\n"
            "   - Automatic Routine Generation: Generate optimized schedules based on the loaded data, considering constraints and requirements.\n"
            "   - Conflict Detection: Automatically detect and resolve scheduling conflicts.\n"
            "   - Resource Allocation: Ensure optimal use of rooms.\n"
            "   - Routine Views: Provides a detailed separate routine for each possible combination of semester, departments and rooms.\n\n"
            "4. Export Routine:\n"
            "   - Excel Export: Export the routine to Excel with the first row and column frozen for better navigation.\n"
            "   - Export View: Export the desired routine from the set of multiple views.\n\n"
            "5. Interactive Data Management:\n"
            "   - Double-click to Sort: Sort data by double-clicking on column headers for better organization.\n"
            "   - Double-click to Edit: Edit cell content directly by double-clicking on them.\n"
            "   - Deselect Rows: Click outside the table to deselect any selected rows, preventing accidental edits.\n\n"
            "6. Progress Tracking:\n"
            "   - During data-intensive operations, a loading interface is shown.\n"
            "   - The loading bar and percentage will update to show progress.\n\n"
            "7. Appearance Settings:\n"
            "   - The application adapts to system appearance settings for a consistent user experience."
        )
        
        features_label = ctk.CTkLabel(features_frame, text=features_text, font=ctk.CTkFont(size=16), justify="left", wraplength=800)
        features_label.pack(pady=10, padx=10)

        # Instructions Section
        instructions_frame = ctk.CTkFrame(about_frame)
        instructions_frame.pack(pady=10, padx=10, fill="x")

        instructions_header = ctk.CTkLabel(instructions_frame, text="Usage Instructions", font=ctk.CTkFont(size=20, weight="bold"))
        instructions_header.pack(pady=10)

        instructions_text = (
            "1. Loading Data:\n"
            "   - Navigate to the 'Departments' tab and click 'Load JSON' to load department data. Ensure the JSON file format is correct.\n"
            "   - Similarly, go to the 'Rooms' tab and click 'Load JSON' to load room data. Verify the room details are accurate.\n"
            "   - Check Data Integrity: After loading, review the data for completeness and accuracy.\n\n"
            "2. Editing Data:\n"
            "   - Use the 'Add Department' button in the 'Departments' tab to add new departments. Fill in all required fields accurately.\n"
            "   - In the 'Rooms' tab, click 'Add Room' to add new rooms. Provide all necessary details for each room.\n"
            "   - Select a row and use the 'Remove Department' or 'Remove Room' buttons to delete entries. Confirm deletions carefully.\n"
            "   - Double-click cells to edit their content directly. Ensure all edits are saved properly.\n\n"
            "3. Generating Routine:\n"
            "   - After loading and editing the necessary data, go to the 'Routine' tab.\n"
            "   - Select the 'Show Routine wrt' button and select the desired parameter to filter the routines accordingly.\n"
            "   - Click the 'Generate Routine' button to create the schedule. Review the generated routine for accuracy.\n"
            "   - Validate Routine: Check the generated routine for conflicts or inconsistencies.\n\n"
            "4. Exporting Routine:\n"
            "   - Use the 'Export Routine' button to save the filtered routine as Excel files.\n"
            "   - Ensure the file name and path are specified correctly during the export process. Verify the exported file for completeness.\n\n"
            "5. Saving Data:\n"
            "   - Save any changes to departmental or room data using the 'Save JSON' button in the respective tabs. Keep backups of important data.\n"
            "   - Regular Backups: Frequently save your data to prevent loss in case of unexpected issues.\n\n"
            "6. Removing Selection:\n"
            "   - Click outside the table in the 'Rooms' tab to deselect any selected rows. This prevents accidental changes or deletions.\n\n"
            "7. Advanced Features:\n"
            "   - Utilize sorting options to organize your data effectively."
        )
        
        instructions_label = ctk.CTkLabel(instructions_frame, text=instructions_text, font=ctk.CTkFont(size=16), justify="left", wraplength=800)
        instructions_label.pack(pady=10, padx=10)

        # Developers Section
        developers_frame = ctk.CTkFrame(about_frame)
        developers_frame.pack(pady=10, padx=10, fill="x")

        developers_header = ctk.CTkLabel(developers_frame, text="About the Developers", font=ctk.CTkFont(size=20, weight="bold"))
        developers_header.pack(pady=10)

        developers_text = (
            "Our team consists of experienced professionals dedicated to creating efficient and user-friendly software solutions. "
            "We aim to make routine management easier and more effective for educational institutions.\n\n"
            "Team Members:\n"
            "- Tridib Paul : (Front-end).\n"
            "- Soham Malakar : (Back-end).\n\n"
            "Our Mission:\n"
            "To provide educational institutions with tools that simplify routine management, enhance productivity, and improve overall efficiency."
        )
        
        developers_details_label = ctk.CTkLabel(developers_frame, text=developers_text, font=ctk.CTkFont(size=16), justify="left", wraplength=800)
        developers_details_label.pack(pady=10, padx=10)


class DepartmentFrame(ctk.CTkFrame):
    def __init__(self, parent, data=None):
        super().__init__(parent)
        self.programs = []
        self.homes = []
        
        self.name_var = ctk.StringVar()
        self.arts_science_var = ctk.StringVar(value="arts")

        self.department_row = ctk.CTkFrame(self, fg_color="transparent")
        self.department_row.grid(row=0, column=0, padx=5, pady=5, sticky="w")

        self.name_label = ctk.CTkLabel(self.department_row, text="Department Name", font=ctk.CTkFont(size=15, weight="bold"))
        self.name_label.pack(side="left", padx=(0, 5))
        
        self.name_entry = ctk.CTkEntry(self.department_row, textvariable=self.name_var, width=150)
        self.name_entry.pack(side="left", padx=(5, 10))

        self.is_arts_dept = ctk.CTkRadioButton(master=self.department_row, text="Arts", variable=self.arts_science_var, value="arts")
        self.is_arts_dept.pack(side="left", padx=(20, 0))
        self.is_science_dept = ctk.CTkRadioButton(master=self.department_row, text="Science", variable=self.arts_science_var, value="science")
        self.is_science_dept.pack(side="left", padx=(0, 20))

        self.add_program_button = ctk.CTkButton(self.department_row, text="Add Program", command=self.add_program)
        self.add_program_button.pack(side="left", padx=(20, 5))
        
        self.remove_program_button = ctk.CTkButton(self.department_row, text="Remove Program", command=self.remove_program)
        self.remove_program_button.pack(side="left", padx=(5, 0))

        self.home_row = ctk.CTkFrame(self, fg_color="transparent")
        self.home_row.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        
        self.homes_label = ctk.CTkLabel(self.home_row, text="Homes")
        self.homes_label.pack(side="left", padx=(0, 5))
        
        self.homes_frame = ctk.CTkFrame(self.home_row, fg_color="transparent")
        self.homes_frame.pack(side="left", padx=5)
        self.remove_home_button = ctk.CTkButton(self.homes_frame, text="-", width=35, command=self.remove_home)
        self.remove_home_button.pack(side="right", padx=(5, 0), anchor="ne")
        self.add_home_button = ctk.CTkButton(self.homes_frame, text="+", width=35, command=self.add_home)
        self.add_home_button.pack(side="right", padx=5, anchor="ne")
        
        self.programs_frame = ctk.CTkFrame(self, fg_color="transparent")
        
        if data:
            self.load_data(data)
        else:
            self.add_home(initial=True)
        
    def add_home(self, initial=False, room_no=""):
        home_frame = ctk.CTkFrame(self.homes_frame, fg_color="transparent")
        home_frame.pack(pady=2, fill="x", expand=True)
        
        home_entry = ctk.CTkEntry(home_frame, width=150)
        home_entry.insert(0, room_no)
        home_entry.pack(side="left", fill="x", expand=True)
        
        if initial:
            self.remove_home_button.configure(state="disabled")
        else:
            self.remove_home_button.configure(state="normal")
        
        self.homes.append(home_frame)
        
    def remove_home(self):
        home_frame = self.homes.pop()
        home_frame.destroy()
        if len(self.homes) == 1:
            self.remove_home_button.configure(state="disabled")


    def add_program(self, program_data=None):
        self.programs_frame.grid(row=2, column=0, columnspan=4, pady=5, sticky="w")
        program_frame = ProgramFrame(self.programs_frame, program_data)
        program_frame.pack(padx=10, pady=10, fill="x", expand=True)
        self.programs.append(program_frame)
        
    def remove_program(self):
        if self.programs:
            program_frame = self.programs.pop()
            program_frame.destroy()
            if len(self.programs) == 0:
                self.programs_frame.grid_forget()
            
        
    def load_data(self, data):
        self.name_var.set(data["name"])
        self.arts_science_var.set(data["type"])
        for i, home in enumerate(data["homes"]):
            if i == 0:
                self.add_home(initial=True, room_no=home["room_id"])
            else:
                self.add_home(room_no=home["room_id"])
        
        for program_data in data["prgms"]:
            self.add_program(program_data)
        
    def get_data(self):
        return {
            "name": self.name_var.get(),
            "type": self.arts_science_var.get(),
            "homes": [{"room_id": home.winfo_children()[0].get()} for home in self.homes],
            "prgms": [program.get_data() for program in self.programs]
        }

class ProgramFrame(ctk.CTkFrame):
    def __init__(self, parent, data=None):
        super().__init__(parent)
        self.semesters = []
        
        self.type_var = ctk.StringVar(value="U.G.")

        self.program_row = ctk.CTkFrame(self, fg_color="transparent")
        self.program_row.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        
        self.name_label = ctk.CTkLabel(self.program_row, text="Program Name")
        self.name_label.pack(side="left", padx=(0, 5))

        self.ug_program_type = ctk.CTkRadioButton(master=self.program_row, text="Undergraduate", variable=self.type_var, value="U.G.")
        self.ug_program_type.pack(side="left", padx=20)
        self.pg_program_type = ctk.CTkRadioButton(master=self.program_row, text="Postgraduate", variable=self.type_var, value="P.G.")
        self.pg_program_type.pack(side="left", padx=20)

        self.add_semester_button = ctk.CTkButton(self.program_row, text="Add Semester", command=self.add_semester)
        self.add_semester_button.pack(side="left", padx=(20, 5))
        
        self.remove_semester_button = ctk.CTkButton(self.program_row, text="Remove Semester", command=self.remove_semester)
        self.remove_semester_button.pack(side="left", padx=(5, 0))
        
        self.semesters_frame = ctk.CTkFrame(self)
        
        if data:
            self.load_data(data)
        
    def add_semester(self, semester_data=None):
        self.semesters_frame.grid(row=1, column=0, columnspan=4, padx=10, pady=10, sticky="w")
        semester_frame = SemesterFrame(self.semesters_frame, semester_data)
        semester_frame.pack(side="left", padx=30, pady=10, fill="y")
        self.semesters.append(semester_frame)
        
    def remove_semester(self):
        if self.semesters:
            semester_frame = self.semesters.pop()
            semester_frame.destroy()
            if len(self.semesters) == 0:
                self.semesters_frame.grid_forget()

        
    def load_data(self, data):
        self.type_var.set(data["type"])
        
        for semester_data in data["sems"]:
            self.add_semester(semester_data)
        
    def get_data(self):
        return {
            "type": self.type_var.get(),
            "sems": [semester.get_data() for semester in self.semesters]
        }

class SemesterFrame(ctk.CTkFrame):
    def __init__(self, parent, data=None):
        super().__init__(parent)
        self.major_practicals = []
        self.minor_practicals = []
        self.mds_practicals = []

        self.no_var = IntegerVar()
        self.strength_var = IntegerVar()

        self.no_label = ctk.CTkLabel(self, text="Semester Number")
        self.no_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")        
        
        self.no_entry = ctk.CTkEntry(self, textvariable=self.no_var, width=150, validate="key", validatecommand=(self.register(self.validate_numeric), '%P'))
        self.no_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        self.no_entry.bind("<FocusOut>", self.default_value)

        self.off_day_entry = ctk.CTkSwitch(self, text="Off-Day required")
        self.off_day_entry.grid(row=0, column=2, padx=5, pady=5, sticky="w")
        
        self.strength_label = ctk.CTkLabel(self, text="Strength")
        self.strength_label.grid(row=1, column=0, padx=5, pady=5, sticky="w")
        
        self.strength_entry = ctk.CTkEntry(self, textvariable=self.strength_var, width=150, validate="key", validatecommand=(self.register(self.validate_numeric), '%P'))
        self.strength_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        self.strength_entry.bind("<FocusOut>", self.default_value)
        
        self.theory_label = ctk.CTkLabel(self, text="Theory")
        self.theory_label.grid(row=2, column=0, padx=5, pady=5, sticky="w")
        
        self.major_theory_var = IntegerVar()
        self.minor_theory_var = IntegerVar()
        self.mds_theory_var = IntegerVar()
        self.envs_theory_var = IntegerVar()
        
        self.major_theory_label = ctk.CTkLabel(self, text="Major")
        self.major_theory_label.grid(row=3, column=1, padx=5, pady=5, sticky="w")
        self.major_theory_entry = ctk.CTkEntry(self, textvariable=self.major_theory_var, width=150, validate="key", validatecommand=(self.register(self.validate_numeric), '%P'))
        self.major_theory_entry.grid(row=3, column=2, padx=5, pady=5, sticky="w")
        self.major_theory_entry.bind("<FocusOut>", self.default_value)
        
        self.minor_theory_label = ctk.CTkLabel(self, text="Minor")
        self.minor_theory_label.grid(row=4, column=1, padx=5, pady=5, sticky="w")
        self.minor_theory_entry = ctk.CTkEntry(self, textvariable=self.minor_theory_var, width=150, validate="key", validatecommand=(self.register(self.validate_numeric), '%P'))
        self.minor_theory_entry.grid(row=4, column=2, padx=5, pady=5, sticky="w")
        self.minor_theory_entry.bind("<FocusOut>", self.default_value)
        
        self.mds_theory_label = ctk.CTkLabel(self, text="MDS")
        self.mds_theory_label.grid(row=5, column=1, padx=5, pady=5, sticky="w")
        self.mds_theory_entry = ctk.CTkEntry(self, textvariable=self.mds_theory_var, width=150, validate="key", validatecommand=(self.register(self.validate_numeric), '%P'))
        self.mds_theory_entry.grid(row=5, column=2, padx=5, pady=5, sticky="w")
        self.mds_theory_entry.bind("<FocusOut>", self.default_value)
        
        self.envs_theory_label = ctk.CTkLabel(self, text="ENVS")
        self.envs_theory_label.grid(row=6, column=1, padx=5, pady=5, sticky="w")
        self.envs_theory_entry = ctk.CTkEntry(self, textvariable=self.envs_theory_var, width=150, validate="key", validatecommand=(self.register(self.validate_numeric), '%P'))
        self.envs_theory_entry.grid(row=6, column=2, padx=5, pady=5, sticky="w")
        self.envs_theory_entry.bind("<FocusOut>", self.default_value)
        
        self.practicals_label = ctk.CTkLabel(self, text="Practical")
        self.practicals_label.grid(row=7, column=0, padx=5, pady=5, sticky="w")
        
        self.major_practicals_label = ctk.CTkLabel(self, text="Major")
        self.major_practicals_label.grid(row=8, column=1, padx=5, pady=5, sticky="w")

        self.add_major_practical_button = ctk.CTkButton(self, text="+", width=50, command=lambda: self.add_practical("major"))
        self.add_major_practical_button.grid(row=8, column=2, padx=5, pady=5, sticky="w")

        self.remove_major_practical_button = ctk.CTkButton(self, text="-",width=50, command=lambda: self.remove_practical("major"))
        self.remove_major_practical_button.grid(row=8, column=3, padx=5, pady=5, sticky="w")
        self.remove_major_practical_button.configure(state="disabled")

        self.major_practicals_frame = ctk.CTkFrame(self, fg_color="transparent")

        self.minor_practicals_label = ctk.CTkLabel(self, text="Minor")
        self.minor_practicals_label.grid(row=11, column=1, padx=5, pady=5, sticky="w")
        
        self.add_minor_practical_button = ctk.CTkButton(self, text="+", width=50, command=lambda: self.add_practical("minor"))
        self.add_minor_practical_button.grid(row=11, column=2, padx=5, pady=5, sticky="w")        

        self.remove_minor_practical_button = ctk.CTkButton(self, text="-", width=50, command=lambda: self.remove_practical("minor"))
        self.remove_minor_practical_button.grid(row=11, column=3, padx=5, pady=5, sticky="w")
        self.remove_minor_practical_button.configure(state="disabled")

        self.minor_practicals_frame = ctk.CTkFrame(self, fg_color="transparent")

        self.mds_practicals_label = ctk.CTkLabel(self, text="MDS")
        self.mds_practicals_label.grid(row=14, column=1, padx=5, pady=5, sticky="w")

        self.add_mds_practical_button = ctk.CTkButton(self, text="+", width=50, command=lambda: self.add_practical("mds"))
        self.add_mds_practical_button.grid(row=14, column=2, padx=5, pady=5, sticky="w")        

        self.remove_mds_practical_button = ctk.CTkButton(self, text="-", width=50, command=lambda: self.remove_practical("mds"))
        self.remove_mds_practical_button.grid(row=14, column=3, padx=5, pady=5, sticky="w")
        self.remove_mds_practical_button.configure(state="disabled")

        self.mds_practicals_frame = ctk.CTkFrame(self, fg_color="transparent")

        if data:
            self.load_data(data)
        
    def add_practical(self, practical_type, practical_data=None):
        if practical_type == "major":
            if len(self.major_practicals) == 0:
                self.cons_major_label = ctk.CTkLabel(self, text="Consecutive Classes")
                self.cons_major_label.grid(row=9, column=2, padx=5, pady=5, sticky="w")
                self.freq_major_label = ctk.CTkLabel(self, text="Frequency")
                self.freq_major_label.grid(row=9, column=3, padx=5, pady=5, sticky="w")
            self.remove_major_practical_button.configure(state="normal")
            self.major_practicals_frame.grid(row=10, column=2, columnspan=8, pady=5, sticky="w")
            major_practical_frame = PracticalFrame(self.major_practicals_frame, practical_data)
            major_practical_frame.pack(fill="x", padx=20, pady=2)
            self.major_practicals.append(major_practical_frame)
        elif practical_type == "minor":
            if len(self.minor_practicals) == 0:
                self.cons_minor_label = ctk.CTkLabel(self, text="Consecutive Classes")
                self.cons_minor_label.grid(row=12, column=2, padx=5, pady=5, sticky="w")
                self.freq_minor_label = ctk.CTkLabel(self, text="Frequency")
                self.freq_minor_label.grid(row=12, column=3, padx=5, pady=5, sticky="w")
            self.remove_minor_practical_button.configure(state="normal")
            self.minor_practicals_frame.grid(row=13, column=2, columnspan=8, pady=5, sticky="w")
            minor_practical_frame = PracticalFrame(self.minor_practicals_frame, practical_data)
            minor_practical_frame.pack(fill="x", padx=20, pady=2)
            self.minor_practicals.append(minor_practical_frame)
        elif practical_type == "mds":
            if len(self.mds_practicals) == 0:
                self.cons_mds_label = ctk.CTkLabel(self, text="Consecutive Classes")
                self.cons_mds_label.grid(row=15, column=2, padx=5, pady=5, sticky="w")
                self.freq_mds_label = ctk.CTkLabel(self, text="Frequency")
                self.freq_mds_label.grid(row=15, column=3, padx=5, pady=5, sticky="w")
            self.remove_mds_practical_button.configure(state="normal")
            self.mds_practicals_frame.grid(row=16, column=2, columnspan=8, pady=5, sticky="w")
            mds_practical_frame = PracticalFrame(self.mds_practicals_frame, practical_data)
            mds_practical_frame.pack(fill="x", padx=20, pady=2)
            self.mds_practicals.append(mds_practical_frame)
        
    def remove_practical(self, practical_type):
        if practical_type == "major" and self.major_practicals:
            major_practical_frame = self.major_practicals.pop()
            major_practical_frame.destroy()
            if len(self.major_practicals) == 0:
                self.remove_major_practical_button.configure(state="disabled")
                self.cons_major_label.grid_forget()
                self.freq_major_label.grid_forget()
                self.major_practicals_frame.grid_forget()
        elif practical_type == "minor" and self.minor_practicals:
            minor_practical_frame = self.minor_practicals.pop()
            minor_practical_frame.destroy()
            if len(self.minor_practicals) == 0:
                self.remove_minor_practical_button.configure(state="disabled")
                self.cons_minor_label.grid_forget()
                self.freq_minor_label.grid_forget()
                self.minor_practicals_frame.grid_forget()
        elif practical_type == "mds" and self.mds_practicals:
            mds_practical_frame = self.mds_practicals.pop()
            mds_practical_frame.destroy()
            if len(self.mds_practicals) == 0:
                self.remove_mds_practical_button.configure(state="disabled")
                self.mds_practicals_frame.grid_forget()
                self.cons_mds_label.grid_forget()
                self.freq_mds_label.grid_forget()

        
    def load_data(self, data):
        self.no_var.set(data["no"])
        if data["hasoffday"]:
            self.off_day_entry.select()
        self.strength_var.set(data["strength"])
        
        self.major_theory_var.set(data["ths"]["major"])
        self.minor_theory_var.set(data["ths"]["minor"])
        self.mds_theory_var.set(data["ths"]["mds"])
        self.envs_theory_var.set(data["ths"]["envs"])
        
        for practical_data in data["prs"]["major"]:
            self.add_practical("major", practical_data)
        
        for practical_data in data["prs"]["minor"]:
            self.add_practical("minor", practical_data)
        
        for practical_data in data["prs"]["mds"]:
            self.add_practical("mds", practical_data)
        
    def get_data(self):
        return {
            "no": self.no_var.get(),
            "hasoffday": self.off_day_entry.get() == 1,
            "strength": self.strength_var.get(),
            "ths": {
                "major": self.major_theory_var.get(),
                "minor": self.minor_theory_var.get(),
                "mds": self.mds_theory_var.get(),
                "envs": self.envs_theory_var.get()
            },
            "prs": {
                "major": [practical.get_data() for practical in self.major_practicals],
                "minor": [practical.get_data() for practical in self.minor_practicals],
                "mds": [practical.get_data() for practical in self.mds_practicals]
            }
        }

    def validate_numeric(self, value):
        if value.isdigit() or value == "":
            return True
        return False

    def default_value(self, event):
        if self.no_entry.get() == "":
            self.no_var.set(0)
        if self.strength_entry.get() == "":
            self.strength_var.set(0)
        if self.major_theory_entry.get() == "":
            self.major_theory_var.set(0)
        if self.minor_theory_entry.get() == "":
            self.minor_theory_var.set(0)
        if self.mds_theory_entry.get() == "":
            self.mds_theory_var.set(0)
        if self.envs_theory_entry.get() == "":
            self.envs_theory_var.set(0)


class PracticalFrame(ctk.CTkFrame):
    def __init__(self, parent, data=None):
        super().__init__(parent)
        
        self.cons_var = IntegerVar()
        self.freq_var = IntegerVar()
        
        self.cons_entry = ctk.CTkEntry(self, textvariable=self.cons_var, width=150, validate="key", validatecommand=(self.register(self.validate_numeric), '%P'))
        self.cons_entry.pack(side="left", padx=5, pady=5)
        self.cons_entry.bind("<FocusOut>", self.default_value)

        self.freq_entry = ctk.CTkEntry(self, textvariable=self.freq_var, width=150, validate="key", validatecommand=(self.register(self.validate_numeric), '%P'))
        self.freq_entry.pack(side="left", padx=5, pady=5)
        self.freq_entry.bind("<FocusOut>", self.default_value)
        
        if data:
            self.load_data(data)
        
    def load_data(self, data):
        self.cons_var.set(data["cons"])
        self.freq_var.set(data["freq"])
        
    def get_data(self):
        return {
            "cons": self.cons_var.get(),
            "freq": self.freq_var.get()
        }

    def validate_numeric(self, value):
        if value.isdigit() or value == "":
            return True
        return False
    
    def default_value(self, event):
        if self.cons_entry.get() == "":
            self.cons_var.set(0)
        if self.freq_entry.get() == "":
            self.freq_var.set(0)


class IntegerVar(ctk.IntVar):
    def get(self):
        value = self._tk.globalgetvar(self._name)
        try:
            return self._tk.getint(value)
        except Exception:
            return 0
        
    
class EditableTreeview(ttk.Treeview):
    def __init__(self, master, headers, data):
        super().__init__(master, columns=headers, show="headings")
        self.headers = headers
        self.data = data
        self.sort_orders = {}
        # Create Treeview style
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Helvetica", 12, "bold"))
        self.init_treeview()
        
    def init_treeview(self):
        for header in self.headers:
            self.heading(header, anchor="center", text=header)
            self.column(header, anchor="center", stretch=True)

        for row in self.data:
            self.insert('', "end", values=row)

        self.bind('<Double-1>', self.on_double_click)

    def on_double_click(self, event):
        region = self.identify_region(event.x, event.y)
        column = self.identify_column(event.x)
        column_index = int(column[1:]) - 1
        if region == "heading":
            column = self.identify_column(event.x)
            self.sort_by_column(column_index)
        else:
            if self.selection():
                item = self.selection()[0]
                x, y, width, height = self.bbox(item, column)
                value = self.item(item, 'values')[column_index]

                # Create a CTkEntry with the same width and height
                entry = ctk.CTkEntry(self, width=width, height=height, validate="key")
                entry.place(x=x, y=y)
                entry.insert(0, value)
                entry.focus()
                entry.bind("<Return>", lambda e: self.update_value(entry, item, column_index))
                entry.bind("<FocusOut>", lambda e: entry.destroy())

                # If the column is for a numeric field, add validation
                if column_index in [1, 4]:
                    entry.configure(validatecommand=(self.register(self.validate_numeric), '%P'))


    def sort_by_column(self, column_index):
            data = [(self.item(item)["values"], item) for item in self.get_children("")]

            # Toggle sort order
            if column_index in self.sort_orders:
                self.sort_orders[column_index] = not self.sort_orders[column_index]
            else:
                self.sort_orders[column_index] = True  # Default to ascending

            reverse = not self.sort_orders[column_index]
            if column_index == 0:
                data.sort(key=lambda x: str(x[0][column_index]), reverse=reverse)
            else:
                data.sort(key=lambda x: x[0][column_index], reverse=reverse)

            # Clear current data
            for item in self.get_children():
                self.delete(item)

            # Insert sorted data
            for values, item in data:
                self.insert("", "end", values=values)


    def update_value(self, entry, item, column_index):
        new_value = entry.get()
        if self.headers[column_index] in ["Room ID"]:
            new_value = new_value if len(new_value) > 0 else "New Room"
        elif self.headers[column_index] in ["Capacity", "Floor"]:
            new_value = new_value if len(new_value) > 0 else 0
        elif self.headers[column_index] in ["AC", "AV"]:
            new_value = "Yes" if new_value.upper() == "Y" or new_value.upper() == "YES" else "No"
        values = list(self.item(item, 'values'))
        values[column_index] = new_value
        self.item(item, values=values)
        entry.destroy()

    def get_data(self):
        data = []
        for row_id in self.get_children():
            row = self.item(row_id)['values']
            data.append(row)
        return data
    
    def validate_numeric(self, value):
        if value.isdigit() or value == "":
            return True
        return False


class LoadingFrame(ctk.CTkFrame):
    def __init__(self, parent):
        super().__init__(parent)
        self.label = ctk.CTkLabel(self, text="Loading... 0%", font=("Arial", 16))
        self.label.pack(pady=10)
        self.progress = ctk.CTkProgressBar(self, orientation="horizontal", mode="determinate")
        self.progress.pack(pady=10)
        
    def set_progress_max(self, max_value):
        self.progress.set(0)
        self.max_value = max_value
        
    def update_progress(self, value):
        progress_value = value / self.max_value
        self.progress.set(progress_value)
        self.label.configure(text=f"Loading... {int(progress_value*100)}%",)


class RoutineTable(Table):
    def __init__(self, parent=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.enable_menus = False
        self.editable = False
        cols = self.model.df.columns.tolist()
        cf = self.columnformats
        cfa = cf['alignment']
        aln = ['center']
        for col in cols:
            cfa[col] = aln

    def doBindings(self):
        """Bind keys and mouse clicks, this can be overriden"""

        self.bind("<Button-1>",self.handle_left_click)
        self.bind("<Double-Button-1>",self.handle_double_click)
        self.bind("<Control-Button-1>", self.handle_left_ctrl_click)
        self.bind("<Shift-Button-1>", self.handle_left_shift_click)

        self.bind("<ButtonRelease-1>", self.handle_left_release)
        if self.ostyp=='darwin':
            #For mac we bind Shift, left-click to right click
            self.bind("<Button-2>", self.handle_right_click)
            self.bind('<Shift-Button-1>',self.handle_right_click)
        else:
            self.bind("<Button-3>", self.handle_right_click)

        self.bind('<B1-Motion>', self.handle_mouse_drag)
        #self.bind('<Motion>', self.handle_motion)

        self.bind("<Control-c>", self.copy)
        #self.bind("<Control-x>", self.deleteRow)
        self.bind("<Delete>", self.clearData)
        self.bind("<Control-v>", self.paste)
        self.bind("<Control-z>", self.undo)
        self.bind("<Control-a>", self.selectAll)
        self.bind("<Control-f>", self.findText)

        self.bind("<Right>", self.handle_arrow_keys)
        self.bind("<Left>", self.handle_arrow_keys)
        self.bind("<Up>", self.handle_arrow_keys)
        self.bind("<Down>", self.handle_arrow_keys)
        #if 'windows' in self.platform:
        self.bind("<MouseWheel>", self.mouse_wheel)
        self.bind('<Button-4>', self.mouse_wheel)
        self.bind('<Button-5>', self.mouse_wheel)
        self.focus_set()
        return


    def handle_right_click(self, event):
        pass

    def sortTable(self, columnIndex=None, ascending=1, index=False):
        pass

    def handle_mouse_drag(self, event):
        if self.startrow is not None and self.startcol is not None:
            super().handle_mouse_drag(event)

    def handleCellEntry(self, row, col):
        self.applyColorMasks(col)
        self.redraw()

    def show(self, callback=None):
        super().show()
        self.rowindexheader.grid_forget()
        self.rowheader.grid_forget()
        self.rowindexheader = IndexHeader(self.parentframe, self, bg='gray25')
        self.rowindexheader.grid(row=0,column=0,rowspan=1,sticky='news')
        self.rowheader = RowHeader(self.parentframe, self, bg='gray25')
        self.rowheader.textcolor = 'white'
        self.rowheader['yscrollcommand'] = self.Yscrollbar.set
        self.rowheader.grid(row=1,column=0,rowspan=1,sticky='news')

    def applyColorMasks(self, col):
        slice_lengths = {
            logic.ClassType.NA: 2,
            logic.ClassType.MAJOR: 8,
            logic.ClassType.MINOR: 8,
            logic.ClassType.MDS: 6,
            logic.ClassType.ENVS: 7
        }
        masks = {
            'lightgreen': lambda x: x.__str__()[0:slice_lengths[logic.ClassType.MAJOR]] == logic.ClassType.MAJOR.name + '(T)',
            'lightcoral': lambda x: x.__str__()[0:slice_lengths[logic.ClassType.MAJOR]] == logic.ClassType.MAJOR.name + '(P)',
            'lightblue': lambda x: x.__str__()[0:slice_lengths[logic.ClassType.MINOR]] == logic.ClassType.MINOR.name + '(T)',
            'pink': lambda x: x.__str__()[0:slice_lengths[logic.ClassType.MINOR]] == logic.ClassType.MINOR.name + '(P)',
            'yellow': lambda x: x.__str__()[0:slice_lengths[logic.ClassType.MDS]] == logic.ClassType.MDS.name + '(T)',
            'orange': lambda x: x.__str__()[0:slice_lengths[logic.ClassType.MDS]] == logic.ClassType.MDS.name + '(P)',
            'lightgray': lambda x: x.__str__()[0:slice_lengths[logic.ClassType.ENVS]] == logic.ClassType.ENVS.name + '(T)',
            'white': lambda x: not any(
                x.__str__()[0:slice_lengths[ct]] == ct.name + '(T)' or
                x.__str__()[0:slice_lengths[ct]] == ct.name + '(P)'
                for ct in logic.ClassType
            )
        }
        for color, mask in masks.items():
            mask_result = self.model.df[col].apply(mask)
            self.setColorByMask(col, mask_result, color)


# Create a custom TableModel to handle the DataFrame with index
class RoutineTableModel(TableModel):
    def __init__(self, dataframe):
        super().__init__(dataframe)
        self.df = dataframe

    def getValueAt(self, rowIndex, colIndex):
        if colIndex < 0:
            return ""
        else:
            return super().getValueAt(rowIndex, colIndex)


# Create and run the application
if __name__ == "__main__":
    app = Routine_Generator()

    # Initialize the application
    ctk.set_appearance_mode("System")

    app.mainloop()
