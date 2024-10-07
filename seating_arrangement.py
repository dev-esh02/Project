import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import random

# Global dictionary to store classroom data
classrooms = {}

# Function to open file dialog and select an Excel file for students
def browse_file():
    filename = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    if filename:
        excel_file_entry.delete(0, tk.END)
        excel_file_entry.insert(0, filename)

# Function to load classroom data from an Excel sheet
def load_classroom_excel():
    filename = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
    )
    if filename:
        try:
            classrooms_df = pd.read_excel(filename)
            if {'Classroom', 'Rows', 'Cols'}.issubset(classrooms_df.columns):
                for _, row in classrooms_df.iterrows():
                    class_name = row['Classroom']
                    rows = int(row['Rows'])
                    cols = int(row['Cols'])
                    classrooms[class_name] = {'rows': rows, 'cols': cols, 'capacity': rows * cols}
                    classroom_listbox.insert(tk.END, f"{class_name} - {rows}x{cols}")
            else:
                messagebox.showerror("Error", "The Excel sheet must contain 'Classroom', 'Rows', and 'Cols' columns.")
        except Exception as e:
            messagebox.showerror("Error", str(e))

# Function to manually add a classroom
def add_classroom():
    class_name = classroom_name_entry.get()
    rows = int(classroom_rows_entry.get())
    cols = int(classroom_cols_entry.get())

    if class_name and rows > 0 and cols > 0:
        classrooms[class_name] = {'rows': rows, 'cols': cols, 'capacity': rows * cols}
        classroom_listbox.insert(tk.END, f"{class_name} - {rows}x{cols}")
        classroom_name_entry.delete(0, tk.END)
        classroom_rows_entry.delete(0, tk.END)
        classroom_cols_entry.delete(0, tk.END)
    else:
        messagebox.showwarning("Invalid Input", "Please enter valid classroom details")

# Function to delete a selected classroom
def delete_classroom():
    try:
        selected_index = classroom_listbox.curselection()
        if selected_index:
            selected_class = classroom_listbox.get(selected_index)
            class_name = selected_class.split(" - ")[0]

            del classrooms[class_name]
            classroom_listbox.delete(selected_index)
        else:
            messagebox.showwarning("Selection Error", "Please select a classroom to delete")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Function to load student data from an Excel sheet
def load_students(excel_file):
    students_df = pd.read_excel(excel_file)
    return students_df

# Function to group students by year and branch
def group_students_by_branch_and_year(students_df):
    grouped_students = students_df.groupby(['Year', 'Branch'])
    student_groups = [group for _, group in grouped_students]
    for group in student_groups:
        group.reset_index(drop=True, inplace=True)
    return student_groups

# Function to check for seat conflicts
def is_adjacent_seat_conflict(seating, r, c, student, students_df):
    rows = len(seating)
    cols = len(seating[0])
    adjacent_coords = [
        (r, c-1), (r, c+1),
        (r-1, c-1), (r-1, c+1), (r+1, c-1), (r+1, c+1)
    ]
    
    student_year = students_df.loc[students_df['Enrollment'] == student, 'Year'].values[0]
    student_branch = students_df.loc[students_df['Enrollment'] == student, 'Branch'].values[0]

    for adj_r, adj_c in adjacent_coords:
        if 0 <= adj_r < rows and 0 <= adj_c < cols:
            adjacent_student = seating[adj_r][adj_c]
            if adjacent_student is not None:
                adjacent_year = students_df.loc[students_df['Enrollment'] == adjacent_student, 'Year'].values[0]
                adjacent_branch = students_df.loc[students_df['Enrollment'] == adjacent_student, 'Branch'].values[0]
                if adjacent_year == student_year and adjacent_branch == student_branch:
                    return True
    return False

# Function to create the seating plan
def create_seating_plan(classrooms, students_df):
    student_groups = group_students_by_branch_and_year(students_df)
    random.shuffle(student_groups)
    
    seating_plan = {}
    student_idx = 0
    group_idx = 0
    
    for class_name, details in classrooms.items():
        rows = details['rows']
        cols = details['cols']
        seating = [[None for _ in range(cols)] for _ in range(rows)]
        
        for r in range(rows):
            for c in range(cols):
                placed = False
                attempts = 0
                while not placed and attempts < len(student_groups):
                    current_group = student_groups[group_idx]
                    if student_idx < len(current_group):
                        student = current_group.iloc[student_idx]['Enrollment']
                        if not is_adjacent_seat_conflict(seating, r, c, student, students_df):
                            seating[r][c] = student
                            placed = True
                        student_idx += 1
                    else:
                        group_idx = (group_idx + 1) % len(student_groups)
                        student_idx = 0
                    attempts += 1
        seating_plan[class_name] = seating
    return seating_plan

# Function to save the seating plan to an Excel file
def save_seating_plan(seating_plan):
    output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if output_file:
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            for class_name, seating in seating_plan.items():
                df = pd.DataFrame(seating)
                df.to_excel(writer, sheet_name=class_name, index=False, header=False)
        messagebox.showinfo("Success", f"Seating plan saved to {output_file}")

# Function to display seating arrangements in a new window
def display_seating_arrangements(seating_plan, students_df):
    seating_window = tk.Toplevel(root)
    seating_window.title("Seating Arrangements")
    seating_window.geometry("800x600")  # Increased window size

    def show_class_seating(class_name):
        seating_text.delete("1.0", tk.END)

        seating = seating_plan[class_name]

        def on_seat_click(student):
            if student:
                student_info = students_df.loc[students_df['Enrollment'] == student]
                if not student_info.empty:
                    student_name = student_info.iloc[0]['Name']
                    student_branch = student_info.iloc[0]['Branch']
                    student_year = student_info.iloc[0]['Year']
                    messagebox.showinfo("Student Info", f"Name: {student_name}\nBranch: {student_branch}\nYear: {student_year}")
                else:
                    messagebox.showwarning("Not Found", "Student information not found.")
            else:
                messagebox.showinfo("Empty Seat", "This seat is empty.")

        seating_text.delete("1.0", tk.END)
        seating_text.insert(tk.END, f"Seating arrangement for {class_name}:\n\n")
        
        for r, row in enumerate(seating):
            seating_text.insert(tk.END, f"Row {r + 1}:\n")
            for c, student in enumerate(row):
                seat_button = tk.Button(
                    seating_window, 
                    text=str(student) if student else "Empty",
                    width=10,
                    command=lambda s=student: on_seat_click(s)
                )
                seating_text.window_create(tk.END, window=seat_button)
            seating_text.insert(tk.END, "\n\n")

        export_button.config(command=lambda: save_seating_plan({class_name: seating}))

    button_frame = tk.Frame(seating_window)
    button_frame.pack(side=tk.LEFT, padx=10, pady=10)

    seating_text = tk.Text(seating_window, width=60, height=25)  # Increased text area size
    seating_text.pack(side=tk.RIGHT, padx=10, pady=10)

    for class_name in seating_plan:
        class_button = tk.Button(button_frame, text=class_name, command=lambda cn=class_name: show_class_seating(cn))
        class_button.pack(pady=5)

    export_button = tk.Button(seating_window, text="Export to Excel", state=tk.NORMAL)
    export_button.pack(pady=5)

# Run seating arrangement and display the results
def run_seating_arrangement():
    try:
        excel_file = excel_file_entry.get()
        students_df = load_students(excel_file)
        seating_plan = create_seating_plan(classrooms, students_df)
        
        display_seating_arrangements(seating_plan, students_df)
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Function to search for a student by enrollment number
def search_student():
    enrollment = enrollment_search_entry.get()
    student_info = students_df.loc[students_df['Enrollment'] == enrollment]
    
    if not student_info.empty:
        student_name = student_info.iloc[0]['Name']
        student_branch = student_info.iloc[0]['Branch']
        student_year = student_info.iloc[0]['Year']
        messagebox.showinfo("Student Info", f"Name: {student_name}\nBranch: {student_branch}\nYear: {student_year}")
    else:
        messagebox.showwarning("Not Found", "Student not found.")

# Main window
root = tk.Tk()
root.title("Seating Arrangement System")

# Load Students Frame
frame_students = tk.Frame(root)
frame_students.pack(pady=10)

tk.Label(frame_students, text="Student Excel File:").grid(row=0, column=0)
excel_file_entry = tk.Entry(frame_students, width=40)
excel_file_entry.grid(row=0, column=1)
browse_button = tk.Button(frame_students, text="Browse", command=browse_file)
browse_button.grid(row=0, column=2)

# Load Classrooms Frame
frame_classrooms = tk.Frame(root)
frame_classrooms.pack(pady=10)

classroom_listbox = tk.Listbox(frame_classrooms, width=50)
classroom_listbox.grid(row=0, column=0, rowspan=4, padx=10)

tk.Label(frame_classrooms, text="Classroom Name:").grid(row=0, column=1)
classroom_name_entry = tk.Entry(frame_classrooms)
classroom_name_entry.grid(row=0, column=2)

tk.Label(frame_classrooms, text="Rows:").grid(row=1, column=1)
classroom_rows_entry = tk.Entry(frame_classrooms)
classroom_rows_entry.grid(row=1, column=2)

tk.Label(frame_classrooms, text="Cols:").grid(row=2, column=1)
classroom_cols_entry = tk.Entry(frame_classrooms)
classroom_cols_entry.grid(row=2, column=2)

add_button = tk.Button(frame_classrooms, text="Add Classroom", command=add_classroom)
add_button.grid(row=3, column=1, pady=5)

delete_button = tk.Button(frame_classrooms, text="Delete Classroom", command=delete_classroom)
delete_button.grid(row=3, column=2, pady=5)

load_classroom_button = tk.Button(frame_classrooms, text="Load Classrooms from Excel", command=load_classroom_excel)
load_classroom_button.grid(row=4, column=0, columnspan=3, pady=10)

# Search Frame for enrollment
frame_search = tk.Frame(root)
frame_search.pack(pady=10)

tk.Label(frame_search, text="Search by Enrollment Number:").grid(row=0, column=0)
enrollment_search_entry = tk.Entry(frame_search, width=20)
enrollment_search_entry.grid(row=0, column=1)

search_button = tk.Button(frame_search, text="Search", command=search_student)
search_button.grid(row=0, column=2)

# Run Button
run_button = tk.Button(root, text="Run Seating Arrangement", command=run_seating_arrangement)
run_button.pack(pady=20)

root.mainloop()
