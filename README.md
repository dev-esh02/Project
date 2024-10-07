# Examination Seating Arrangement System

## Description
The **Examination Seating Arrangement System** is a Python-based application that automates the seating arrangement process for students during exams. It ensures that no two students from the same branch and year are seated next to each other. The application allows manual addition or deletion of classrooms, loads classroom and student data from Excel files, and provides a final seating arrangement that can be exported to Excel.

## Features
- **Manual Classroom Management**: Add or remove classrooms with specified rows and columns.
- **Load Classroom Data**: Load classroom data from an Excel sheet.
- **Load Student Data**: Load student data from an Excel sheet, containing details like enrollment number, name, branch, and year.
- **Seating Arrangement Generation**: Automatically generate seating arrangements while ensuring students from the same branch and year are not adjacent.
- **Student Information**: Click on any seat to view the student's details.
- **Search by Enrollment Number**: Search for a student by entering their enrollment number.
- **Export to Excel**: Save the final seating arrangement for each classroom as an Excel file.

## Installation

1. **Clone the repository**:
    ```bash
    git clone https://github.com/your-repository-url.git
    cd seating_arrangement_system
    ```

2. **Install dependencies**:
    Create a virtual environment (optional but recommended):
    ```bash
    python -m venv venv
    source venv/bin/activate  # For Windows: venv\Scripts\activate
    ```

    Install the required packages:
    ```bash
    pip install -r requirements.txt
    ```

3. **Install additional libraries**:
    The application requires the following libraries:
    - `tkinter`: For the GUI
    - `pandas`: For data manipulation
    - `openpyxl`: For reading Excel files
    - `xlsxwriter`: For writing Excel files

    You can install these packages manually:
    ```bash
    pip install pandas openpyxl XlsxWriter
    ```

## Usage

1. **Run the Application**:
    ```bash
    python seating_arrangement.py
    ```

2. **Add Classrooms**:
   - You can add classrooms manually by specifying the classroom name, rows, and columns, or load classrooms from an Excel file (which must contain columns: `Classroom`, `Rows`, `Cols`).

3. **Load Students**:
   - Browse and select an Excel file containing student information (with columns: `Enrollment`, `Name`, `Branch`, and `Year`).

4. **Generate Seating Arrangement**:
   - Once classrooms and students are loaded, click "Run Seating Arrangement" to automatically generate the seating plan.

5. **View and Export Seating**:
   - Click on any classroom to view its seating arrangement.
   - Click on any seat to see the details of the student assigned to that seat.
   - Export the seating arrangement to an Excel file.

## Excel File Formats

### Classroom Excel Format:
| Classroom | Rows | Cols |
|-----------|------|------|
| LT001     | 10   | 10   |
| LT002     | 10   | 10   |

### Student Excel Format:
| Enrollment    | Name            | Branch | Year |
|---------------|-----------------|--------|------|
| 0801CS221049  | Devesh Sharma   | CSE    | 3rd  |
| 0801CS221079  | Kushagra Saxena | CSE    | 3rd  |

## Known Issues
- Ensure that the Excel files used for classrooms and students adhere strictly to the formats mentioned above.
- `xlsxwriter` must be installed to export the seating plan.

## Acknowledgments
- **Pandas** for handling Excel data.
- **Tkinter** for the GUI.
