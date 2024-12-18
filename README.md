# Basic School Management System

A simple School Management System implemented in Python. This system allows for the management of Students, Staffs, and Courses using a command-line interface (CLI) and stores data in an Excel file.

## Features

The system provides basic CRUD (Create, Read, Update, Delete) functionalities for three main fields:

1. **Students**
   - Add new students
   - Display all students
   - Search for a student by ID
   - Update student details by ID
   - Delete a student by ID

2. **Staff**
   - Add new staff members
   - Display all staff members
   - Search for staff by ID
   - Update staff details by ID
   - Delete staff by ID

3. **Courses**
   - Add new courses
   - Display all courses
   - Search for a course by ID
   - Update course details by ID
   - Delete a course by ID

## Technologies Used

- **Python**: The programming language used for the implementation.
- **tabulate**: A library used to display data in a tabular format in the CLI.
- **openpyxl**: A library used for reading and writing Excel files, which serves as the database for the system.

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/Manish-Let-It-Be/School-Management-System.git
   cd school-management-system
   ```

2. Install the required libraries:
   ```bash
   pip install tabulate openpyxl
   ```

3. Run the program:
   ```bash
   python sms.py
   ```

## Screenshots

### CLI Interface
![CLI Interface](screenshots/cli_interface.png)

### Excel File
![Excel File](screenshots/excel_file.png)

## Usage

Once the program is running, you will be presented with a menu to choose between managing students, staff, or courses. Follow the prompts to perform the desired operations.
