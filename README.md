# 🏥 Hospital Management System

A modern desktop **Hospital Management System** developed using **Python**, **CustomTkinter**, and **MySQL**.

The application provides a clean graphical user interface for managing patient records with full **CRUD operations**, **data validation**, **search functionality**, and **CSV/Excel export** support.

---

## 📸 Application Preview

<p align="center">
  <img src="screenshots/hospital_management.png" alt="Hospital Management System" width="100%">
</p>

---

# ✨ Features

### 👤 Patient Management
- Add new patient records
- Update existing patient information
- Delete patient records
- View all stored records
- Search patients by Reference Number

### 💊 Prescription Management
- Generate prescription details
- Display medication information
- Store prescription-related data

### ✅ Data Validation
- Issue Date validation
- Expiry Date validation
- Date of Birth validation
- User-friendly error messages

### 📁 Export Options
- Export patient records to CSV
- Export patient records to formatted Excel (.xlsx)

### 🖥 User Interface
- Modern interface built with **CustomTkinter**
- Responsive table view
- Clean and intuitive layout

### 💾 Database
- MySQL integration
- Secure database connection using configuration file
- Full CRUD operations

---

# 🛠 Technologies Used

| Technology | Purpose |
|------------|---------|
| Python | Application development |
| CustomTkinter | Modern GUI |
| Tkinter (Treeview) | Data table |
| MySQL | Database |
| mysql-connector-python | Database connection |
| OpenPyXL | Excel export |
| CSV | CSV export |
| Object-Oriented Programming | Application architecture |

---

# 📂 Project Structure

```text
Hospital-Management-System-Python
│
├── main.py
├── config.py
├── requirements.txt
├── .gitignore
├── README.md
│
├── sql/
│   └── hospital.sql
│
└── screenshots/
    └── hospital_management.png
```

---

# 🚀 Getting Started

## 1️⃣ Clone the repository

```bash
git clone https://github.com/bluecaterpillarz/hospital-management-system-python.git
```

---

## 2️⃣ Navigate to the project folder

```bash
cd hospital-management-system-python
```

---

## 3️⃣ Install the required packages

```bash
pip install -r requirements.txt
```

---

## 4️⃣ Configure the database

Create a file named **config.py** using the template below:

```python
DB_CONFIG = {
    "host": "localhost",
    "username": "your_username",
    "password": "your_password",
    "database": "your_database"
}
```

---

## 5️⃣ Create the database

Import the SQL script located in the **sql/** directory into your MySQL server.

---

## 6️⃣ Run the application

```bash
python main.py
```

---

# 📋 Requirements

- Python 3.10+
- MySQL Server
- Required Python packages listed in `requirements.txt`

---

# 📈 Future Improvements

- Login & Authentication
- User Roles (Admin / Doctor / Receptionist)
- Appointment Management
- Patient History
- Dashboard & Statistics
- PDF Prescription Export
- Dark Mode
- Search by Patient Name
- Search by NHS Number

---

# 💡 Highlights

- ✔ Modern desktop GUI
- ✔ Object-Oriented Design
- ✔ MySQL Database Integration
- ✔ Data Validation
- ✔ CSV Export
- ✔ Excel Export
- ✔ Search Functionality
- ✔ Error Handling

---

# 👩‍💻 Author

**Zeynep Ercan**

🔗 GitHub  
https://github.com/bluecaterpillarz

🔗 LinkedIn  
https://www.linkedin.com/in/zeynep-ercan-bluecaterpillarz/

---

# 📄 License

This project was developed for educational purposes and portfolio presentation.
