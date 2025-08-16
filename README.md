# Hospital Management System (HMS)

### Project Overview

This is an end-to-end Hospital Management System built using Microsoft Access. The project demonstrates practical database design principles, form automation with VBA (Visual Basic for Applications), and report generation. It is designed to be a lightweight, offline solution for managing patient records, appointments, and billing in a small clinic or hospital setting.

### Key Features

* **Patient & Doctor Management:** Forms for adding, updating, and managing patient and doctor records. Includes a friendly, calculated Patient ID for easy lookup.
* **Intelligent Appointment Scheduling:** An appointment form with combo box lookups for patients and doctors. VBA code prevents scheduling conflicts for doctors by checking for overlapping appointments.
* **Automated Billing & Invoicing:**
    * A main form (`frmBills`) with a subform (`sfrmBillItems`) for adding services to an invoice.
    * VBA automatically populates `UnitPrice` and calculates `LineTotal` as services are added.
    * A `Recalculate` button computes the `Subtotal`, `Tax`, and `Total` for the entire bill.
* **Dynamic Queries:** Custom queries to retrieve specific data, including a daily doctor's schedule, patient visit history (using parameters), and monthly revenue totals.
* **Professional Reports:** Print-ready reports for daily appointments, patient lists, and a detailed invoice (using a main report and subreport).
* **User-Friendly Interface:** A main menu (switchboard) with command buttons for one-click navigation to all major forms and reports.

### Technologies Used

* **Microsoft Access:** The core database engine and development environment.
* **Access SQL:** For data retrieval and dynamic queries.
* **VBA (Visual Basic for Applications):** For form automation, validation, and custom business logic.

### Setup and Installation

1.  **Download the Database:** Download the `HospitalMS.accdb` file from this repository.
2.  **Enable Macros:** Microsoft Access will block VBA code by default. To enable it, open the database file and follow the security warning prompts to enable content. For long-term use, add the folder containing the database to Access's **Trusted Locations** (File > Options > Trust Center > Trust Center Settings > Trusted Locations).
3.  **Split the Database (Recommended):** For better stability and to allow multiple users, it is highly recommended to split the database into a front-end (forms, reports, queries) and a back-end (tables). You can do this by navigating to `Database Tools > Access Database`.

### Usage

* Open the `frmMainMenu` form to navigate the application.
* Use the forms to input sample data for `Patients`, `Doctors`, and `Services` first.
* Then, add `Appointments` and create `Bills` to test all the automation features.
* Run the queries and reports to view your data.

### Screenshots

_You can add your screenshots here to show the key features of your project._

* **Main Invoice Form:** _(Show the `frmBill` with populated data)_
* **Relationships Window:** _(Show the tables and links)_
* **Appointment Form:** _(Show the `frmAppointments` form with the combo boxes)_
* **Example Report:** _(Show a print preview of the invoice or daily schedule report)_

### Credits

This project was developed by Chethan Vakit.

---
