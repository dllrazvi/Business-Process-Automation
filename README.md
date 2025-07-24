# Business Process Automation

This project demonstrates a simple RPA (Robotic Process Automation) scenario developed in UiPath, aimed at automating job deadline management using Excel spreadsheets. The automation helps identify job listings that are about to expire soon and highlights them for quick attention.

## 📁 Project Overview

The automation performs the following tasks:

1. **Reads Excel data** from the `NoDeskJobs.xlsx` file.
2. **Iterates through job entries**, checking each application deadline.
3. **Highlights in red** the rows where the deadline is within the next 3 days.
4. **Writes back the updated table** to the same Excel file.

This workflow simulates how a user might quickly spot urgent job applications from a long list.

## 🔧 Tools & Technologies

- **UiPath Studio** – for designing the workflow.
- **UiPath Excel Activities** – for reading, writing, and formatting Excel files.
- **Excel Spreadsheet** – used as the job list data source.
- **Visual Basic Expression Language** – used in conditional checks inside the workflow.

## 📂 File Structure

- `Main.xaml` – The main workflow file.
- `HighlightDeadlinesWorkflow.xaml` – A subcomponent focused on identifying deadlines.
- `HipoJobs.xlsx` / `NoDeskJobs.xlsx` – Input Excel files with job postings.
- `project.json` – Contains metadata and dependency information for the UiPath project.

## 🧠 Logic Highlight

Inside the `Main.xaml` workflow:
- A **Read Range** activity loads job listings.
- A **For Each Row** loop checks if the deadline is ≤ 3 days.
- If the condition is met, the row is **highlighted in red** using a **Set Range Color** activity.
- The table is then written back using a **Write Range** activity.

## 📸 Screenshots

Visual examples of the workflow and interface can be found in the [Screenshots folder](https://github.com/dllrazvi/Business-Process-Automation/tree/main/Screenshots).

---

This automation represents an entry-level UiPath solution but demonstrates fundamental RPA principles such as data input/output, conditional logic, and UI-based automation through Excel.
