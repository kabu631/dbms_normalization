# Prerequisite: pip install python-docx
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
import os

# ---------------------------------------------------------------------------
# Styling and Helper Functions (Matching format.docx: Arial, Table Grid)
# ---------------------------------------------------------------------------

def set_cell_background(cell, fill_hex):
    tcPr = cell._element.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=70, bottom=70, left=100, right=100):
    tcPr = cell._element.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)

def make_row_cant_split(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

def make_row_header(row):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

def add_grid_table(doc, headers, rows, col_widths=None, header_bg="F2F2F2"):
    table = doc.add_table(rows=len(rows) + 1, cols=len(headers))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    # Header Row
    hdr_row = table.rows[0]
    make_row_header(hdr_row)
    make_row_cant_split(hdr_row)
    for i, title in enumerate(headers):
        cell = hdr_row.cells[i]
        cell.text = title
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        set_cell_background(cell, header_bg)
        set_cell_margins(cell, 80, 80, 110, 110)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15
        for run in p.runs:
            run.font.name = "Arial"
            run.font.bold = True
            run.font.size = Pt(9.5)
            run.font.color.rgb = RGBColor(0, 0, 0)

    # Data Rows
    for r_idx, row_data in enumerate(rows):
        row = table.rows[r_idx + 1]
        make_row_cant_split(row)
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.text = str(val)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_margins(cell, 60, 60, 100, 100)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.15
            for run in p.runs:
                run.font.name = "Arial"
                run.font.size = Pt(9)
                run.font.color.rgb = RGBColor(30, 30, 30)

    # Set Column Widths if provided
    if col_widths and len(col_widths) == len(headers):
        for row in table.rows:
            for i, w in enumerate(col_widths):
                row.cells[i].width = Inches(w)

    p_space = doc.add_paragraph()
    p_space.paragraph_format.space_before = Pt(0)
    p_space.paragraph_format.space_after = Pt(3)
    return table

def add_para(doc, text, bold=False, italic=False, size=10, align=None, space_before=2, space_after=3):
    p = doc.add_paragraph()
    if align:
        p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    run = p.add_run(text)
    run.font.name = "Arial"
    run.font.bold = bold
    run.font.italic = italic
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor(0, 0, 0)
    return p

def add_q_heading(doc, text):
    return add_para(doc, text, bold=True, size=11, space_before=11, space_after=2)

def add_lead_para(doc, lead, body):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.15
    r_lead = p.add_run(lead + ": ")
    r_lead.font.name = "Arial"
    r_lead.font.bold = True
    r_lead.font.size = Pt(9.5)
    r_lead.font.color.rgb = RGBColor(0, 0, 0)

    r_body = p.add_run(body)
    r_body.font.name = "Arial"
    r_body.font.size = Pt(9.5)
    r_body.font.color.rgb = RGBColor(30, 30, 30)
    return p

# ---------------------------------------------------------------------------
# Load format.docx Template
# ---------------------------------------------------------------------------
template_path = "format_template_backup.docx" if os.path.exists("format_template_backup.docx") else "format.docx"
doc = Document(template_path)

# Remove old brief answers starting from ANSWER KEY & STEP-BY-STEP SOLUTIONS
ans_idx = None
for i, p in enumerate(doc.paragraphs):
    if "ANSWER KEY" in p.text:
        ans_idx = i
        break

if ans_idx is not None:
    paras_to_remove = doc.paragraphs[ans_idx:]
    for p in paras_to_remove:
        p._element.getparent().remove(p._element)

# Remove any trailing empty paragraphs at the end of Part A
while doc.paragraphs and not doc.paragraphs[-1].text.strip():
    p = doc.paragraphs[-1]
    p._element.getparent().remove(p._element)

# Add Page Break before Part B
doc.add_page_break()

# ---------------------------------------------------------------------------
# Part B: Answer Key & Step-by-Step Solutions Header
# ---------------------------------------------------------------------------
add_para(doc, "ANSWER KEY & STEP-BY-STEP SOLUTIONS", bold=True, size=18, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=6, space_after=2)
add_para(doc, "Suggested solutions. Equivalent lossless and dependency-preserving 3NF decompositions may be accepted when correctly justified.", italic=True, size=10, align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=8)

# ===========================================================================
# Solution 1: Student–Course Repeating Groups
# ===========================================================================
add_q_heading(doc, "Q1. Student–Course Repeating Groups")
add_lead_para(doc, "1NF Violation", "The original STUDENT_RAW relation violates 1NF because the 'Courses' column contains non-atomic, comma-delimited strings ('DBMS, OS'). 1NF strictly requires that every column contain only atomic (indivisible) values with no repeating groups.")
add_lead_para(doc, "Candidate Key", "(Student_ID, Course). Once flattened, Course is required alongside Student_ID to uniquely identify each tuple.")
add_para(doc, "Flattened 1NF Relation Table (Atomic Tuples):", bold=True, size=9.5, space_before=2, space_after=2)
add_grid_table(doc, 
    ["Student_ID (PK)", "Student_Name", "Course (PK)"],
    [
        ["S01", "Ram", "DBMS"],
        ["S01", "Ram", "OS"],
        ["S02", "Sita", "DBMS"],
        ["S02", "Sita", "AI"],
        ["S03", "Hari", "OS"],
        ["S03", "Hari", "AI"]
    ],
    col_widths=[2.2, 2.5, 2.3]
)

# ===========================================================================
# Solution 2: Employee–Department
# ===========================================================================
add_q_heading(doc, "Q2. Employee–Department")
add_lead_para(doc, "Candidate Key", "Emp_ID. Since Emp_ID determines Emp_Name and Dept_ID, and Dept_ID determines Dept_Name and Dept_Location, Emp_ID+ = {Emp_ID, Emp_Name, Dept_ID, Dept_Name, Dept_Location} = R.")
add_lead_para(doc, "Transitive Dependency", "Emp_ID → Dept_ID and Dept_ID → {Dept_Name, Dept_Location}. The non-key attributes Dept_Name and Dept_Location are transitively dependent on primary key Emp_ID through Dept_ID (which is not a superkey). This violates 3NF.")
add_para(doc, "3NF Relational Decomposition Table:", bold=True, size=9.5, space_before=2, space_after=2)
add_grid_table(doc,
    ["Relation Name", "Primary Key (PK)", "Foreign Key (FK)", "Non-Key Attributes", "Dependency Preserved"],
    [
        ["EMPLOYEE", "Emp_ID", "Dept_ID → DEPARTMENT(Dept_ID)", "Emp_Name", "Emp_ID → Emp_Name, Dept_ID"],
        ["DEPARTMENT", "Dept_ID", "None", "Dept_Name, Dept_Location", "Dept_ID → Dept_Name, Dept_Location"]
    ],
    col_widths=[1.5, 1.3, 1.8, 1.3, 1.1]
)

# ===========================================================================
# Solution 3: Student Marks
# ===========================================================================
add_q_heading(doc, "Q3. Student Marks")
add_lead_para(doc, "Candidate Key", "(Student_ID, Subject_ID). Both attributes are required together to uniquely determine Marks.")
add_lead_para(doc, "Partial Dependencies", "Student_ID → Student_Name and Subject_ID → Subject_Name. Both determinants are strict proper subsets of candidate key (Student_ID, Subject_ID), violating 2NF.")
add_para(doc, "2NF & 3NF Relational Decomposition Table:", bold=True, size=9.5, space_before=2, space_after=2)
add_grid_table(doc,
    ["Relation Name", "Primary Key (PK)", "Foreign Key (FK)", "Non-Key Attributes", "Eliminated Anomaly"],
    [
        ["STUDENT", "Student_ID", "None", "Student_Name", "Student name duplication across subjects"],
        ["SUBJECT", "Subject_ID", "None", "Subject_Name", "Subject name repeated per enrollment"],
        ["RESULT", "(Student_ID, Subject_ID)", "Student_ID → STUDENT,\nSubject_ID → SUBJECT", "Marks", "Pure intersection associative entity"]
    ],
    col_widths=[1.3, 1.6, 1.7, 1.3, 1.1]
)

# ===========================================================================
# Solution 4: Attribute Closure
# ===========================================================================
add_q_heading(doc, "Q4. Attribute Closure")
add_lead_para(doc, "Given", "R(A, B, C, D, E), F = {A → B, B → C, C → D, D → E}")
add_para(doc, "Step-by-Step Closure Derivation for A+:", bold=True, size=9.5, space_before=2, space_after=2)
add_grid_table(doc,
    ["Iteration (i)", "Current Set X^(i)", "FD Applied (W → Y)", "Attributes Added", "Updated Set X^(i+1)"],
    [
        ["0", "{A}", "Initial Basis", "—", "{A}"],
        ["1", "{A}", "A → B (since A ⊆ {A})", "{B}", "{A, B}"],
        ["2", "{A, B}", "B → C (since B ⊆ {A, B})", "{C}", "{A, B, C}"],
        ["3", "{A, B, C}", "C → D (since C ⊆ {A, B, C})", "{D}", "{A, B, C, D}"],
        ["4", "{A, B, C, D}", "D → E (since D ⊆ {A, B, C, D})", "{E}", "{A, B, C, D, E} = R"]
    ],
    col_widths=[1.1, 1.4, 1.8, 1.2, 1.5]
)
add_lead_para(doc, "Candidate Key Evaluation", "A+ = {A, B, C, D, E} = R. Because A determines all attributes of R and consists of a single attribute (no proper subset exists), A is a minimal Candidate Key.")

# ===========================================================================
# Solution 5: Identify the Normal Form & 2NF Decomposition
# ===========================================================================
add_q_heading(doc, "Q5. Identify the Normal Form")
add_lead_para(doc, "Given", "R(A, B, C, D), Candidate Key = AB, FDs: AB → C, A → D")
add_lead_para(doc, "1NF Evaluation", "R is in 1NF assuming atomic attribute domains.")
add_lead_para(doc, "2NF Evaluation", "In FD A → D, determinant A is a proper subset of candidate key AB (A ⊂ AB), and dependent D is non-prime. This is a partial dependency, so R is NOT in 2NF.")
add_para(doc, "2NF Relational Decomposition Table:", bold=True, size=9.5, space_before=2, space_after=2)
add_grid_table(doc,
    ["Relation", "Attributes", "Primary Key (PK)", "Foreign Key (FK)", "Satisfied Normal Form", "Resolved FD"],
    [
        ["R1", "A, D", "A", "None", "2NF / 3NF / BCNF", "Partial dependency A → D"],
        ["R2", "A, B, C", "AB", "A → R1(A)", "2NF / 3NF / BCNF", "Full key dependency AB → C"]
    ],
    col_widths=[0.8, 1.2, 1.3, 1.3, 1.3, 1.1]
)

# ===========================================================================
# Solution 6: Order Detail
# ===========================================================================
add_q_heading(doc, "Q6. Order Detail")
add_lead_para(doc, "Candidate Key", "(Order_ID, Product_ID). (Order_ID, Product_ID)+ = {Order_ID, Product_ID, Customer_Name, Product_Name, Product_Price, Quantity} = R.")
add_lead_para(doc, "Partial Dependencies", "Order_ID → Customer_Name (Order_ID ⊂ CK) and Product_ID → {Product_Name, Product_Price} (Product_ID ⊂ CK).")
add_para(doc, "Final 3NF Decomposition Table:", bold=True, size=9.5, space_before=2, space_after=2)
add_grid_table(doc,
    ["Relation Name", "Primary Key (PK)", "Foreign Key (FK)", "Non-Key Attributes", "Role"],
    [
        ["ORDER", "Order_ID", "None", "Customer_Name", "Order master entity"],
        ["PRODUCT", "Product_ID", "None", "Product_Name, Product_Price", "Product catalog entity"],
        ["ORDER_DETAIL", "(Order_ID, Product_ID)", "Order_ID → ORDER,\nProduct_ID → PRODUCT", "Quantity", "Composite transaction table"]
    ],
    col_widths=[1.5, 1.6, 1.7, 1.3, 0.9]
)

# ===========================================================================
# Solution 7: Hospital Appointment
# ===========================================================================
add_q_heading(doc, "Q7. Hospital Appointment")
add_lead_para(doc, "Candidate Key", "(Patient_ID, Doctor_ID)")
add_lead_para(doc, "Partial Dependencies", "Patient_ID → Patient_Name and Doctor_ID → {Doctor_Name, Specialization}. Both determinants are proper subsets of candidate key (Patient_ID, Doctor_ID).")
add_para(doc, "3NF Decomposition Table:", bold=True, size=9.5, space_before=2, space_after=2)
add_grid_table(doc,
    ["Relation Name", "Primary Key (PK)", "Foreign Key (FK)", "Non-Key Attributes", "Role"],
    [
        ["PATIENT", "Patient_ID", "None", "Patient_Name", "Master Patient Entity"],
        ["DOCTOR", "Doctor_ID", "None", "Doctor_Name, Specialization", "Master Doctor Entity"],
        ["APPOINTMENT", "(Patient_ID, Doctor_ID)", "Patient_ID → PATIENT,\nDoctor_ID → DOCTOR", "Appointment_Date", "Appointment Transaction"]
    ],
    col_widths=[1.4, 1.6, 1.7, 1.3, 1.0]
)

# ===========================================================================
# Solution 8: University Enrollment
# ===========================================================================
add_q_heading(doc, "Q8. University Enrollment")
add_lead_para(doc, "Candidate Key", "(Student_ID, Course_ID)")
add_lead_para(doc, "Dependencies", "Partial: Student_ID → Student_Name, Course_ID → {Course_Name, Instructor_ID}. Transitive: Course_ID → Instructor_ID → Instructor_Name.")
add_para(doc, "3NF Decomposition Table:", bold=True, size=9.5, space_before=2, space_after=2)
add_grid_table(doc,
    ["Relation Name", "Primary Key (PK)", "Foreign Key (FK)", "Non-Key Attributes", "Role"],
    [
        ["STUDENT", "Student_ID", "None", "Student_Name", "Student Master Entity"],
        ["INSTRUCTOR", "Instructor_ID", "None", "Instructor_Name", "Instructor Master Entity"],
        ["COURSE", "Course_ID", "Instructor_ID → INSTRUCTOR", "Course_Name, Instructor_ID", "Course Entity with Dept FK"],
        ["ENROLLMENT", "(Student_ID, Course_ID)", "Student_ID → STUDENT,\nCourse_ID → COURSE", "Grade", "Enrollment associative table"]
    ],
    col_widths=[1.4, 1.6, 1.7, 1.3, 1.0]
)

# ===========================================================================
# Solution 9: Find Candidate Keys
# ===========================================================================
add_q_heading(doc, "Q9. Find Candidate Keys")
add_lead_para(doc, "Essential Attributes Analysis", "RHS attributes in F = {B, C, E, F}. Attributes absent from the RHS of all FDs = {A, D}. Because no FD can generate A or D, attributes A and D must appear in EVERY candidate key of R.")
add_para(doc, "Iterative Closure Calculations Table:", bold=True, size=9.5, space_before=2, space_after=2)
add_grid_table(doc,
    ["Candidate Set (X)", "Step-by-Step Closure Derivation", "Closure Result (X+)", "Is Candidate Key?"],
    [
        ["{A}", "A → B, B → C", "{A, B, C}", "No (≠ R)"],
        ["{D}", "No FD triggers", "{D}", "No (≠ R)"],
        ["{A, D}", "AD^(0)={A,D} → AD^(1)={A,B,D} (via A→B) → AD^(2)={A,B,C,D} (via B→C) → AD^(3)={A,B,C,D,E} (via CD→E) → AD^(4)={A,B,C,D,E,F} = R (via E→F)", "{A, B, C, D, E, F} = R", "YES (Candidate Key)"]
    ],
    col_widths=[1.2, 3.2, 1.5, 1.1]
)
add_lead_para(doc, "Conclusion", "AD is the unique minimal Candidate Key of relation R.")

# ===========================================================================
# Solution 10: Library Loan
# ===========================================================================
add_q_heading(doc, "Q10. Library Loan")
add_lead_para(doc, "Candidate Key", "(Member_ID, Book_ID, Loan_Date)")
add_lead_para(doc, "Dependencies", "Partial: Member_ID → Member_Name; Book_ID → {Book_Title, Author_ID, Publisher_ID}. Transitive: Author_ID → Author_Name; Publisher_ID → Publisher_Name.")
add_para(doc, "3NF Decomposition Table (5 Relations):", bold=True, size=9.5, space_before=2, space_after=2)
add_grid_table(doc,
    ["Relation Name", "Primary Key (PK)", "Foreign Key (FK)", "Attributes / Role", "Resolved Dependency"],
    [
        ["MEMBER", "Member_ID", "None", "Member_Name", "Member_ID → Member_Name"],
        ["AUTHOR", "Author_ID", "None", "Author_Name", "Author_ID → Author_Name"],
        ["PUBLISHER", "Publisher_ID", "None", "Publisher_Name", "Publisher_ID → Publisher_Name"],
        ["BOOK", "Book_ID", "Author_ID → AUTHOR,\nPublisher_ID → PUBLISHER", "Book_Title, Author_ID, Publisher_ID", "Book_ID → Book_Title, Author_ID, Publisher_ID"],
        ["LOAN", "(Member_ID, Book_ID, Loan_Date)", "Member_ID → MEMBER,\nBook_ID → BOOK", "Return_Date", "Transaction checkout record"]
    ],
    col_widths=[1.2, 1.8, 1.6, 1.4, 1.0]
)

# ===========================================================================
# Solution 11: Sales Database
# ===========================================================================
add_q_heading(doc, "Q11. Sales Database")
add_lead_para(doc, "Candidate Key", "(Order_ID, Product_ID)")
add_lead_para(doc, "Decomposition Strategy", "Eliminating partial dependencies on (Order_ID, Product_ID) yields 2NF relations ORDER, PRODUCT, and ORDER_ITEM. Subsequently eliminating transitive dependencies (Customer_ID → Customer_Name, Salesperson_ID → Salesperson_Name, Category_ID → Category_Name) completes 3NF.")
add_para(doc, "Final 3NF Multi-Tier Schema Table:", bold=True, size=9.5, space_before=2, space_after=2)
add_grid_table(doc,
    ["Relation Name", "Primary Key (PK)", "Foreign Key (FK)", "Non-Key Attributes", "Addressed Level"],
    [
        ["CUSTOMER", "Customer_ID", "None", "Customer_Name", "Transitive (3NF)"],
        ["SALESPERSON", "Salesperson_ID", "None", "Salesperson_Name", "Transitive (3NF)"],
        ["CATEGORY", "Category_ID", "None", "Category_Name", "Transitive (3NF)"],
        ["PRODUCT", "Product_ID", "Category_ID → CATEGORY", "Product_Name, Unit_Price", "Partial (2NF) + FK"],
        ["ORDER", "Order_ID", "Customer_ID → CUSTOMER,\nSalesperson_ID → SALESPERSON", "Customer_ID, Salesperson_ID", "Partial (2NF) + FKs"],
        ["ORDER_ITEM", "(Order_ID, Product_ID)", "Order_ID → ORDER,\nProduct_ID → PRODUCT", "Quantity", "Full Key Composite"]
    ],
    col_widths=[1.3, 1.6, 1.7, 1.3, 1.1]
)

# ===========================================================================
# Solution 12: Course Registration
# ===========================================================================
add_q_heading(doc, "Q12. Course Registration")
add_lead_para(doc, "Candidate Key", "(Student_ID, Course_ID)")
add_lead_para(doc, "2NF Evaluation", "Not in 2NF because Student_ID → Student_Name and Course_ID → {Course_Name, Dept_ID} are partial dependencies.")
add_lead_para(doc, "3NF Evaluation", "Course_ID → Dept_ID → Dept_Name is a transitive dependency.")
add_para(doc, "3NF Relational Architecture Table:", bold=True, size=9.5, space_before=2, space_after=2)
add_grid_table(doc,
    ["Relation Name", "Primary Key (PK)", "Foreign Key (FK)", "Attributes", "Role"],
    [
        ["STUDENT", "Student_ID", "None", "Student_ID, Student_Name", "Student Master Entity"],
        ["DEPARTMENT", "Dept_ID", "None", "Dept_ID, Dept_Name", "Department Master Entity"],
        ["COURSE", "Course_ID", "Dept_ID → DEPARTMENT", "Course_ID, Course_Name, Dept_ID", "Course Entity with Dept FK"],
        ["REGISTRATION", "(Student_ID, Course_ID)", "Student_ID → STUDENT,\nCourse_ID → COURSE", "Student_ID, Course_ID, Grade", "Enrollment associative table"]
    ],
    col_widths=[1.4, 1.6, 1.7, 1.3, 1.0]
)

# ===========================================================================
# Solution 13: Employee–Project Assignment
# ===========================================================================
add_q_heading(doc, "Q13. Employee–Project Assignment")
add_lead_para(doc, "Candidate Key", "(Emp_ID, Project_ID)")
add_lead_para(doc, "Dependencies", "Partial: Emp_ID → {Emp_Name, Dept_ID}; Project_ID → {Project_Name, Manager_ID}. Transitive: Emp_ID → Dept_ID → Dept_Name; Project_ID → Manager_ID → Manager_Name.")
add_para(doc, "3NF Relational Schema Table (5 Relations):", bold=True, size=9.5, space_before=2, space_after=2)
add_grid_table(doc,
    ["Relation Name", "Primary Key (PK)", "Foreign Key (FK)", "Attributes", "Preserved Dependency"],
    [
        ["DEPARTMENT", "Dept_ID", "None", "Dept_ID, Dept_Name", "Dept_ID → Dept_Name"],
        ["EMPLOYEE", "Emp_ID", "Dept_ID → DEPARTMENT", "Emp_ID, Emp_Name, Dept_ID", "Emp_ID → Emp_Name, Dept_ID"],
        ["MANAGER", "Manager_ID", "None", "Manager_ID, Manager_Name", "Manager_ID → Manager_Name"],
        ["PROJECT", "Project_ID", "Manager_ID → MANAGER", "Project_ID, Project_Name, Manager_ID", "Project_ID → Project_Name, Manager_ID"],
        ["ASSIGNMENT", "(Emp_ID, Project_ID)", "Emp_ID → EMPLOYEE,\nProject_ID → PROJECT", "Emp_ID, Project_ID, Hours", "(Emp_ID, Project_ID) → Hours"]
    ],
    col_widths=[1.3, 1.5, 1.6, 1.5, 1.1]
)

# ===========================================================================
# Solution 14: University Timetable
# ===========================================================================
add_q_heading(doc, "Q14. University Timetable")
add_lead_para(doc, "Candidate Keys", "Enrollment key = (Student_ID, Course_ID); Class session schedule key = (Course_ID, Day, Time).")
add_lead_para(doc, "Dependencies", "Partial: Student_ID → Student_Name, Course_ID → {Course_Name, Instructor_ID, Department_ID}. Transitive: Instructor_ID → Instructor_Name, Department_ID → Department_Name, Room_ID → Room_Capacity.")
add_para(doc, "3NF Relational Schema Table (7 Relations):", bold=True, size=9.5, space_before=2, space_after=2)
add_grid_table(doc,
    ["Relation Name", "Primary Key (PK)", "Foreign Key (FK)", "Attributes"],
    [
        ["STUDENT", "Student_ID", "None", "Student_ID, Student_Name"],
        ["INSTRUCTOR", "Instructor_ID", "None", "Instructor_ID, Instructor_Name"],
        ["DEPARTMENT", "Department_ID", "None", "Department_ID, Department_Name"],
        ["ROOM", "Room_ID", "None", "Room_ID, Room_Capacity"],
        ["COURSE", "Course_ID", "Instructor_ID → INSTRUCTOR,\nDepartment_ID → DEPARTMENT", "Course_ID, Course_Name, Instructor_ID, Department_ID"],
        ["CLASS_SESSION", "(Course_ID, Day, Time)", "Course_ID → COURSE,\nRoom_ID → ROOM", "Course_ID, Day, Time, Room_ID"],
        ["ENROLLMENT", "(Student_ID, Course_ID)", "Student_ID → STUDENT,\nCourse_ID → COURSE", "Student_ID, Course_ID, Grade"]
    ],
    col_widths=[1.4, 1.8, 1.8, 2.0]
)

# ===========================================================================
# Solution 15: Complete 3NF Decomposition (Dual-Chain)
# ===========================================================================
add_q_heading(doc, "Q15. Complete 3NF Decomposition")
add_lead_para(doc, "Closure Computation", "A^(0)={A} → A^(1)={A,B,C} (via A→B, A→C) → A^(2)={A,B,C,D,E} (via B→D, C→E) → A^(3)={A,B,C,D,E,F,G} = R (via D→F, E→G). Since A+ = R and A is minimal, A is the unique Candidate Key.")
add_lead_para(doc, "Dual Transitive Chains", "Chain 1: A → B → D → F (violates 3NF at B→D and D→F). Chain 2: A → C → E → G (violates 3NF at C→E and E→G).")
add_para(doc, "3NF Decomposition Table (5 Relations):", bold=True, size=9.5, space_before=2, space_after=2)
add_grid_table(doc,
    ["Relation", "Primary Key", "Foreign Key (FK)", "Attributes", "Preserved Dependency"],
    [
        ["R1", "A", "None", "A, B, C", "A → B, A → C"],
        ["R2", "B", "B → R1(B)", "B, D", "B → D"],
        ["R3", "D", "D → R2(D)", "D, F", "D → F"],
        ["R4", "C", "C → R1(C)", "C, E", "C → E"],
        ["R5", "E", "E → R4(E)", "E, G", "E → G"]
    ],
    col_widths=[1.1, 1.2, 1.5, 1.5, 1.7]
)

# ===========================================================================
# Solution 16: Mixed Dependency Challenge
# ===========================================================================
add_q_heading(doc, "Q16. Mixed Dependency Challenge")
add_lead_para(doc, "Key & Classification", "Candidate Key = ABC. Partial: A → D (A ⊂ ABC), BC → E (BC ⊂ ABC). Transitive: E → F (since ABC → E and E → F where E is non-key).")
add_lead_para(doc, "2NF Intermediate", "R1(A, D), R2(B, C, E, F), R3(A, B, C).")
add_para(doc, "Final 3NF Decomposition Table:", bold=True, size=9.5, space_before=2, space_after=2)
add_grid_table(doc,
    ["Relation", "Primary Key", "Foreign Key (FK)", "Attributes", "Normal Form Status"],
    [
        ["R1", "A", "None", "A, D", "3NF (A → D resolved)"],
        ["R2", "BC", "None", "B, C, E", "3NF (BC → E resolved)"],
        ["R3", "E", "None", "E, F", "3NF (E → F resolved)"],
        ["R4", "ABC", "A → R1(A),\nBC → R2(BC)", "A, B, C", "3NF (Candidate Key relation)"]
    ],
    col_widths=[1.1, 1.3, 1.6, 1.5, 1.5]
)

# ===========================================================================
# Solution 17: Multiple Candidate Keys & Formal 3NF Verification
# ===========================================================================
add_q_heading(doc, "Q17. Multiple Candidate Keys")
add_para(doc, "Attribute Closures Table:", bold=True, size=9.5, space_before=2, space_after=2)
add_grid_table(doc,
    ["Attribute Set (X)", "Closure Steps", "Result (X+)", "Candidate Key Status"],
    [
        ["A", "A → B", "{A, B}", "No"],
        ["B", "B → A", "{A, B}", "No"],
        ["AC", "AC → B (via A→B), AC → D (via AC→D)", "{A, B, C, D} = R", "YES (Candidate Key 1)"],
        ["BD", "BD → A (via B→A), BD → C (via BD→C)", "{A, B, C, D} = R", "YES (Candidate Key 2)"]
    ],
    col_widths=[1.5, 2.5, 1.8, 1.2]
)
add_lead_para(doc, "Prime Attributes", "{A, C} ∪ {B, D} = {A, B, C, D}. Every single attribute in R is prime!")
add_para(doc, "Formal 3NF Verification Test Table:", bold=True, size=9.5, space_before=2, space_after=2)
add_grid_table(doc,
    ["Functional Dependency (X → Y)", "Is X Superkey?", "Is Y Prime?", "3NF Condition Satisfied?", "Justification"],
    [
        ["A → B", "No (A+ = {A,B})", "YES (B ∈ BD)", "SATISFIED", "Dependent B is prime"],
        ["B → A", "No (B+ = {A,B})", "YES (A ∈ AC)", "SATISFIED", "Dependent A is prime"],
        ["AC → D", "YES (AC+ = R)", "YES (D ∈ BD)", "SATISFIED", "Determinant AC is Superkey"],
        ["BD → C", "YES (BD+ = R)", "YES (C ∈ AC)", "SATISFIED", "Determinant BD is Superkey"]
    ],
    col_widths=[1.5, 1.4, 1.1, 1.4, 1.6]
)
add_lead_para(doc, "Conclusion", "Because every FD satisfies the formal 3NF condition (either determinant is superkey or dependent is prime), relation R is already in 3NF.")

# ===========================================================================
# Solution 18: Lossless and Dependency Preservation
# ===========================================================================
add_q_heading(doc, "Q18. Lossless and Dependency Preservation")
add_lead_para(doc, "Candidate Key", "A+ = {A, B, C, D} = R. Minimal candidate key is A.")
add_para(doc, "Dependency Preservation Analysis Table:", bold=True, size=9.5, space_before=2, space_after=2)
add_grid_table(doc,
    ["Original FD", "Sub-Relation", "Local FD Set", "Preservation Status"],
    [
        ["A → B", "R1(A, B)", "A → B ∈ F1", "Preserved directly in R1"],
        ["B → C", "R2(B, C)", "B → C ∈ F2", "Preserved directly in R2"],
        ["C → D", "R3(C, D)", "C → D ∈ F3", "Preserved directly in R3"]
    ],
    col_widths=[1.5, 1.5, 2.0, 2.0]
)
add_para(doc, "Lossless Join Proof Table (Binary Intersection Method):", bold=True, size=9.5, space_before=2, space_after=2)
add_grid_table(doc,
    ["Step", "Sub-relations Joined", "Common Attributes (Intersection)", "Functional Dependency Holding", "Superkey Condition Met?"],
    [
        ["Step 1", "R1(A,B) ⋈ R2(B,C)", "R1 ∩ R2 = {B}", "B → C in R2 (B+ = {B,C,D})", "YES: B is superkey of R2. R12(A,B,C) is lossless."],
        ["Step 2", "R12(A,B,C) ⋈ R3(C,D)", "R12 ∩ R3 = {C}", "C → D in R3 (C+ = {C,D})", "YES: C is superkey of R3. R(A,B,C,D) is lossless."]
    ],
    col_widths=[1.0, 1.6, 1.6, 1.5, 1.3]
)
add_lead_para(doc, "Conclusion", "The decomposition is both strictly dependency-preserving and guarantees a lossless join.")

# ===========================================================================
# Solution 19: 3NF Synthesis Challenge
# ===========================================================================
add_q_heading(doc, "Q19. 3NF Synthesis Challenge")
add_lead_para(doc, "Minimal Cover Derivation", "1. Split RHS: {A→B, A→C, B→D, CD→E, E→F, F→G, G→H}. 2. Check extraneous LHS in CD→E: C+={C}, D+={D}; neither derives E, so both C and D are required. 3. Check redundant FDs: none are redundant. F_min = {A→B, A→C, B→D, CD→E, E→F, F→G, G→H}. Candidate Key = A.")
add_para(doc, "3NF Synthesis Table (Bernstein's Algorithm):", bold=True, size=9.5, space_before=2, space_after=2)
add_grid_table(doc,
    ["Relation", "Determinant", "Attributes", "Primary Key", "Contains Candidate Key?", "Preserved FDs"],
    [
        ["R1", "A", "A, B, C", "A", "YES (Contains CK A)", "A → B, A → C"],
        ["R2", "B", "B, D", "B", "No", "B → D"],
        ["R3", "CD", "C, D, E", "(C, D)", "No", "CD → E"],
        ["R4", "E", "E, F", "E", "No", "E → F"],
        ["R5", "F", "F, G", "F", "No", "F → G"],
        ["R6", "G", "G, H", "G", "No", "G → H"]
    ],
    col_widths=[1.0, 1.1, 1.3, 1.0, 1.4, 1.2]
)
add_lead_para(doc, "Lossless Join Property", "Because relation R1 contains the candidate key A, the standard synthesis theorem guarantees that this decomposition is both lossless-join and dependency-preserving without requiring an additional relation.")

# ===========================================================================
# Solution 20: Full Real-World Normalization Challenge
# ===========================================================================
add_q_heading(doc, "Q20. Full Real-World Normalization Challenge")
add_lead_para(doc, "Candidate Key", "(Order_ID, Product_ID)")
add_lead_para(doc, "2NF Stage", "Extract ORDER(Order_ID, Customer_ID, Salesperson_ID), PRODUCT(Product_ID, Product_Name, Category_ID, Supplier_ID, Unit_Price), and ORDER_ITEM(Order_ID, Product_ID, Quantity).")
add_lead_para(doc, "3NF Stage", "Extract master relations CUSTOMER, SALESPERSON, CATEGORY, and SUPPLIER to eliminate transitive dependencies.")
add_para(doc, "Complete 3NF Decomposed Schema Table (7 Relations):", bold=True, size=9.5, space_before=2, space_after=2)
add_grid_table(doc,
    ["Relation Name", "Primary Key (PK)", "Foreign Key (FK)", "Attributes", "Eliminated Anomaly"],
    [
        ["CUSTOMER", "Customer_ID", "None", "Customer_ID, Customer_Name, Customer_City", "Client data update anomalies"],
        ["SALESPERSON", "Salesperson_ID", "None", "Salesperson_ID, Salesperson_Name", "Repeated staff data"],
        ["CATEGORY", "Category_ID", "None", "Category_ID, Category_Name", "Redundant category labels"],
        ["SUPPLIER", "Supplier_ID", "None", "Supplier_ID, Supplier_Name", "Vendor master repetition"],
        ["ORDER", "Order_ID", "Customer_ID, Salesperson_ID", "Order_ID, Customer_ID, Salesperson_ID", "Partial dependency on Order_ID"],
        ["PRODUCT", "Product_ID", "Category_ID, Supplier_ID", "Product_ID, Product_Name, Category_ID, Supplier_ID, Unit_Price", "Partial dependency on Product_ID"],
        ["ORDER_ITEM", "(Order_ID, Product_ID)", "Order_ID, Product_ID", "Order_ID, Product_ID, Quantity", "Full composite key association"]
    ],
    col_widths=[1.3, 1.5, 1.6, 1.5, 1.1]
)
add_para(doc, "Operational Redundancy & Anomaly Prevention Analysis:", bold=True, size=9.5, space_before=2, space_after=2)
add_grid_table(doc,
    ["Anomaly Type", "Manifestation in Unnormalized Schema", "Resolution in 3NF Schema"],
    [
        ["Update Anomaly", 
         "If a supplier or customer moves or updates their name, every single order row containing that entity must be scanned and updated. Missing any row leads to data inconsistency.", 
         "Customer and Supplier data exist in exactly one place (CUSTOMER and SUPPLIER tables). An update requires modifying exactly one tuple."],
        ["Insertion Anomaly", 
         "A new product cannot be added until an order is placed for it, because the composite primary key requires Order_ID. Likewise, a new customer cannot be registered without an order.", 
         "Entities can be independently inserted into CUSTOMER, PRODUCT, SUPPLIER, or SALESPERSON tables without requiring dummy order records."],
        ["Deletion Anomaly", 
         "If an order record is cancelled or deleted, deleting that row inadvertently erases the associated customer, salesperson, or product definitions from the system.", 
         "Deleting an order or order item record leaves master entities in CUSTOMER, PRODUCT, and SUPPLIER completely intact."]
    ],
    col_widths=[1.3, 2.8, 2.9]
)

# ===========================================================================
# Quick Revision — Normalization Checklist (Matching format.docx style)
# ===========================================================================
add_q_heading(doc, "Quick Revision — Normalization Checklist")
checklist_items = [
    "Check atomicity/repeating groups → 1NF.",
    "Find candidate key(s), especially composite keys.",
    "Find partial dependencies → 2NF.",
    "Find transitive dependencies → 3NF.",
    "Mark PKs and FKs in every final relation.",
    "When required, verify lossless join and dependency preservation."
]

for idx, item in enumerate(checklist_items):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    p.paragraph_format.left_indent = Inches(0.25)
    r_num = p.add_run(f"{idx + 1}. ")
    r_num.font.name = "Arial"
    r_num.font.bold = True
    r_num.font.size = Pt(10)
    r_num.font.color.rgb = RGBColor(0, 0, 0)

    r_txt = p.add_run(item)
    r_txt.font.name = "Arial"
    r_txt.font.size = Pt(10)
    r_txt.font.color.rgb = RGBColor(0, 0, 0)

# ---------------------------------------------------------------------------
# Save Document
# ---------------------------------------------------------------------------
output_filename = "DBMS_Normalization_Worksheet_and_Solutions.docx"
doc.save(output_filename)
print(f"Successfully generated: {output_filename}")

# Also update format.docx so both files reflect the exact required format
doc.save("format.docx")
print("Successfully updated: format.docx")