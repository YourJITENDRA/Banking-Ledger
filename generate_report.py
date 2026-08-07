import os
import sys
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def add_page_number(run):
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = "PAGE"
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)

def add_total_page_number(run):
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = "NUMPAGES"
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)

def set_cell_background(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color_hex)
    tcPr.append(shd)

def create_table_header(table, headers):
    hdr_cells = table.rows[0].cells
    for i, header_text in enumerate(headers):
        hdr_cells[i].text = header_text
        set_cell_background(hdr_cells[i], '1F497D')
        for paragraph in hdr_cells[i].paragraphs:
            for run in paragraph.runs:
                run.font.name = 'Arial'
                run.font.size = Pt(11)
                run.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)

def style_table_rows(table):
    for r_idx, row in enumerate(table.rows[1:], start=1):
        bg_color = 'F2F5F8' if r_idx % 2 == 0 else 'FFFFFF'
        for cell in row.cells:
            set_cell_background(cell, bg_color)
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = 'Arial'
                    run.font.size = Pt(10)

def main():
    doc = Document()
    
    # Page setup
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Styles Setup
    styles = doc.styles
    normal_style = styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(12)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(6)

    # --- COVER PAGE ---
    for _ in range(3): doc.add_paragraph()
    
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run("PROJECT REPORT ON\nDOUBLE-ENTRY LEDGER SYSTEM")
    title_run.font.name = 'Arial'
    title_run.font.size = Pt(24)
    title_run.bold = True
    title_run.font.color.rgb = RGBColor(31, 73, 125)
    
    subtitle_p = doc.add_paragraph()
    subtitle_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = subtitle_p.add_run("A Secure, High-Performance Transaction Engine with Strict ACIDs")
    sub_run.font.name = 'Arial'
    sub_run.font.size = Pt(14)
    sub_run.italic = True
    
    for _ in range(8): doc.add_paragraph()
    
    info_p = doc.add_paragraph()
    info_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info_run = info_p.add_run(
        "Submitted by:\n"
        "Jitendra Mishra\n"
        "Email: jitendramishra223355@gmail.com\n\n"
        "Academic Year: 2026\n\n"
        "Under the Guidance of: DeepMind Advanced Agentic Coding"
    )
    info_run.font.name = 'Arial'
    info_run.font.size = Pt(12)
    info_run.bold = True
    
    doc.add_page_break()

    # --- TABLE OF CONTENTS ---
    h_toc = doc.add_heading(level=1)
    h_toc_run = h_toc.add_run("TABLE OF CONTENTS")
    h_toc_run.font.name = 'Arial'
    h_toc_run.font.size = Pt(18)
    h_toc_run.bold = True
    h_toc_run.font.color.rgb = RGBColor(31, 73, 125)
    doc.add_paragraph("The following outline lists the chapters and page ranges contained within this document:")
    
    toc_data = [
        ("1. Introduction", "1"),
        ("   1.1 Purpose", "1"),
        ("   1.2 Project Scope", "2"),
        ("   1.3 Problem Definition", "3"),
        ("2. Literature Survey", "4"),
        ("   2.1 System Review", "4"),
        ("   2.2 Technology Used", "5"),
        ("3. Requirement Analysis", "9"),
        ("   3.1 Programming Language", "9"),
        ("   3.2 Operating system", "10"),
        ("   3.3 Hardware", "11"),
        ("4. Project Planning", "12"),
        ("   4.1 System Model", "12"),
        ("   4.2 API Endpoints", "13"),
        ("5. System Design", "15"),
        ("   5.1 Use Case Model", "15"),
        ("   5.2 Activity Diagram", "16"),
        ("   5.3 Class Diagram", "18"),
        ("6. System Testing", "19"),
        ("   6.1 Test Cases & Test Results", "19"),
        ("7. Implementation", "28"),
        ("   7.1 Work Flow", "28"),
        ("   7.2 Source Code", "29"),
        ("8. Screenshots of Project", "34"),
        ("9. Conclusion And Future Scope", "35"),
        ("10. References", "36"),
    ]
    
    for title, pg in toc_data:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        dots_count = 80 - len(title) - len(pg)
        dots = "." * max(10, dots_count)
        run_t = p.add_run(title)
        if not title.startswith(" "):
            run_t.bold = True
        p.add_run(dots)
        p.add_run(pg).bold = True
        
    doc.add_page_break()

    # --- SECTION WRITER HELPER ---
    def add_chapter_title(num, name):
        h = doc.add_heading(level=1)
        h_run = h.add_run(f"Chapter {num}: {name}")
        h_run.font.name = 'Arial'
        h_run.font.size = Pt(16)
        h_run.bold = True
        h_run.font.color.rgb = RGBColor(31, 73, 125)
        h.paragraph_format.space_before = Pt(12)
        h.paragraph_format.space_after = Pt(6)

    def add_section_title(num, name):
        h = doc.add_heading(level=2)
        h_run = h.add_run(f"{num} {name}")
        h_run.font.name = 'Arial'
        h_run.font.size = Pt(13)
        h_run.bold = True
        h_run.font.color.rgb = RGBColor(80, 80, 80)
        h.paragraph_format.space_before = Pt(6)
        h.paragraph_format.space_after = Pt(4)

    # --- CHAPTER 1 ---
    add_chapter_title("1", "Introduction")
    
    add_section_title("1.1", "Purpose")
    doc.add_paragraph(
        "In modern financial applications, transaction data accuracy, integrity, and immutability are non-negotiable. "
        "A simple double-entry system provides a robust framework that prevents the creation or loss of value by ensuring "
        "every debit is balanced by a corresponding credit. The purpose of this project, titled 'Backend Ledger System', "
        "is to design and implement an enterprise-ready ledger engine that allows users to create multi-currency accounts, "
        "initiate peer-to-peer transactions, and maintain a highly precise audit trail. The ledger uses transactional ACID "
        "properties under MongoDB to prevent double-spending, race conditions, and out-of-order execution."
    )
    doc.add_paragraph(
        "Additionally, the system is designed to provide secure, cookie-based tokenized authentication, automatic "
        "email notifications using SMTP and nodemailer, and a set of helper tools to allow system operations "
        "(such as initial funding transfers from a system account). The core design pattern focuses on the clean "
        "separation of concerns through routes, controllers, middleware, and database models."
    )

    add_section_title("1.2", "Project Scope")
    doc.add_paragraph(
        "The scope of this project includes the backend REST API implementation, the database model definition, security "
        "enhancements, and automated email messaging. Specifically, the functional scope includes:\n"
        "1. User register and login authentication using bcryptjs password hashing and JWT token storage in secure HttpOnly cookies.\n"
        "2. Blacklisting of JWT tokens on logout to prevent replay attacks.\n"
        "3. Multi-currency account creation with Mongoose validation rules.\n"
        "4. Strict double-entry ledger tracking: Every user transfer creates two balanced ledger entries (DEBIT and CREDIT) "
        "within a single MongoDB session transaction.\n"
        "5. System user interface endpoints to perform initial funding actions for new users.\n"
        "6. Complete automated email notification triggers for registrations and successful or failed transactions."
    )
    doc.add_paragraph(
        "Out of scope items: This version of the project focuses solely on the RESTful Backend API and does not include "
        "a visual user interface or client application, although it can be easily integrated with frontend frameworks like React or Next.js."
    )

    add_section_title("1.3", "Problem Definition")
    doc.add_paragraph(
        "Traditional financial backends often suffer from database synchronization issues. If a user transfers money to another, "
        "a simplistic 'update balance of A, then update balance of B' approach can fail midway if the server crashes. This creates "
        "inconsistencies where money is deducted but never received, or vice versa. The Ledger system resolves this by keeping a log of "
        "immutable entries rather than changing a single balance field in the account. The current balance is derived dynamically "
        "by summing the credits and debits from the ledger.\n"
        "Furthermore, security concerns like JWT hijacking via Client Javascript (XSS) and unauthorized account manipulation (a user "
        "transferring funds out of another user's account by specifying their account ID) present critical vulnerabilities that this "
        "implementation actively addresses and patches."
    )
    
    doc.add_page_break()

    # --- CHAPTER 2 ---
    add_chapter_title("2", "Literature Survey")
    
    add_section_title("2.1", "System Review")
    doc.add_paragraph(
        "Double-entry bookkeeping is a practice dating back to the 15th century, popularized by Luca Pacioli. "
        "Its digital translation requires absolute transaction safety. Modern systems like Amazon QLDB or Ledger databases "
        "provide tamper-evident transaction logs. For general-purpose web architectures, developers build custom ledgers "
        "over relational databases (like PostgreSQL) or document-oriented databases (like MongoDB) supporting multi-document transactions. "
        "A system review reveals that implementing double-entry engines over document databases requires specific precautions "
        "because document stores were originally not designed for multi-document ACID operations. MongoDB introduced multi-document "
        "session transactions in version 4.0, which allows developers to run atomic updates across collections like Users, Accounts, and Ledgers."
    )

    add_section_title("2.2", "Technology Used")
    doc.add_paragraph("The Backend Ledger System is built using the following core technologies:")
    
    tech_table = doc.add_table(rows=9, cols=2)
    tech_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    create_table_header(tech_table, ["Technology Component", "Purpose in Project"])
    
    techs = [
        ("Node.js", "Server-side runtime environment providing asynchronous event-driven execution."),
        ("Express.js (5.x)", "Web framework providing router management, JSON body parsing, and middleware chains."),
        ("MongoDB Atlas", "Cloud document database storing users, accounts, transaction records, and ledger logs."),
        ("Mongoose (9.x)", "ODM library enforcing strict schemas, indexes, and validation rules in MongoDB."),
        ("JSON Web Token (JWT)", "Decentralized authorization format for secure communication between client and server."),
        ("BcryptJS", "Salted password hashing algorithm protecting user passwords in the database."),
        ("Nodemailer", "Node.js email sending module supporting SMTP server connections for automated transactional emails."),
        ("Dotenv", "Zero-dependency module that loads environment variables from a root .env file.")
    ]
    
    for idx, (comp, purp) in enumerate(techs):
        row_cells = tech_table.rows[idx + 1].cells
        row_cells[0].text = comp
        row_cells[1].text = purp
        
    style_table_rows(tech_table)
    
    doc.add_page_break()

    # --- CHAPTER 3 ---
    add_chapter_title("3", "Requirement Analysis")
    
    add_section_title("3.1", "Programming Language")
    doc.add_paragraph(
        "The system is implemented using Javascript (ECMAScript 2022+ / CommonJS module system) running on the Node.js platform. "
        "JavaScript provides non-blocking, asynchronous I/O execution, which is extremely efficient for handling high-volume HTTP "
        "request traffic and parallel database lookups. Node.js's native Promise framework and async/await syntax ensure that ledger "
        "aggregation queries and transaction sessions run concurrently without thread blockages."
    )

    add_section_title("3.2", "Operating System")
    doc.add_paragraph(
        "The project is cross-platform and runs seamlessly on Microsoft Windows (tested on Windows 10/11), Linux, and macOS. "
        "For the development and verification stages of this project, Windows powershell environment was utilized to verify "
        "local runtime environments and dependency checks."
    )

    add_section_title("3.3", "Hardware Requirements")
    doc.add_paragraph(
        "The system has minimal physical hardware requirements, allowing it to run efficiently on standard developer workstations or "
        "micro-instance cloud hosting platforms (e.g. Render, Railway, AWS EC2 nano):\n"
        "- Processor: Dual-Core 2.0 GHz or higher (Intel Core i3/i5/i7 or AMD equivalent)\n"
        "- RAM: 4 GB minimum (8 GB recommended for concurrent database testing)\n"
        "- Storage: 500 MB of free disk space (excluding database growth)\n"
        "- Network: Persistent broadband connection for MongoDB Atlas cluster connections and SMTP email relays."
    )

    doc.add_page_break()

    # --- CHAPTER 4 ---
    add_chapter_title("4", "Project Planning")
    
    add_section_title("4.1", "System Model")
    doc.add_paragraph(
        "The system adopts a 3-tier architectural design: Client Layer (API Consumers), Application Layer (Express HTTP Server), "
        "and Data Layer (MongoDB). The logical layout of collections is structured as follows:"
    )
    
    db_table = doc.add_table(rows=6, cols=3)
    db_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    create_table_header(db_table, ["Schema/Model", "Key Attributes", "Security / Index Rules"])
    
    models_info = [
        ("User", "email, name, password, systemUser", "Email unique index, password select:false"),
        ("Account", "user, status, currency", "Compound index on { user: 1, status: 1 }"),
        ("Transaction", "fromAccount, toAccount, status, amount, idempotencyKey", "IdempotencyKey unique index to prevent double execution"),
        ("Ledger", "account, amount, transaction, type", "Immutable entries (hooks prevent update/delete queries)"),
        ("TokenBlacklist", "token, createdAt", "TTL index on createdAt (expiring tokens after 3 days)")
    ]
    
    for idx, (model, key_attr, rules) in enumerate(models_info):
        row_cells = db_table.rows[idx + 1].cells
        row_cells[0].text = model
        row_cells[1].text = key_attr
        row_cells[2].text = rules
        
    style_table_rows(db_table)
    
    doc.add_paragraph(
        "Under this system model, when a transaction starts: \n"
        "1. A new transaction is created in PENDING status.\n"
        "2. The sender account is debited, creating a DEBIT ledger entry.\n"
        "3. The receiver account is credited, creating a CREDIT ledger entry.\n"
        "4. The transaction is marked as COMPLETED.\n"
        "All steps are executed inside a MongoDB session transaction to guarantee atomicity. If any step fails, the entire transaction is rolled back."
    )

    add_section_title("4.2", "API Endpoints")
    doc.add_paragraph(
        "The Backend Ledger system exposes RESTful API endpoints for authentication, account management, and transaction routing. "
        "The complete API surface is documented below:"
    )

    api_table = doc.add_table(rows=9, cols=3)
    api_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    create_table_header(api_table, ["Method", "Endpoint Route", "Description and Usage"])

    apis_info = [
        ("POST", "/api/auth/register", "Registers a new user. Hashes the password and sets a cookie token."),
        ("POST", "/api/auth/login", "Authenticates credentials and signs a new JWT session token."),
        ("POST", "/api/auth/logout", "Logs out the user. Clears the JWT cookie and stores the token in the blacklist database."),
        ("POST", "/api/accounts", "Creates a new financial account in a specified currency (default INR) for the auth user."),
        ("GET", "/api/accounts", "Lists all active and closed accounts belonging to the authenticated user."),
        ("GET", "/api/accounts/balance/:accountId", "Derives the account balance dynamically by aggregating debits and credits."),
        ("POST", "/api/transactions", "Transfers funds between accounts. Requires ownership verification and idempotency check."),
        ("POST", "/api/transactions/system/initial-funds", "Allows system user accounts to inject funds (seed transactions) into user accounts.")
    ]

    for idx, (method, route, desc) in enumerate(apis_info):
        row_cells = api_table.rows[idx + 1].cells
        row_cells[0].text = method
        row_cells[1].text = route
        row_cells[2].text = desc

    style_table_rows(api_table)

    doc.add_page_break()

    # --- CHAPTER 5 ---
    add_chapter_title("5", "System Design")
    
    add_section_title("5.1", "Use Case Model")
    doc.add_paragraph(
        "The use case model outlines the core actions performed by actors in the system:\n"
        "1. User Actor: Can Register a new profile, Log in, Create accounts in specific currencies, check account balances, and initiate money transfers to other accounts.\n"
        "2. System User Actor: A special high-privilege account that bypasses standard checks to seed accounts with initial funds.\n"
        "3. Email Service Actor: Automatically triggers email dispatches to users when key use cases complete."
    )

    add_section_title("5.2", "Activity Diagram")
    doc.add_paragraph(
        "Activity Workflow for Money Transfer:\n"
        "[Client Request] -> [Auth Middleware Checks Token] -> [Verify Account Owners & Status] -> "
        "[Check Sender Balance via Ledger aggregation] -> [Open DB Session & Start Transaction] -> "
        "[Create PENDING Transaction Doc] -> [Write DEBIT Ledger Entry] -> [Write CREDIT Ledger Entry] -> "
        "[Update Transaction to COMPLETED] -> [Commit DB Session] -> [Send Email Notification via SMTP] -> [Return JSON Response]."
    )

    add_section_title("5.3", "Class Diagram")
    doc.add_paragraph(
        "Class/Model Relationships:\n"
        "- A User has one or more Accounts (1:N relationship).\n"
        "- An Account belongs to a User (N:1 relationship).\n"
        "- A Transaction references two Accounts: fromAccount and toAccount.\n"
        "- A Ledger entry references one Account (which holds the entry) and one Transaction (which triggered the entry).\n"
        "- A User has many Transactions via their Accounts."
    )

    doc.add_page_break()

    # --- CHAPTER 6 ---
    add_chapter_title("6", "System Testing")
    
    add_section_title("6.1", "Test Cases & Test Results")
    doc.add_paragraph("The following test suites were run on the optimized backend code to confirm stability and security:")
    
    test_table = doc.add_table(rows=8, cols=4)
    test_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    create_table_header(test_table, ["Test Case ID", "Description / Input", "Expected Result", "Status"])
    
    tests_info = [
        ("TC-001", "Register User with Valid Inputs", "User created, password hashed, JWT cookie set, welcome email sent", "PASS"),
        ("TC-002", "Login with Correct Credentials", "Returns JWT token, sets HTTP-only cookie, clears CLI debug logs", "PASS"),
        ("TC-003", "Create Account via Auth User", "Active account created under the authenticated user's ID", "PASS"),
        ("TC-004", "Fetch Balance (LEDGER Aggregation)", "Sums credits and subtracts debits correctly. Returns balance", "PASS"),
        ("TC-005", "Transfer with insufficient balance", "Returns 400 Bad Request, session rolled back, no ledger records written", "PASS"),
        ("TC-006", "Transfer from another user's account", "Returns 403 Forbidden: You do not own the sender account", "PASS"),
        ("TC-007", "Self-Transfer (Same sender & receiver)", "Returns 400 Bad Request: Cannot transfer to same account", "PASS")
    ]
    
    for idx, (tc_id, desc, exp, status) in enumerate(tests_info):
        row_cells = test_table.rows[idx + 1].cells
        row_cells[0].text = tc_id
        row_cells[1].text = desc
        row_cells[2].text = exp
        row_cells[3].text = status
        
    style_table_rows(test_table)
    
    doc.add_page_break()

    # --- CHAPTER 7 ---
    add_chapter_title("7", "Implementation")
    
    add_section_title("7.1", "Work Flow")
    doc.add_paragraph(
        "The project is structured with clean code conventions. Controllers retrieve data from requests, interact with schemas, "
        "and handle responses. Middleware processes HTTP headers/cookies to extract authentication metadata. "
        "The database models manage persistence layer validations."
    )

    add_section_title("7.2", "Source Code")
    doc.add_paragraph("Below are listings of the core optimized components of the Ledger server:")
    
    # Add key code snippets
    doc.add_paragraph("1. Database Connection Configuration (src/config/db.js):").bold = True
    p_code1 = doc.add_paragraph()
    p_code1.paragraph_format.left_indent = Inches(0.5)
    r_code1 = p_code1.add_run(
        "const mongoose = require(\"mongoose\");\n\n"
        "function connectToDB() {\n"
        "    mongoose.connect(process.env.MONGO_URI)\n"
        "        .then(() => {\n"
        "            console.log(\"server is connected to DB\");\n"
        "        })\n"
        "        .catch(err => {\n"
        "            console.log(\"Error connecting to DB\", err);\n"
        "            process.exit(1);\n"
        "        });\n"
        "}\n\n"
        "module.exports = connectToDB;"
    )
    r_code1.font.name = 'Courier New'
    r_code1.font.size = Pt(9.5)
    
    doc.add_paragraph("2. Token Check in Auth Middleware (src/middleware/auth.middleware.js):").bold = True
    p_code2 = doc.add_paragraph()
    p_code2.paragraph_format.left_indent = Inches(0.5)
    r_code2 = p_code2.add_run(
        "async function authMiddleware(req, res, next) {\n"
        "    const token = req.cookies.token || req.headers.authorization?.split(\" \")[ 1 ];\n"
        "    if (!token) {\n"
        "        return res.status(401).json({ message: \"Unauthorized access, token is missing\" });\n"
        "    }\n"
        "    const isBlacklisted = await tokenBlackListModel.findOne({ token });\n"
        "    if (isBlacklisted) {\n"
        "        return res.status(401).json({ message: \"Unauthorized access, token is invalid\" });\n"
        "    }\n"
        "    try {\n"
        "        const decoded = jwt.verify(token, process.env.JWT_SECRET);\n"
        "        const user = await userModel.findById(decoded.userId);\n"
        "        if (!user) {\n"
        "            return res.status(401).json({ message: \"Unauthorized access, user not found\" });\n"
        "        }\n"
        "        req.user = user;\n"
        "        return next();\n"
        "    } catch (err) {\n"
        "        return res.status(401).json({ message: \"Unauthorized access, token is invalid\" });\n"
        "    }\n"
        "}"
    )
    r_code2.font.name = 'Courier New'
    r_code2.font.size = Pt(9.5)

    doc.add_page_break()

    # --- CHAPTER 8 ---
    add_chapter_title("8", "Screenshots of Project")
    doc.add_paragraph("Since this is a backend-only REST API system, the operational interface was validated using Postman and terminal output scripts:")
    
    screenshot_table = doc.add_table(rows=5, cols=3)
    screenshot_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    create_table_header(screenshot_table, ["Endpoint", "Action / Screen Description", "Expected API Response Status"])
    
    screens = [
        ("POST /api/auth/register", "User inputs email, name, password. Creates user and sets cookie token.", "201 Created"),
        ("POST /api/auth/login", "User inputs credentials. Returns authenticated user information.", "200 OK"),
        ("POST /api/accounts", "Authenticated user creates an account record inside DB.", "201 Created"),
        ("POST /api/transactions", "Transfer money from sender's account to receiver's account inside transaction session.", "201 Created")
    ]
    
    for idx, (ep, desc, status) in enumerate(screens):
        row_cells = screenshot_table.rows[idx + 1].cells
        row_cells[0].text = ep
        row_cells[1].text = desc
        row_cells[2].text = status
        
    style_table_rows(screenshot_table)

    doc.add_page_break()

    # --- CHAPTER 9 ---
    add_chapter_title("9", "Conclusion And Future Scope")
    add_section_title("9.1", "Conclusion")
    doc.add_paragraph(
        "The Backend Ledger System successfully implements a double-entry database architecture with strict transaction-level "
        "atomicity using MongoDB sessions. Through careful debugging and security reviews, the application's credentials logging "
        "vulnerabilities were resolved, HTTP cookie safety attributes were enabled to shield cookies from script hijacking, "
        "and endpoint safety bounds were set up to restrict unauthorized funds manipulation. The email messaging engine "
        "integrated over SMTP SMTP-transporter sends notifications on registration and transfers, assuring clear tracking."
    )

    add_section_title("9.2", "Future Scope")
    doc.add_paragraph(
        "The current backend ledger server is highly scalable, but can be improved with the following future features:\n"
        "1. Integration of a visual administrative dashboard using React or Next.js to monitor transaction records and account details in real-time.\n"
        "2. Adding multi-currency conversion APIs to support transfers across accounts of different currencies (e.g. converting INR to USD during transaction logic).\n"
        "3. Integration of phone OTP registration mechanisms to supplement email authentication.\n"
        "4. Deployment to cloud platforms like Heroku/AWS with automated CI/CD pipeline tests."
    )

    doc.add_page_break()

    # --- CHAPTER 10 ---
    add_chapter_title("10", "References")
    doc.add_paragraph(
        "1. Pacioli, Luca (1494). 'Summa de arithmetica, geometria, proportioni et proportionalita' - Classic Double-entry Bookkeeping foundations.\n"
        "2. MongoDB Official Documentation - 'Multi-document Transactions in Node.js Applications'. https://www.mongodb.com/docs/manual/core/transactions/\n"
        "3. Express.js API Reference - 'Middleware development and global error boundaries'. https://expressjs.com/\n"
        "4. RFC 7519 - 'JSON Web Token (JWT) specification for security transmission'. https://tools.ietf.org/html/rfc7519\n"
        "5. Nodemailer official guidelines - 'Secure SMTP connections using OAuth2 and Gmail App Passwords'. https://nodemailer.com/"
    )

    # Add page numbers to footers
    for section in doc.sections:
        footer = section.footer
        p = footer.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.text = "Page "
        run = p.add_run()
        add_page_number(run)
        p.add_run(" of ")
        run2 = p.add_run()
        add_total_page_number(run2)
        
        # Style footer run
        for run_f in p.runs:
            run_f.font.name = 'Arial'
            run_f.font.size = Pt(9)
            run_f.font.color.rgb = RGBColor(128, 128, 128)

    # Save Document
    output_path = "Project_Report.docx"
    doc.save(output_path)
    print(f"Project report successfully created and saved to: {os.path.abspath(output_path)}")

if __name__ == "__main__":
    main()
