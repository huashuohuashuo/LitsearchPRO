# LitSearchPro

LitSearchPro is an integrated research, laboratory management, and academic workflow platform designed for university laboratories, research groups, and teaching teams.

The software combines literature search, document management, research project organization, collaborative workflows, laboratory equipment borrowing, laboratory reservation, safety approval, hazardous chemical inventory management, and lightweight web access into one unified system.

## Overview

LitSearchPro was originally developed to support daily research and laboratory administration in an academic environment. It aims to reduce repetitive administrative work, improve the traceability of laboratory approvals, and provide students, mentors, laboratory administrators, and system administrators with a consistent digital workflow.

The system includes both a desktop client and a lightweight web portal. The desktop client provides a full-featured interface for research and administrative work, while the web portal allows users to complete essential actions such as account registration, approvals, equipment borrowing, laboratory reservation, chemical use applications, and signature upload from a browser.

## Key Features

### Research and Literature Workflow

- Literature search and academic resource access
- Local literature library management
- PDF reading and document organization
- Research project and experiment record management
- Academic writing workflow support
- Systematic review and evidence organization tools
- Research subscription, tracking, and reporting features

### Account and Role Management

- User registration and approval workflow
- Role-based permission control
- Student, mentor, laboratory administrator, chemical warehouse administrator, and super administrator roles
- Electronic signature upload and management
- Real-name display for users while keeping account identifiers unique

### Laboratory Equipment Borrowing

- Equipment registration and import
- Equipment borrowing application
- Approval and rejection workflows
- Return application and status tracking
- Cross-team equipment borrowing support
- Borrowing records and approval logs
- Web-based equipment management and CSV import support

### Laboratory Reservation and Safety Management

- Public and team laboratory management
- Laboratory reservation application
- Multi-participant reservation confirmation
- Mentor approval and laboratory administrator approval workflow
- Safety commitment confirmation
- Upload of experiment procedure, chemical usage, and safety plan spreadsheets
- Reservation PDF generation with electronic signatures
- Laboratory access status tracking
- Experiment completion report and administrator confirmation
- Laboratory blacklist and reservation permission management
- Laboratory announcement and reservation channel control

### Hazardous Chemical Management

- Chemical warehouse creation and administrator assignment
- Chemical purchase registration and stock-in approval
- Chemical withdrawal application
- Co-user confirmation workflow
- Mentor and warehouse administrator approval workflow
- Chemical inventory tracking
- Chemical disposal report and approval
- PDF application form generation with electronic signatures
- Chemical stock summary by warehouse, chemical type, mentor, and team
- Time-limited authorization code mechanism for controlled chemical access

### Lightweight Web Portal

- Browser-based access for common workflows
- Account registration and approval
- Equipment borrowing, approval, and import
- Laboratory reservation and approval
- Hazardous chemical application and approval
- Electronic signature upload
- PDF download support
- Workflow experience aligned with the desktop client

## System Design

LitSearchPro follows a client-server architecture. The desktop client is implemented as a Python-based graphical application, while the server handles shared data, collaboration, approval workflows, web access, and persistent records.

The system focuses on:

- Data traceability
- Permission isolation
- Approval accountability
- Local and server-side PDF record management
- Practical usability in laboratory and university environments
- Extensibility for future institutional customization

## Target Users

This project is suitable for:

- University laboratories
- Research groups
- Teaching laboratories
- Laboratory safety management teams
- Academic departments
- Research administrators
- Student research teams
- Developers interested in academic workflow systems

## Open Source Purpose

This project is released as open source to encourage collaboration, transparency, and further development of digital research and laboratory management tools.

Developers may use this codebase to:

- Study the implementation of academic workflow software
- Build customized laboratory management systems
- Extend the approval and safety management modules
- Integrate local institutional requirements
- Improve usability, security, and scalability
- Contribute new features and bug fixes

## Notes

Before deploying this software in a production environment, please review and adapt the permission model, data storage strategy, backup policy, security configuration, and institutional compliance requirements according to your local regulations.

This project is intended as a practical research and laboratory management framework. Institutions and developers are encouraged to customize it responsibly for their own operational needs.

## License

Please refer to the license file included in this repository.

## Contribution

Contributions are welcome. You may submit issues, bug reports, feature suggestions, documentation improvements, or pull requests.

Recommended contribution areas include:

- User interface improvements
- Web portal optimization
- Database migration support
- Security hardening
- Internationalization
- Automated testing
- Deployment documentation
- Workflow customization
- PDF template refinement
