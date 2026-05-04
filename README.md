# Production Document Generator (FC, FMEA & CP)

## 1. Overview
The Production Document Generator is a Windows-based professional utility designed to automate the creation of industrial manufacturing documentation. By processing a user-defined sequence of production phases, the software populates standardized Excel templates to generate interconnected Flow Charts (FC), Failure Mode and Effects Analysis (FMEA) reports, and Control Plans (CP).

## 2. Core Capabilities
* **Centralized Management:** A dedicated graphical interface for the configuration, sequencing, and modification of production routings.
* **Automated Resource Architecture:** Upon execution, the system maintains all necessary master databases and Excel templates within a localized, hidden directory named "Sistema_Produzione" to ensure environment integrity.
* **Intelligent Data Mapping:** Integration of fuzzy logic algorithms to associate manual phase entries with standardized functional blocks stored in the master databases.
* **Multi-Variant Control Plan Composition:** A specialized interface allows for the mapping of a single production phase to multiple database variants or supplemental blank entries, supporting extended selection modifiers.
* **Bilateral Documentation Export:** Simultaneous generation of internal and customer-facing documentation. The system automatically filters sensitive internal phases and merges both versions into a unified final workbook.
* **Graphic and Typographic Standardization:** Utilization of Excel COM objects for workbook merging to preserve high-resolution graphical assets, logos, and strict adherence to the Calibri Light typographic standard.

## 3. System Requirements
* **Operating System:** Microsoft Windows.
* **Required Software:** Microsoft Excel (required for COM object interoperability and workbook merging).
* **Environment:** Python 3.x execution environment with initialized dependencies.

## 4. Project Structure
* **main.py:** Primary application entry point and user interface logic.
* **Sistema_Produzione/** (Hidden): Secure repository for master databases (MASTERS-FMEA.xlsx, MASTERS-CP.xlsx) and standardized templates.
* **generate_flowchart.py:** Engine for dynamic shape generation and arrow routing within the Flow Chart.
* **generate_fmea.py:** Module for FMEA compilation and automated footnote management.
* **generate_cp.py:** Module for Control Plan construction and multi-phase composition logic.
* **LeggiMasterFMEA.py / LeggiMasterCP.py:** Backend utilities for database structural formatting and cleaning.

## 5. User Manual

### 5.1 Initialization
Execute the main application script. On the initial run, the system will verify or extract the required resource files into the "Sistema_Produzione" directory. The interface will display the active paths for templates and databases in the top configuration panel.

### 5.2 Phase Definition and Queue Management
Use the input section to build the manufacturing sequence:
1. **Operation Number:** Define the numerical sequence (e.g., 10, 20, 30).
2. **Phase Identification:** Input the Phase Name. Specific component identifiers should be entered in the designated field; the system will automatically format these in parentheses for the final output while excluding them from database search strings to maintain matching accuracy.
3. **Operational Attributes:** Specify the Internal/External status and the responsible Department.
4. **Visibility Control:** Use the "Customer" checkbox to flag phases that must be included in the customer-facing version of the documents.
5. **Queue Commands:** Add the phase to the project. Phases can be reordered via the directional arrows or edited by double-clicking the entry in the list view.

### 5.3 Control Plan Composition Logic
When initiating Control Plan generation, if the system identifies multiple functional variants in the database for a specific phase, the "Control Plan Composer" will activate:
1. **Selection:** Identify the required blocks from the database results on the left panel.
2. **Assignment:** Move selected blocks to the right panel. Blank rows may be added for manual post-process notation.
3. **Sequencing:** Use the vertical controls to define the internal order of the selected blocks for that specific operation.
4. **Confirmation:** Use the Enter key or the confirmation button to commit the composition.

### 5.4 Document Finalization and Export
1. **Project Metadata:** Enter the Product Name and the current Revision index.
2. **Export Scope:** Select the specific document required or use the comprehensive "TUTTO" command to generate a unified Master Workbook.
3. **Interface Locking:** The application will display a modal progress indicator and lock inputs during the generation process to ensure data consistency.
4. **Output:** Finalized documents are exported to a structured directory named "[Product Name] Pratiche" located in the application root.