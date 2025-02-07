# FuzzySheets (WIP)

FuzzySheets is a utility project designed to enhance the robustness of Excel parsers through the systematic mutation of Excel files. This tool facilitates fuzzy testing by introducing a variety of controlled mutations that mimic common data anomalies and formatting variations typically encountered in real-world Excel documents.

## Purpose

The primary goals of FuzzySheets are to:

- **Improve Parser Reliability:**  
  By generating mutated versions of Excel files, the tool helps uncover potential issues in Excel parser implementations, ensuring they can gracefully handle diverse data scenarios.

- **Enable Fuzzy Testing:**  
  The project supports fuzz testing by applying systematic mutations—such as altering number formats, modifying date representations, and introducing text encoding anomalies—to simulate edge cases and data inconsistencies.

- **Provide a Modular Architecture:**  
  The core mutation engine is designed as a reusable .NET Standard library, making it independent of any specific Excel parser (like EPPlus). This allows for easy integration into other projects or replacement of the underlying Excel parsing library without affecting the mutation logic.

## Key Features

- **Core Mutation Engine:**  
  - A dedicated library that applies various mutations to Excel files.
  - Supports both data and numeric mutations.
  - Configurable via JSON, enabling easy import/export of mutation settings.
  - Batch processing capabilities to generate multiple mutated files in one go.

- **Standalone GUI Application:**  
  - Built using .NET MAUI, the GUI provides a user-friendly interface for selecting files, configuring mutations, and executing the mutation process.
  - Separates UI concerns from the core logic, ensuring a clean, maintainable codebase.

- **Generic Domain Model:**  
  - Uses custom domain objects (e.g., `Workbook`, `Worksheet`, `Cell`) and interfaces to abstract the Excel file structure.
  - Decouples the mutation engine from the Excel parsing library, allowing for future flexibility and easier testing.

## Architectural Overview

- **Separation of Concerns:**  
  FuzzySheets is split into a core library (handling the mutation logic) and a UI layer (handling user interactions). This modularity ensures that each component can be developed, tested, and maintained independently.

- **Generic Representation of Excel Files:**  
  A custom domain model represents Excel workbooks, worksheets, and cells. This abstraction provides flexibility to handle multi-sheet workbooks and to extend functionality (e.g., cell formatting or formula mutations) in future iterations.

- **Extensibility:**  
  The project is designed with future enhancements in mind, such as adding live previews, supporting additional mutation types (like formula errors), and integrating the mutation engine into larger applications.

## Conclusion

FuzzySheets is a strategic utility designed to drive improvements in Excel parsing reliability by simulating real-world data mutations. Its modular architecture, built on .NET Core and .NET MAUI, ensures that the mutation engine remains both flexible and reusable across various contexts. This foundation sets the stage for further enhancements and integration into broader testing ecosystems.
