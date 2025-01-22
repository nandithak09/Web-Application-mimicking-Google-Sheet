# Web Application Mimicking Google Sheets

This web application mimics the functionality of Google Sheets. It allows users to interact with a spreadsheet interface, enter data, perform mathematical functions, apply data quality operations, and visualize data through charts. The app supports key features such as cell formatting, formulas (e.g., SUM, AVERAGE), and saving/loading data. The application has been designed with performance and security in mind.

## Features

- Spreadsheet Interface:
  - Mimics the Google Sheets UI (rows, columns, formula bar).
  - Drag functionality for cells, formulas, and selections.
  - Basic cell formatting: bold, italic, font size, color.
  - Add, delete, and resize rows/columns.

- Mathematical Functions:
  - SUM: Sum of a range of cells.
  - AVERAGE: Average of a range of cells.
  - MAX: Maximum value from a range of cells.
  - MIN: Minimum value from a range of cells.
  - COUNT: Count cells with numerical values.

- Data Quality Functions:
  - TRIM: Removes leading and trailing whitespace.
  - UPPER: Converts text to uppercase.
  - LOWER: Converts text to lowercase.
  - REMOVE_DUPLICATES: Removes duplicate rows.
  - FIND_AND_REPLACE: Replace specific text in cells.

- Bonus Features:
  - Additional mathematical functions: (e.g., AVERAGE, MAX, MIN).
  - Complex formula support: Relative and absolute references.
  - Save/Load functionality: Save spreadsheet state to localStorage.
  - Data Visualization: Generate line charts using **Chart.js.


  ## Tech Stack

- *HTML*: For structuring the content and user interface (UI).
- *CSS*: For styling the spreadsheet and its components.
- *JavaScript*: For the core functionality of the web application, including data management, user interaction, and formula evaluation.
- *Chart.js*: For generating charts and visualizing data.


## Data Structures

1. *sheetData (2D Array)*:
   - This is the primary data structure used to store the content of the spreadsheet.
   - It's a 2D array where each element represents a cell in the sheet, and its value can be text, numbers, or formulas.
   - The reason for using a 2D array is that it allows quick access to any cell by its row and column indices.

   ```javascript
   let sheetData = []; // 2D array to store spreadsheet content

2. *dependencyGraph (Object)*:

   - This is used to track dependencies between cells that contain formulas.
   - The keys are cell references (e.g., "A1", "B2"), and the values are arrays of dependent cells that should be recalculated when the original cell is updated.
   - This data structure helps us manage the dynamic nature of formulas in the spreadsheet and ensures that updates to one cell automatically trigger recalculations of cells that depend on it .

   let dependencyGraph = {}; // Object to store cell dependencies

3. *Drag-and-Drop Variables*:

   - isDragging: A flag to determine if a drag operation is active.
   - dragStartCell: A reference to the starting cell when the drag operation begins.
    These variables allow for simulating the drag-and-drop behavior to copy or fill values between cells.

   let isDragging = false; // Flag to track drag operation state
   let dragStartCell = null; // Reference to the cell where drag started


## Why these Data Structures
   
1. 2D Array (sheetData): A 2D array is ideal for representing a grid of rows and columns, which is exactly how a spreadsheet works. It allows for fast and efficient updates to individual cells by referencing their indices.

2. Object (dependencyGraph): The use of an object to track cell dependencies is essential for managing formulas and ensuring that dependent cells are recalculated when the referenced cells change. An object allows for quick lookups and efficient updates to dependent cells.

3. Event-Driven Logic (Drag-and-Drop): By using flags like isDragging and references like dragStartCell, the application can handle complex drag operations and update the sheet dynamically in response to user actions.

## Security Considerations

To ensure the security of the application, the following measures have been implemented:

- Data Sanitization: All user inputs are validated to ensure no malicious code (e.g., XSS) is injected into the application. User data is processed and sanitized before being stored.
- No External Server Interaction: The application uses localStorage for saving data, which is stored locally on the user's machine. This ensures that the application is self-contained and does not require external server interaction.
- Formula Parsing: The formula parsing only allows a set of predefined mathematical functions (SUM, AVERAGE, etc.) to prevent arbitrary code execution.

## Performance Enhancements

- Efficient DOM Manipulation: The spreadsheet data is stored in memory (in the sheetData array) to minimize direct DOM updates. Changes are only reflected when necessary, improving performance.
- LocalStorage: Instead of relying on network requests for saving/loading data, the app uses the localStorage API to store and retrieve spreadsheet data quickly, which also reduces network load.
- Batch Updates: Changes to multiple cells or rows are batched to minimize reflows and repaints in the DOM, improving responsiveness.

## Installation Instructions

To run the web application:

1. Download or clone the repository.
2. Open index.html in your browser.

No server setup or backend configuration is required, as the app is fully client-side.

## Usage Instructions

1. Start the Application:
Open the HTML file in a browser to interact with the spreadsheet.
The sheet is initialized with 10 rows and 10 columns. Users can add more rows/columns as needed.
2. Input Data:
Cells are editable, and users can input numbers, text, and formulas.
3. Apply Functions: 
Use the toolbar or input functions (e.g., =SUM(A1:B5)) to perform calculations and apply data quality operations.
4. Drag-and-Drop:
Click and drag a cell to copy its content or formula to other cells.
5. Save and Load:
Click the "Save" button to store the current sheet data in localStorage.
Click "Load" to retrieve the saved sheet data.
6. Generate Chart:
Use the "Generate Chart" button to visualize the data in a line chart.

## Contribution Guidelines

1. Fork the repository.
2. Make changes in your fork and test them thoroughly.
3. Open a pull request describing your changes.
