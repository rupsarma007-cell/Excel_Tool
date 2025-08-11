Excel Data Assistant - Release Notes
We're excited to announce a significant update to your Excel Data Assistant! This release brings a host of new features, powerful data manipulation tools, and crucial stability improvements to enhance your data analysis and management experience.

✨ New Features
This update introduces several new capabilities to make your workflow smoother and more efficient:

Merge Excel Files (Folder): You can now easily combine all Excel files from a selected folder into a single DataFrame. Find this under File > Merge Excel Files (Folder)....

Split Sheet by Column Value: Divide your current worksheet into multiple Excel files based on the unique values in a chosen column. Access this via File > Split Sheet by Column Value....

Batch Export: Export your current DataFrame into multiple formats (Excel, CSV, PDF) simultaneously to a chosen directory. Look for File > Batch Export....

Advanced Search: Perform sophisticated searches using multiple keywords across selected columns and view matching rows in a dedicated window. Available under Tools > Advanced Search....

Fill Missing Values: Quickly fill empty (NaN) cells in a selected column with a specified value (e.g., 0, 'N/A', mean, or median). Use Tools > Fill Missing Values....

Convert Data Type: Easily change the data type of a selected column to numeric, datetime, or string. Access this via Tools > Convert Data Type....

Split Column: Divide a text column into multiple new columns using a custom delimiter. Find this under Tools > Split Column....

Auto-Number Rows: Add a new column to your DataFrame with sequential numbering for easy row identification. Available under Tools > Auto-Number Rows....

Conditional Formatting Report: Generate reports for rows that meet specific conditional criteria like "Top N," "Bottom N," "Duplicates," "Greater Than," or "Less Than." Access this via Tools > Conditional Formatting Report....

Date Filters: Filter your data to show only rows within a specified date range in a selected datetime column. Find this under Tools > Date Filters....

Compare Two Columns (Two Files): Analyze and compare data between two columns from different Excel files, providing insights into unique and common values. Available under Tools > Compare Two Columns (Two Files).

Future Enhancements Info: A new section under Help > Future Enhancements... providing a glimpse into upcoming planned features.

🛠️ Improvements & Bug Fixes
We've addressed several issues and enhanced existing functionalities for improved performance and reliability:

Descriptive Statistics Dropdown Fix: The unique values dropdown and subsequent filtering in the "Descriptive Statistics" feature (Analyze > Descriptive Statistics) have been significantly improved. The application now correctly handles type conversion for filtered values (numeric, datetime, boolean, and string), ensuring accurate statistics regardless of the data type in the selected column. This resolves the previous issue where filtering might not have worked as expected.

General Performance: Long-running operations, such as file loading and complex data processing, now run asynchronously in separate threads, preventing the user interface from freezing.

UI/UX Enhancements: The application now utilizes the ttkbootstrap library, providing a more modern and visually appealing interface.

Treeview Cell Editing: You can now directly edit cell values within the data preview Treeview by double-clicking on a cell, or using the "Edit selected cell" button.

PDF Export Dependency Handling: The PDF export feature is now optional; if the fpdf library is not installed, the feature will be gracefully disabled with an informative message.

Robust Data Saving: Improved error handling for saving Excel files, providing more specific feedback on issues like permission errors.

🚀 How to Get the Latest Version
To experience these new features and improvements, simply run the provided Python script. Ensure you have all necessary libraries installed (pandas, openpyxl, matplotlib, ttkbootstrap, and optionally fpdf for PDF features) using pip install <library_name>.

We hope these updates significantly enhance your experience with the Excel Data Assistant!
