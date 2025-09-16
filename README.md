This SOLIDWORKS VBA macro automates the process of exporting drawings for all components within an active assembly. It organizes the output by creating a dedicated folder for each component and naming the exported PDF file with the component's name and revision number.

---
## Main Routine: `ExportDrawingsToCustomFolders`

This is the primary subroutine that executes the entire workflow. ⚙️

1.  **Initialization**: The macro first hooks into the active SOLIDWORKS application and checks to ensure the currently open document is an assembly. If not, it alerts the user and stops.
2.  **Destination Selection**: It prompts the user to select a parent (root) folder where all the exported files will be saved. This is handled by the `GetParentFolderPath` function.
3.  **Component Iteration**: The macro then iterates through every component in the assembly.
4.  **Processing Logic (per component)**:
    * It **skips** any virtual components (components saved internally within the assembly) as they do not have their own separate files.
    * It determines the component's **drawing file path** by assuming it has the same name and location as the part file, but with a `.SLDDRW` extension instead of `.SLDPRT`.
    * If the corresponding drawing file exists, it **creates a subfolder** within the parent destination folder, named after the component.
    * It opens the drawing, retrieves the latest revision using the `GetLatestRevisionFromDrawing` function, and sanitizes the revision string.
    * Finally, it **exports the drawing as a PDF** into the newly created subfolder. The PDF is named using the format: `PartName-Revision.pdf`.
    * The drawing is closed after the PDF is saved to manage memory.
5.  **Completion**: Once all components have been processed, a message box notifies the user that the task is complete and indicates the destination folder.

---
## Helper Function: `GetLatestRevisionFromDrawing`

This function is responsible for finding the latest revision number from a given SOLIDWORKS drawing. 🧐

* **Primary Method**: It first attempts to read the revision from a custom property named **"Revision"** in the drawing file. If this property exists and has a value, it's used immediately.
* **Fallback Method**: If the custom property is not found, the function scans the drawing's **revision table**. It iterates through the first column of the table to find the highest numerical value, which it assumes is the latest revision.
* **Return Value**: It returns the found revision as a string, formatted as `RevX.X`. If no revision can be found, it defaults to `Rev0.0`.

---
## Helper Function: `GetParentFolderPath`

This simple utility function provides a standard Windows folder browser dialog. 📁

* It uses the `Shell.Application` object, which allows it to function without any dependencies on Microsoft Office applications (like Excel or Word).
* It prompts the user with the message "Select Parent Folder..." and returns the full path of the selected folder. If the user cancels the dialog, it returns an empty string.
