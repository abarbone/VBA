# Style Guide for VBA Scripts

This Style Guide applies to all files used to store VBA Code.

## Common Requirements

1. Save files using the ".vb" file extension.
2. File names should include one of the following prefixes.
  * "block_"
  * "function_"
  * "sub_"
  * "module_"
3. Excluding the prefix, the filename should follow the UpperCamelCase convention.
4. Excluding the underscore in the prefix, do not include any special characters in file names.
5. Do not include spaces in the file names.
6. Files intended to be included in a specific Office Program should be filed into the folder named for that Office Program.
7. Tab size should be four (4) spaces to match the standard behavior of the Visual Basic Editor within the Office Programs.
8. Include one empty line of code at the end of all files.
9. Use empty lines to break apart lines of code to improve readability.

## Blocks

### Purpose

Blocks are used for lines of code that are not themselves a complete Function or Subroutine.
Blocks are intended to be used to piece together sophisticated Functions or Subroutines.
Blocks are typically used for gathering and manipulating information stored in file.
Blocks commonly gather and manipulate the contents of structured tables.

### Tab Behavior

* Line Labels should not be tabbed (tabs are removed automatically from Line Labels within the Visual Basic Editor).
* All other lines of code should start with at least a single tab.

### Header

Start each Block with the following Header.

> ''' Block Start '''
>
> ' Name: [NameOfBlock]
>
> ' Description: [Description of what the Block does]
>
> ' Commit date: [Date the Block was committed to the Main branch]

### Comments

Include comments for introducing sections of code (e.g. declaring variables and identifying indepdendent variables).

Identify lines of code that require the user to make changes similar to the example below.

> NameOfTable = "Folders" '<- input name of table.

Identify lines of code where the user may want to change to the name of variables similar to the example below.

> Dim GenArray() As String '<- adjust the name of the array as necessary

### Footer

End each block with the line below:

> ''' Block End '''

## Functions

### Purpose

To store Function procedures. 
Refer to Microsoft's documentation for information related to Function procedures.

### Tab Behavior

* The first line of code that introduces the Function should not be tabbed.
* The final line of code that ends the Function should not be tabbed.
* Line Labels should not be tabbed (tabs are removed automatically from Line Labels within the Visual Basic Editor).
* All other lines of code should start with at least a single tab.

### Header 

Include the following header for each Function.
The header should appear below the first line of code that introduces the Function.

> ' Description: [Description of what the Function does]
>
> ' Commit date: [Date the Function was committed to the Main branch]

### Comments

Include descriptive comments throughout the Function.

## Subroutines

### Purpose

To store Subroutines. 
Refer to Microsoft's documentation for information related to Subroutines.

### Tab Behavior

* The first line of code that introduces the Subroutine should not be tabbed.
* The final line of code that ends the Subroutine should not be tabbed.
* Line Labels should not be tabbed (tabs are removed automatically from Line Labels within the Visual Basic Editor).
* All other lines of code should start with at least a single tab.
* Tabs should be used to align code per the Outline.

### Header

Include the following Header for each Subroutine.
The Header should appear below the first line of code that introduces the Subroutine.

> ' Description: [Description of what the Subroutine does]
>
> ' Commit date: [Date the Subroutine was committed to the Main branch]

### Outline

All Subroutines should follow the outline indicated with the Headings below.

1. Setup.  
    1.1. Set error behavior.  
    1.2. Declare variables.  
    1.3. Populate variables.  
2. Check inputs.  
3. Perform actions.  
4. Closeout.  

These Headings should appear in all Subroutines. 
If there is no code for a required Heading, include a comment indicating there is no code for that section.

Include additional Headings as necessary to make the code clear. 
Include comments beneath headings if a lengthy description is needed.

Headings should be indented according to their Level.
For instance, a Level 1 Heading (e.g. "1. Setup.") should be indented once.
A Level 2 Heading (e.g. "1.2. Declare variables.") should be indented twice.
The code beneath each heading should be indented the number of times that the heading is indented plus one.
This makes the code collapsible at each Heading.

### Comments

The Outline and Headings serve to provide some basic commenting to the code.
Include additional comments throughout to improve clarity.

## Modules

### Purpose

To store a collection of closely-related or interdependent Functions and Subroutines.
Blocks should not be included in a Module.

### Formatting

The Functions and Subroutines included in a Module should follow the requirements above.
