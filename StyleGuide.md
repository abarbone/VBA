# Style Guide for VBA Scripts
This Style Guide applies to all files used to store VBA Code.

## Common Requirements
1. Save files using the ".vb" file extension.
2. File names should include one of the following prefixes.
  * "block_"
  * "function_"
  * "module_"
  * "sub_"
3. Excluding the prefix, the filename should use the UpperCamelCase convention.
4. Do not include any spaces or special characters in file names.
5. Files intended to be included in a specific Office Program should be filed into the folder named for that Office Program

## Blocks

### Purpose
Blocks are used for lines of code that perform actions between functions and subroutines.
Blocks are typically used for gathering information stored in tables within a file.

### Header
Start each block like the example below:
> ''' Block Start '''
>
> ' Name: [NameOfBlock]
>
> ' Description: [Description of what the block does]
>
> ' Commit date: [Date the block was committed to the Main branch]
