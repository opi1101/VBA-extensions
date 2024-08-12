# VBA extension modules
These modules are a collection of useful functions and subroutines for Visual Basic applications. 
General purpose was to make readable code, extend VBA functionality and to speed up developement.

| Module name | Description |
| ------------- | ------------- |
| `libcore` | Contains utility functions for VBA language. Type extensions: `Array`, `String`, `Path`, `Directory`, `File` and more. |
| `libexcel` | Contains utility functions for Excel types. Type extensions: `Workbook`, `Worksheet`, `Range`... |

<br>

> [!IMPORTANT]
> Target OS is `Microsoft Windows`.
> 
> `libcore` has functions that reference Excel functions. These functions will create an Excel application, if needed.
> `libexcel` module references `libcore` module functions. Don't forget to import both when using `libexcel`.

# How to
- Import modules to your projects, or copy specific functions,
- Function names start with the name of the type they intended to extend,
  * `Array` functions start with the word 'Array...()', ex. `ArrayDimensionCount()`,
  * `Range` extensions start with the word 'Range...()', ex. `RangeIsEmpty()`,
  * and so on...
- Functions use late binding so it does not require you to add library references,
  * I recommend using early binding, but late binding allows you to manage application behaviour,
- Most Excel extensions don't mess with `ScreenUpdating` or `EnableEvents` properties (unless needed),
- Excel sheet, column & row parameters are variants,
  * You can either use string, number or object references as well,
- Some functions raise errors, make sure to handle them,
- I don't use RubberDuck, but i did leave RubberDuck annotations to describe functions/modules,
- I tend to trigger errors to check valid data (some might find it a bad habit),

# TODO
- [ ] Documentation
