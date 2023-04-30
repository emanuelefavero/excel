# Excel

This is a cheat sheet repo for Excel

> Note: In this tutorial I will be using Google Sheets, but the same concepts apply to Excel

## Keyboard Shortcuts

| Shortcut        | Description                |
| --------------- | -------------------------- |
| Tab             | Go to next cell in the row |
| cmd + shift + v | Paste values without style |

## Tips

### Formulas

- You can add formulas to a cell by typing `=` and then the formula
- After typing `=` you can click on a cell to add it to the formula (e.g. `=C4*D4`)
- You can drag the bottom right corner of a cell to copy the formula to other cells
- Select multiple cells and drag the bottom right corner to copy the formula to multiple cells
- If you copy a cell with a formula and paste it into another cell, the formula will be updated to use the new cell's position

### Functions

- `=MAX(A1:A10)` - Returns the maximum value in the range
- `=MIN(A1:A10)` - Returns the minimum value in the range
- `=SUM(A1:A10)` - Returns the sum of the values in the range
- `=AVERAGE(A1:A10)` - Returns the average of the values in the range
- `=COUNT(A1:A10)` - Returns the number of values in the range
- `=COUNTA(A1:A10)` - Returns the number of non-empty values in the range
- `=COUNTBLANK(A1:A10)` - Returns the number of empty values in the range
- `=COUNTIF(A1:A10, ">10")` - Returns the number of values in the range that are greater than 10
- `=COUNTIFS(A1:A10, ">10", B1:B10, "<20")` - Returns the number of values in the range that are greater than 10 and less than 20
- `=SUMIF(A1:A10, ">10")` - Returns the sum of the values in the range that are greater than 10
- `=SUMIFS(A1:A10, ">10", B1:B10, "<20")` - Returns the sum of the values in the range that are greater than 10 and less than 20
- `=AVERAGEIF(A1:A10, ">10")` - Returns the average of the values in the range that are greater than 10
- `=AVERAGEIFS(A1:A10, ">10", B1:B10, "<20")` - Returns the average of the values in the range that are greater than 10 and less than 20
- `=IF(A1>10, "Greater than 10", "Less than 10")` - Returns "Greater than 10" if the value in A1 is greater than 10, otherwise returns "Less than 10"

### Absolute References

- You can use `$` to make a reference absolute (e.g. `$A$1`)
- If you copy a cell with an absolute reference and paste it into another cell, the reference will not be updated

> Note: It is useful to use absolute references when you are referencing a cell that you do not want to change (e.g. a constant)
>
> BEWARE: If the '$' symbol is placed before the column letter (e.g. `$A1`) then the column will not change when the formula is copied, but the row will and vice versa (e.g. `A$1`)

### Conditional Formatting

Conditional formatting allows you to change the style of a cell based on its value

- Select the cells you want to apply the conditional formatting to
- Click on the "Format" menu and select "Conditional formatting"
- Select the type of conditional formatting you want to apply (e.g. "Greater than")
- Enter the value you want to compare the cell to (e.g. 10), add the percentage if you want to compare it to a percentage (e.g. 10%, or use 0.1 for 10%)
- Select the style you want to apply to the cell if the condition is met (e.g. "Bold")
- Click on the "Done" button

### General Tips

- You can use the `=` operator to convert a value to a number (e.g. `=A1*2`)
- You can click on a column header to select the entire column
- You can add a column by right clicking on a column header and selecting "Insert 1 left"
- You can calculate the number of days between two dates by subtracting them (e.g. `=A1-A2`)
- Double click on the column right border to auto resize the column to fit the content
- You can rotate text in a cell by selecting the cell and then clicking on the "Text Rotation" button in the toolbar
