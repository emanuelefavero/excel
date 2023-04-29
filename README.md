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
