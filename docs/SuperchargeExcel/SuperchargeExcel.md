# 2: Concept: Loading Data

> Highlight (yellow) - Location 493

Business users can think of a dimension table as a lookup table and a fact table as a data (or transactions) table.

> Highlight (yellow) - Location 546

A data table contains transactional information — in this case sales transactions. Lookup tables contain information about logical groups of objects, such as customers, products, time (calendar), etc.

> Highlight (yellow) - Location 550

Each of these tables must have a unique ID column, such as ProductNumber, CustomerNumber, etc. These unique columns are sometimes called keys.

> Highlight (yellow) - Location 557

Unlike in other database programs (Power Pivot is actually a database), there is no other type of table join available in Power Pivot

> Highlight (yellow) - Location 558

The data table may contain none, one, or many rows of data for each row in the lookup table. The following “Here’s How” shows how to join tables.

> Highlight (yellow) - Location 584

Because the relationships are always one-to-many, the joins are specifically single-directional. Always drag from the data table up to the lookup table, not the other way around

> Highlight (yellow) - Location 635

The Collie layout methodology involves placing the lookup tables at the top of the window and the data tables at the bottom.

> Highlight (yellow) - Location 640

In the IT world, lookup tables are referred to as dimension tables, and data tables are called fact tables.

> Highlight (yellow) - Location 648

It is no longer necessary to bring data from the lookup tables into the data tables by using VLOOKUP(). Instead, you can simply load the lookup tables and join them with a relationship.

> Highlight (yellow) - Location 654

A key feature of a lookup table is that it contains one and only one row for each individual item in the table, and it has as many columns as needed to describe the object.

> Highlight (yellow) - Location 664

In this case, the Sales table contains one column (technically called a foreign key) that matches each of the keys in each lookup table (technically called a primary key). Stated differently, the Sales data table has four foreign key columns: a date, a customer number, a product number, and a territory key.

> Highlight (yellow) - Location 670

Ideally, data tables should have very few columns but as many rows as needed to bring in all the data records. Data tables normally have lots of rows (sometimes in the tens of millions or even billions).

> Highlight (yellow) - Location 778

The data connections you create in Power Query are relative to your computer. When you send a workbook and data source to another user, that person will have to edit the data connection so that it will work on his or her own PC.

# 3. Concept: Measures

> Highlight (yellow) - Location 826

A measure is simply a DAX formula that instructs Power Pivot to do a calculation on data. In a sense, a measure is a lot like a formula in a cell in Excel. The main difference, however, between a formula in a cell in Excel and a measure is that a measure always operates over the entire data model, not over just a few cells in a spreadsheet.

## Techniques for Writing DAX Measures

> Highlight (yellow) - Location 839-840

You can write a measure in the formula bar in the Power Pivot window, as shown below. If you use this method, you must specify the measure name followed by a colon and then the formula. Note that there can be no spaces between the measure name, the colon, and the equals sign.

> Matt Allington. Supercharge Excel (Kindle Locations 843-844).

You can write and edit measures in any empty cell in the calculation area at the bottom of the Power Pivot window, as shown below. Note that you need to use a colon here, too.

> Matt Allington. Supercharge Excel (Kindle Locations 846-847).

You can write measures in the Measure dialog in Excel, as shown below. You can open this dialog from within Excel by navigating to the Power Pivot tab (see #1 below) and clicking on the Measures button (# 2) and then New Measure. In general, Excel users should write DAX in the Measure dialog box in Excel. And it is normally best to first create a pivot table that provides some context for a measure you are about to write. If you do it this way, you immediately see the measure in a pivot table when you click OK

### Writing Measures

Follow these steps:

1. Create a new, blank pivot table connected to the data model.
2. Add data to the rows.
3. Click inside the pivot table, go to the Power Pivot tab and select Add Measure.
4. Make sure to place the measure in the same table that the data comes from.
5. Use descriptive & unique names
6. Write the DAX formula & click 'Check Formula' to validate
7. Select the desired data format & click 'OK'

#### Avoiding Implicit Measures

Implicit Measures are what traditional pivot tables used when you used 'Count of x' or 'Sum of y' in the values pane.

> Matt Allington. Supercharge Excel (Kindle Locations 945-946).

It is best practice in DAX to always type the table name before every column name inside your formulas

Use the 'Manage Measures' dialog box to view, edit or delete measures from the workbook. Since the Measure dialog box is modal you can't do anything while it is open. If you are having problems with a DAX formula you can either set the measure `=1` or wrap the formula in double quotes so it doesn't evaluate. If you need to use double quotes in the formula, replace the double quotes in the formula with single quotes then come back to it later.

# 4. DAX Topic: SUM(), COUNT(), COUNTROWS(), MIN(), MAX(), COUNTBLANK(), DIVIDE()

## Aggregate Functions

These functions take inputs from a column or table and aggregate the contents. You have to tell Power Pivot how to aggregate the data so it returns just a single value to each cell in the pivot table. The iterable nature of applying these aggregations to each cell is what makes power pivot so powerful.

### Practice Problem:

Using the data we imported from Access, create the following measures:

1. [Total Sales]
   `=SUM ( Sales[ExtendedAmount] )`
2. [Total Cost]
   `=SUM ( Sales[ProductStandardCost] )`
3. [Total Margin $]
   `=[Total Sales] - [Total Cost]`
4. [Total Margin %]
   `=DIVIDE ( [Total Margin], [Total Sales], "" )`
5. [Total Sales Tax Paid]
   `=SUM ( Sales[TaxAmt] )`
6. [Total Sales Including Tax]
   `=[Total Sales] + [Total Sales Tax Paid]`
7. [Total Order Quantity]
   `=SUM ( Sales[OrderQuantity] )`
8. [Total Number of Products]
   `=COUNT ( [ProductKey] )`
9. [Total Number of Customers]
   `=COUNT ( Customers[CustomerKey] )`
10. [Total Products Using COUNTROWS]
    `=COUNTROWS ( Products )`
11. [Total Customers Using COUNTROWS]
    `=COUNTROWS ( Customers )`
12. [Total Customers using DISTINCTCOUNT]
    `=DISTINCTCOUNT ( Customers[CustomerKey] )`
13. [Count of Occupation]
    `=DISTINCTCOUNT ( Customers[Occupation] )`
14. [Count of Country]
    `=DISTINCTCOUNT ( Territory[Country] )`
15. [Customers that Have Purchased]
    `=DISTINCTCOUNT ( Sales[CustomerKey] )`
16. [Maximum Tax Paid on a Product]
    `=MAX ( Sales[TaxAmt] )`
17. [Minimum Price Paid for a Product]
    `=MIN ( Sales[SalesAmount] )`
18. [Average Price Paid for a Product]
    `=AVERAGE ( Sales[SalesAmount] )`
19. [Customers Without Address Line 2]
    `=COUNTBLANK ( [AddressLine2] )`
20. [Products Without Weight Values]
    `=COUNTBLANK ( Products[Weight] )`
21. [Margin %]
    `=DIVIDE ( [Total Margin], [Total Sales] )`
22. [Markup %]
    `=DIVIDE ( [Total Margin], [Total Cost] )`
23. [Tax %]
    `=DIVIDE ( [Total Sales Tax Paid], [Total Sales] )`

### Pivot Table Conditional Formatting

Highlight the cells, click 'Conditional Formatting' and select the visualization. It's easier to select one of the cells, apply the formatting, then use the pop-up box to apply the formatting to the _rows only_, _leaving out the grand total_.

### When are Measures Automatically Added?

The pivot table must be selected before writing the measure AND you have to save the measure without creating an error on save. If the formula isn't checked and creates an error the measure will not be automatically added to the pivot table after the error is fixed.

### Filter Context

> Matt Allington. Supercharge Excel (Kindle Locations 1451-1458).

Note that the first measure, [Customers Without Address Line 2], is being filtered by the pivot table (i.e., Customers[ Occupation] on Rows), and the values in the pivot table change with each row. But the second measure, [Products Without Weight Values], is not filtered; the values don’t change for each row in the pivot table. The technical term for filtering behavior in Power Pivot is _filter context_.

> Matt Allington. Supercharge Excel (Kindle Locations 1468-1470).

If you use DIVIDE() instead of the slash operator (/) for division, DAX returns a blank where you would otherwise get a divide-by-zero error. Given that a pivot table will filter out blank rows by default, a blank row is a much better option than an error.If you don’t specify the alternate result, a blank value is returned when there is a divide-by-zero error.
