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

### Initial Filter Context

It is important that you learn to “read” the initial filter context from your visuals because it will help you understand how each value in a visual is calculated. And it is important to refer to the full table name and column name because that forces you to look, check, and confirm exactly which tables and columns you are using in your visuals.

The filters automatically flow from the “one” side of the relationship to the “many” side of the relationship, in the direction of the arrow; or you can think of the filters as flowing from the lookup table to the data table. Whatever terms you use, it’s always downhill. The connected table - the Sales table - is then also filtered.

It's important to remember that all cells are evaluated on their own, without regard for any other cell in the visual, even if the cell is a subtotal or grand total cell.

# 6. Concept: Lookup Tables & Data Tables

## Data Tables

Data tables are typically the largest tables loaded into Power BI. Examples of data tables are Sales, Budget, Exchange Rates, General Ledgers and stock counts. There is no limitation on how often similar transactions can occur and be stored in a data table.

## Lookup Tables

Lookup tables tend to be smaller and wider than data tables. Some examples include Customers, Products and Calendars. Lookup tables **must** have a uniquely identifying code of some type to uniquely differentiate each row in the table - the primary key.

## Denormalising Tables

In old Excel pivot tables had to be based on one, single, table. You would then write a bunch of lookup functions to bring information from other tables into the source of your pivot table data. The problem with this is the duplication of data - files quickly become bloated and inefficient.

> The more unique values a column has, the less the data will be compressed by the PBI data model. The number of columns in your data tables is much more important than the number of rows.

Fewer columns and more rows is better than more columns and fewer rows, especially for large tables.

## Joining Tables Using Relationships

To avoid repetitive data, keep that data in separate subtables. For example, if the sales table contains the unique product key, it can fetch any extra information from the product master table whenever it's needed. Once the relationship is created the tables work together as if they were a single unit without the need to create duplicate data in the Sales table.

## Schemas

### Star Schemas

The data table is the center of the star. In our case, this is the Sales table. There are supporting lookup tables that add information using primary & foreign keys. You could use vlookups to fetch the columns from other tables, but that's not necessary with relationships.

## General Advice

1. Keep data tables long and skinny. Get rid of extra columns by unpivoting tables
2. Move repeating attribute columns from data tables to lookup tables.
3. If your lookup tables are joined to other lookup tables, consider flatening them into a single, wider, lookup table.

# 7. DAX Topic: The Basic Iterators `SUMX()` and `AVERAGEX()`

Iterative functions have _row context_. This means that the function is "aware" of which row it is referencing at any point in time.

## `SUMX(table, expression)`

SUMX creates a row context in the specified table then iterates through each row, one at a time, evaluating the expression for each row it gets to before it finally adds the interim results for each row. There is no need to wrap the columns in an aggregation function when using X-functions because that's essentially what these function are doing.

During each step of the iteration process, the column names in the expression are only referring to a single cell - the one at the intersection of the single column and the current row. This works like a running total. One row at a time, the single value in the `Sales[ExtendedAmount]` column is added to the single value of the `Sales[TaxAmt]` column.

### When to use X-Functions vs Aggegators

When data doesn't contain a line total. For example, if your table has a column for quantity and another for price per unit, but doesn't include Total Sales. These are useful functions for when you need to calculate an average or multiply values. The example in the book uses the multiplication of the averages at the grand total level. Think of which level the aggregation needs to apply to. Things can start getting funky if you start doing row-level operations where they should be column-wise / categorical (think: averages).

# 8. DAX Topic: Calculated Columns

Generally, favor measures & Power Query over calculated columns. You should, however, use calculated columns if:

1. You need to filter / slice a visual based on the results of a column
2. You can't bring the column of data you need in from your source data or by using Power Query

Ideally, you want to push the column as far back to the source as possible. If you can add the column in Access or Power Query, do it there instead of a calculated column.

# 9. DAX Topic: `CALCULATE()`

`Calculate(expression, filter 1, filter 2, filter n...)`

The `CALCULATE()` function alters the filter context coming from the visual by applying none, one or more filters **prior** to evaluating the expression. Taking a filter from a lookup table and propagating it to the data table is what the Power Pivot engine in Power BI was built and optimized to do. Just think about the interactions between visuals and how changing the context of one visual affects the rest of the visuals on that page.
