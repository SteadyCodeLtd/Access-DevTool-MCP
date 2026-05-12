# Usage Guide: New Interop-Enabled Features

This guide shows how to use the five new features added via the Interop refactoring.

## 1. Get Table Definitions

### Purpose
Inspect the complete schema of your database, including all tables and their field definitions.

### Usage

```bash
# Get definitions for ALL tables
mcp_tool: get_table_definitions

# Returns JSON:
{
  "Tables": [
    {
      "Name": "tblOrderHeader",
      "FieldCount": 15,
      "Fields": [
        {
          "Name": "OrderID",
          "Type": 4,
          "TypeName": "Long",
          "Size": 4,
          "Required": true,
          "AllowZeroLength": false
        },
        {
          "Name": "CustomerName",
          "Type": 10,
          "TypeName": "Text",
          "Size": 100,
          "Required": false,
          "AllowZeroLength": false
        }
      ]
    }
  ]
}
```

### Use Cases

1. **Database Documentation**
   - Generate automatic schema documentation
   - Export to markdown or HTML
   - Version control database structure

2. **Data Validation**
   - Verify field types and sizes
   - Check required fields
   - Identify fields allowing zero-length strings

3. **Migration Planning**
   - Compare schemas before/after updates
   - Identify schema drift in cloned databases
   - Detect incompatibilities

4. **Schema Analysis**
   - Find all Text fields over 255 characters
   - Identify currency fields across database
   - Locate date/time fields

### Example: Finding Large Text Fields

```
// Pseudo-code example
tables = get_table_definitions()
for each table in tables:
    for each field in table.Fields:
        if field.TypeName == "Text" AND field.Size > 255:
            print "Large text field: " + table.Name + "." + field.Name
```

---

## 2. Get Table Definition

### Purpose
Get detailed information about a single, specific table.

### Usage

```bash
# Get definition for specific table
mcp_tool: get_table_definition
parameters:
  table_name: "tblOrderHeader"

# Returns same structure as above, but for single table only
```

### Use Cases

1. **Focused Analysis**
   - Inspect structure of large tables without loading all tables
   - Performance optimization for large databases
   - Detailed table documentation

2. **Validation**
   - Verify a table exists before import
   - Check field definitions before data operations
   - Pre-flight checks before migrations

3. **Debugging**
   - Diagnose field-related issues
   - Verify type conversions
   - Check field constraints

### Example: Table Validation Before Import

```
// Check table before importing data
table_def = get_table_definition("tblOrderHeader")
if table_def.Fields[0].TypeName != "Long":
    error("First field should be Long integer ID")
if table_def.Fields[1].Size < 100:
    error("Customer name field too small")
proceed_with_import()
```

---

## 3. Get Objects by Type

### Purpose
Get all objects of a specific type using enum-based type selection.

### Usage

```bash
# Get all forms
mcp_tool: get_objects_by_type
parameters:
  object_type: 2  # acForm = 2

# Get all tables
mcp_tool: get_objects_by_type
parameters:
  object_type: 0  # acTable = 0

# Returns JSON:
{
  "Objects": [
    {
      "Name": "frmOrderEntry",
      "IsLoaded": false
    },
    {
      "Name": "frmCustomerList",
      "IsLoaded": true
    }
  ]
}
```

### Object Type Enum

| Type | Code | Name |
|------|------|------|
| Table | 0 | acTable |
| Query | 1 | acQuery |
| Form | 2 | acForm |
| Report | 3 | acReport |
| Macro | 4 | acMacro |
| Module | 5 | acModule |

### Use Cases

1. **Selective Enumeration**
   - Get only forms, not all objects
   - Get only reports for documentation
   - Get only modules for code analysis

2. **Object Management**
   - Find all unloaded forms
   - List only active/loaded objects
   - Track object inventory by type

3. **Batch Operations**
   - Process all modules for code review
   - Export all reports
   - Generate list of all forms

### Example: Find Unloaded Forms

```
// Find forms that aren't currently loaded
forms = get_objects_by_type(2)  // 2 = acForm
unloaded = filter(forms, f => f.IsLoaded == false)
for each form in unloaded:
    print "Unloaded form: " + form.Name
```

### Example: Count All Queries

```
// Count queries in database
queries = get_objects_by_type(1)  // 1 = acQuery
print "Total queries: " + queries.length
```

---

## 4. Compile VBA with Errors

### Purpose
Compile all VBA code and get detailed error information.

### Usage

```bash
# Compile VBA and get errors
mcp_tool: compile_vba_with_errors

# Returns JSON:
{
  "Success": false,
  "Message": "Compilation failed",
  "Errors": [
    "Module: modUtility, Line 42: Syntax Error - Expected ')'"
  ]
}
```

### Use Cases

1. **Code Quality Assurance**
   - Validate VBA before deployment
   - Catch syntax errors early
   - Verify code integrity

2. **CI/CD Integration**
   - Automated code checks
   - Pre-commit validation
   - Build pipeline quality gates

3. **Development**
   - Quick syntax validation
   - Identify compilation errors
   - Fix issues before testing

### Example: Pre-Deployment Validation

```
// Validate code before allowing deployment
compile_result = compile_vba_with_errors()
if compile_result.Success == false:
    print "Build FAILED:"
    for each error in compile_result.Errors:
        print "  - " + error
    abort_deployment()
else:
    print "All VBA code compiled successfully"
    proceed_with_deployment()
```

### Example: Error Reporting

```
// Generate error report
compile_result = compile_vba_with_errors()
error_count = compile_result.Errors.length
if error_count > 0:
    report = "VBA Compilation Report\n"
    report += "Date: " + now() + "\n"
    report += "Status: FAILED\n"
    report += "Errors: " + error_count + "\n\n"
    for each error in compile_result.Errors:
        report += "  " + error + "\n"
    save_report(report)
```

---

## 5. Get Database Summary

### Purpose
Get a quick overview of all objects in the database.

### Usage

```bash
# Get database summary
mcp_tool: get_database_summary

# Returns JSON:
{
  "Tables": 47,
  "Forms": 23,
  "Reports": 12,
  "Queries": 34,
  "Modules": 8,
  "Macros": 5,
  "TotalFields": 892
}
```

### Use Cases

1. **Database Inventory**
   - Quick asset count
   - Documentation generation
   - Capacity planning

2. **Health Checks**
   - Verify database structure intact
   - Monitor growth over time
   - Identify anomalies

3. **Reporting**
   - Database statistics reports
   - Audit documentation
   - Change tracking

### Example: Database Health Check

```
// Verify database structure
summary = get_database_summary()
print "Database Health Report"
print "  Tables: " + summary.Tables
print "  Forms: " + summary.Forms
print "  Reports: " + summary.Reports
print "  Queries: " + summary.Queries
print "  Modules: " + summary.Modules
print "  Macros: " + summary.Macros
print "  Total Fields: " + summary.TotalFields

// Check for anomalies
if summary.Tables == 0:
    alert("WARNING: No tables found!")
if summary.Forms > 100:
    alert("WARNING: Too many forms (" + summary.Forms + ")")
```

### Example: Generate Documentation

```
// Generate simple database documentation
summary = get_database_summary()
doc = "# Database Structure Report\n\n"
doc += "## Object Counts\n"
doc += "| Object Type | Count |\n"
doc += "|---|---|\n"
doc += "| Tables | " + summary.Tables + " |\n"
doc += "| Forms | " + summary.Forms + " |\n"
doc += "| Reports | " + summary.Reports + " |\n"
doc += "| Queries | " + summary.Queries + " |\n"
doc += "| Modules | " + summary.Modules + " |\n"
doc += "| Macros | " + summary.Macros + " |\n"
doc += "| **Total Fields** | **" + summary.TotalFields + "** |\n"
save_markdown_file(doc, "database_summary.md")
```

---

## Comparison: Old vs New Approaches

### Example 1: Getting All Tables

#### Old Way
```
get_forms()
get_reports()
get_modules()
get_queries()
// ... multiple calls needed
```

#### New Way
```
get_objects_by_type(acTable=0)  // Single call, type-safe
```

### Example 2: Inspecting Table Structure

#### Old Way
No built-in way to inspect detailed field definitions

#### New Way
```
table = get_table_definition("tblOrders")
for each field in table.Fields:
    print field.Name + ": " + field.TypeName + " (" + field.Size + ")"
```

### Example 3: Validating Code

#### Old Way
```
compile_vba()  // No error details returned
```

#### New Way
```
result = compile_vba_with_errors()  // Get detailed error list
if result.Success == false:
    for each error in result.Errors:
        log(error)
```

---

## Error Handling

### Example: Robust Error Handling

```
try:
    table_def = get_table_definition("tblOrders")
except "Table not found":
    print "ERROR: Table 'tblOrders' does not exist"
    create_table("tblOrders")
except error:
    print "ERROR: Unexpected error - " + error.message
```

### Common Errors

| Error | Cause | Solution |
|-------|-------|----------|
| Table not found | Misspelled name | Check table name in get_table_definitions() |
| Invalid object type | Wrong enum value | Use codes 0-5 for table/query/form/report/macro/module |
| Database not connected | No active connection | Call connect_access first |

---

## Best Practices

1. **Use Enums for Type Safety**
   ```
   // Good
   get_objects_by_type(acForm=2)
   
   // Avoid
   get_objects_by_type(2)  // Magic number
   ```

2. **Cache Results When Possible**
   ```
   // First call
   summary = get_database_summary()
   
   // Use cached value multiple times
   table_count = summary.Tables
   form_count = summary.Forms
   ```

3. **Validate Before Operations**
   ```
   table_def = get_table_definition("target_table")
   if table_def.Fields[0].Type != "Long":
       // Adjust operation based on field type
   ```

4. **Use in Combination**
   ```
   // Get summary first
   summary = get_database_summary()
   
   // Then drill down
   tables = get_objects_by_type(acTable=0)
   for each table in tables:
       table_def = get_table_definition(table.Name)
       // Process each table with full details
   ```

---

## Integration Examples

### Example 1: Database Validation Script

```
// Validate entire database structure
print "Starting database validation..."

// Check connectivity
if not is_connected():
    print "ERROR: Not connected to database"
    exit

// Get overview
summary = get_database_summary()
print "Database contains:"
print "  " + summary.Tables + " tables"
print "  " + summary.Forms + " forms"
print "  " + summary.Reports + " reports"

// Validate table structures
tables = get_objects_by_type(acTable=0)
for each table in tables:
    table_def = get_table_definition(table.Name)
    
    // Check for ID field
    has_id = any(f => f.Name == "ID", table_def.Fields)
    if not has_id:
        print "WARNING: Table '" + table.Name + "' has no ID field"
    
    // Check for large text fields
    for each field in table_def.Fields:
        if field.TypeName == "Text" and field.Size > 255:
            print "INFO: Large text field in " + table.Name + "." + field.Name

// Validate code
print "Compiling VBA code..."
compile_result = compile_vba_with_errors()
if compile_result.Success:
    print "✓ All VBA compiled successfully"
else:
    print "✗ VBA compilation errors:"
    for each error in compile_result.Errors:
        print "  " + error

print "Validation complete!"
```

### Example 2: Generate Database Documentation

```
// Auto-generate database documentation

summary = get_database_summary()
doc = "# Database Documentation\n\n"
doc += "Generated: " + now() + "\n\n"

doc += "## Overview\n"
doc += "- Tables: " + summary.Tables + "\n"
doc += "- Forms: " + summary.Forms + "\n"
doc += "- Reports: " + summary.Reports + "\n"
doc += "- Total Fields: " + summary.TotalFields + "\n\n"

doc += "## Tables\n"
tables = get_objects_by_type(acTable=0)
for each table in tables:
    table_def = get_table_definition(table.Name)
    doc += "### " + table.Name + "\n"
    doc += "| Field | Type | Size |\n"
    doc += "|---|---|---|\n"
    for each field in table_def.Fields:
        doc += "| " + field.Name + " | " + field.TypeName + " | " + field.Size + " |\n"
    doc += "\n"

save_file(doc, "database_doc.md")
print "Documentation saved to database_doc.md"
```

---

## Performance Considerations

1. **Bulk Operations**
   - `get_table_definitions()` loads all tables - use `get_table_definition()` for single table
   - `get_database_summary()` is fastest for overview
   - `get_objects_by_type()` is faster than getting all objects then filtering

2. **Caching**
   - Cache `get_database_summary()` results (rarely changes)
   - Cache `get_table_definitions()` per session (large result set)
   - Don't cache `get_objects_by_type()` frequently (may change during session)

3. **Batch Processing**
   ```
   // Get all tables once
   all_tables = get_table_definitions()
   
   // Process without repeated calls
   for each table in all_tables:
       process(table)  // Fast - no additional calls
   ```

---

## Conclusion

The new Interop-enabled features provide powerful database analysis and validation capabilities while maintaining backward compatibility with existing tools. Use them to improve database reliability, documentation, and code quality.
