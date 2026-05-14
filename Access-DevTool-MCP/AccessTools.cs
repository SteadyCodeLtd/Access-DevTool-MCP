using System.ComponentModel;
using System.Runtime.Versioning;
using System.Text.Json;
using Microsoft.Win32;
using ModelContextProtocol.Server;

namespace AccessAceMcp;

[McpServerToolType]
[SupportedOSPlatform("windows")]
public sealed class AccessTools(PipeChannel channel)
{
    // ── Connection ───────────────────────────────────────────────────────────────

    [McpServerTool(Name = "connect_access")]
    [Description("Connect to an Access database via COM automation (required before all other tools)")]
    public Task<string> ConnectAccess(
        [Description("Full path to the .accdb or .mdb file")] string database_path)
    {
        var bitness = GetAccessBitness();
        if (string.Equals(bitness, "x64", StringComparison.OrdinalIgnoreCase))
            return Task.FromResult(JsonSerializer.Serialize(new
            {
                success = false,
                error = "64-bit Microsoft Access is installed but is not supported. " +
                        "This MCP server requires 32-bit (x86) Access. " +
                        "Please install the 32-bit version of Microsoft Office/Access."
            }));

        return channel.CallAsync("connect_access", new { database_path });
    }

    /// <summary>
    /// Detects the bitness of the installed Microsoft Access by inspecting the registry.
    /// Checks Office versions 16.0, 15.0, and 14.0 in order.
    /// Returns "x86", "x64", or "Not Found".
    /// </summary>
    [SupportedOSPlatform("windows")]
    private static string GetAccessBitness()
    {
        foreach (var version in new[] { "16.0", "15.0", "14.0" })
        {
            var result = GetAccessBitnessForVersion(version);
            if (result != "Not Found")
                return result;
        }
        return "Not Found";
    }

    [SupportedOSPlatform("windows")]
    private static string GetAccessBitnessForVersion(string officeVersion)
    {
        try
        {
            using var hklm64 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64);

            // Prefer the Outlook key — it carries an explicit Bitness value ("x86" or "x64")
            // and is present on any full Office installation that includes Access.
            var outlookKey = hklm64.OpenSubKey($@"SOFTWARE\Microsoft\Office\{officeVersion}\Outlook");
            if (outlookKey != null)
            {
                var bitness = outlookKey.GetValue("Bitness")?.ToString();
                if (!string.IsNullOrWhiteSpace(bitness))
                    return bitness; // "x86" or "x64"
            }

            // Fall back to Access-specific keys.
            // 64-bit installs write to the native registry hive;
            // 32-bit installs on 64-bit Windows write under WOW6432Node.
            var foundNative = hklm64.OpenSubKey(
                $@"SOFTWARE\Microsoft\Office\{officeVersion}\Access\InstallRoot") != null;
            var found32Bit  = hklm64.OpenSubKey(
                $@"SOFTWARE\WOW6432Node\Microsoft\Office\{officeVersion}\Access\InstallRoot") != null;

            if (foundNative && !found32Bit) return "x64";
            if (!foundNative && found32Bit) return "x86";
            if (foundNative && found32Bit)  return "x86"; // both present → mixed install, assume x86
        }
        catch { /* registry unavailable — treated as Not Found */ }

        return "Not Found";
    }

    [McpServerTool(Name = "disconnect_access")]
    [Description("Close the COM connection and quit the Access instance")]
    public Task<string> DisconnectAccess()
        => channel.CallAsync("disconnect_access");

    [McpServerTool(Name = "is_connected")]
    [Description("Check whether a database is currently connected")]
    public Task<string> IsConnected()
        => channel.CallAsync("is_connected");

    // ── Application ──────────────────────────────────────────────────────────────

    [McpServerTool(Name = "get_application_info")]
    [Description("Get basic Access application metadata (name, version, db path, visibility)")]
    public Task<string> GetApplicationInfo()
        => channel.CallAsync("get_application_info");

    [McpServerTool(Name = "get_option")]
    [Description("Get an Access application option value")]
    public Task<string> GetOption(
        [Description("Access option name")] string option_name)
        => channel.CallAsync("get_option", new { option_name });

    [McpServerTool(Name = "set_option")]
    [Description("Set an Access application option value")]
    public Task<string> SetOption(
        [Description("Access option name")] string option_name,
        [Description("Option value as string/number/bool")] string value)
        => channel.CallAsync("set_option", new { option_name, value });

    [McpServerTool(Name = "eval_expression")]
    [Description("Evaluate an Access expression via Application.Eval")]
    public Task<string> EvalExpression(
        [Description("Expression text")] string expression)
        => channel.CallAsync("eval_expression", new { expression });

    // ── Procedures & Commands ────────────────────────────────────────────────────

    [McpServerTool(Name = "run_procedure")]
    [Description("Run a VBA procedure/function via Application.Run")]
    public Task<string> RunProcedure(
        [Description("Procedure/function name")] string procedure_name,
        [Description("JSON array of arguments (optional)")] string? arguments_json = null)
        => channel.CallAsync("run_procedure", new { procedure_name, arguments_json });

    [McpServerTool(Name = "run_command")]
    [Description("Run Access DoCmd.RunCommand by command id")]
    public Task<string> RunCommand(
        [Description("Integer command id")] int command_id)
        => channel.CallAsync("run_command", new { command_id });

    [McpServerTool(Name = "invoke_application_method")]
    [Description("Invoke any Application method by name with positional args")]
    public Task<string> InvokeApplicationMethod(
        [Description("Application method name")] string method_name,
        [Description("JSON array of arguments (optional)")] string? arguments_json = null)
        => channel.CallAsync("invoke_application_method", new { method_name, arguments_json });

    [McpServerTool(Name = "invoke_docmd_method")]
    [Description("Invoke any DoCmd method by name with positional args")]
    public Task<string> InvokeDoCmdMethod(
        [Description("DoCmd method name")] string method_name,
        [Description("JSON array of arguments (optional)")] string? arguments_json = null)
        => channel.CallAsync("invoke_docmd_method", new { method_name, arguments_json });

    // ── Generic Object Operations ─────────────────────────────────────────────────

    [McpServerTool(Name = "export_object_to_text")]
    [Description("Export any Access object definition via SaveAsText")]
    public Task<string> ExportObjectToText(
        [Description("Access object type integer (e.g., 2 form, 3 report, 5 query, -32761 module, -32766 macro)")] int object_type,
        [Description("Object name")] string object_name)
        => channel.CallAsync("export_object_to_text", new { object_type, object_name });

    [McpServerTool(Name = "import_object_from_text")]
    [Description("Import/replace any Access object via LoadFromText")]
    public Task<string> ImportObjectFromText(
        [Description("Access object type integer")] int object_type,
        [Description("Object name")] string object_name,
        [Description("Object definition text")] string object_data)
        => channel.CallAsync("import_object_from_text", new { object_type, object_name, object_data });

    [McpServerTool(Name = "delete_object")]
    [Description("Delete any Access object by type and name")]
    public Task<string> DeleteObject(
        [Description("Access object type integer")] int object_type,
        [Description("Object name")] string object_name)
        => channel.CallAsync("delete_object", new { object_type, object_name });

    // ── Enumeration ───────────────────────────────────────────────────────────────

    [McpServerTool(Name = "get_forms")]
    [Description("List all forms in the database")]
    public Task<string> GetForms()
        => channel.CallAsync("get_forms");

    [McpServerTool(Name = "get_reports")]
    [Description("List all reports in the database")]
    public Task<string> GetReports()
        => channel.CallAsync("get_reports");

    [McpServerTool(Name = "get_modules")]
    [Description("List all standard and class VBA modules")]
    public Task<string> GetModules()
        => channel.CallAsync("get_modules");

    [McpServerTool(Name = "get_macros")]
    [Description("List all macros")]
    public Task<string> GetMacros()
        => channel.CallAsync("get_macros");

    [McpServerTool(Name = "get_queries")]
    [Description("List all queries including their SQL text")]
    public Task<string> GetQueries()
        => channel.CallAsync("get_queries");

    [McpServerTool(Name = "get_vba_projects")]
    [Description("List VBA projects and their components")]
    public Task<string> GetVbaProjects()
        => channel.CallAsync("get_vba_projects");

    // ── VBA Code ──────────────────────────────────────────────────────────────────

    [McpServerTool(Name = "get_vba_code")]
    [Description("Get the complete VBA source of a module")]
    public Task<string> GetVbaCode(
        [Description("Module name")] string module_name)
        => channel.CallAsync("get_vba_code", new { module_name });

    [McpServerTool(Name = "set_vba_code")]
    [Description("Replace the entire source of a VBA module")]
    public Task<string> SetVbaCode(
        [Description("Module name")] string module_name,
        [Description("Complete replacement source")] string code)
        => channel.CallAsync("set_vba_code", new { module_name, code });

    [McpServerTool(Name = "add_vba_procedure")]
    [Description("Append a VBA procedure to an existing module")]
    public Task<string> AddVbaProcedure(
        [Description("Module name")] string module_name,
        [Description("Procedure name (reference)")] string procedure_name,
        [Description("Procedure source to append")] string code)
        => channel.CallAsync("add_vba_procedure", new { module_name, procedure_name, code });

    [McpServerTool(Name = "compile_vba")]
    [Description("Compile all VBA modules")]
    public Task<string> CompileVba()
        => channel.CallAsync("compile_vba");

    // ── Forms ─────────────────────────────────────────────────────────────────────

    [McpServerTool(Name = "open_form")]
    [Description("Open a form in design view")]
    public Task<string> OpenForm(
        [Description("Form name")] string form_name)
        => channel.CallAsync("open_form", new { form_name });

    [McpServerTool(Name = "close_form")]
    [Description("Close an open form without saving")]
    public Task<string> CloseForm(
        [Description("Form name")] string form_name)
        => channel.CallAsync("close_form", new { form_name });

    [McpServerTool(Name = "get_form_controls")]
    [Description("List all controls on a form with type and key properties")]
    public Task<string> GetFormControls(
        [Description("Form name")] string form_name)
        => channel.CallAsync("get_form_controls", new { form_name });

    [McpServerTool(Name = "get_control_properties")]
    [Description("Get all properties of a specific control")]
    public Task<string> GetControlProperties(
        [Description("Form name")] string form_name,
        [Description("Control name")] string control_name)
        => channel.CallAsync("get_control_properties", new { form_name, control_name });

    [McpServerTool(Name = "set_control_property")]
    [Description("Set a property on a form control and save the form")]
    public Task<string> SetControlProperty(
        [Description("Form name")] string form_name,
        [Description("Control name")] string control_name,
        [Description("Property name")] string property_name,
        [Description("New value")] string value)
        => channel.CallAsync("set_control_property", new { form_name, control_name, property_name, value });

    [McpServerTool(Name = "export_form_to_text")]
    [Description("Export a form definition via SaveAsText")]
    public Task<string> ExportFormToText(
        [Description("Form name")] string form_name)
        => channel.CallAsync("export_form_to_text", new { form_name });

    [McpServerTool(Name = "import_form_from_text")]
    [Description("Import/replace a form via LoadFromText")]
    public Task<string> ImportFormFromText(
        [Description("Form name")] string form_name,
        [Description("Form definition text")] string form_data)
        => channel.CallAsync("import_form_from_text", new { form_name, form_data });

    [McpServerTool(Name = "delete_form")]
    [Description("Permanently delete a form")]
    public Task<string> DeleteForm(
        [Description("Form name")] string form_name)
        => channel.CallAsync("delete_form", new { form_name });

    // ── Reports ───────────────────────────────────────────────────────────────────

    [McpServerTool(Name = "export_report_to_text")]
    [Description("Export a report definition via SaveAsText")]
    public Task<string> ExportReportToText(
        [Description("Report name")] string report_name)
        => channel.CallAsync("export_report_to_text", new { report_name });

    [McpServerTool(Name = "import_report_from_text")]
    [Description("Import/replace a report via LoadFromText")]
    public Task<string> ImportReportFromText(
        [Description("Report name")] string report_name,
        [Description("Report definition text")] string report_data)
        => channel.CallAsync("import_report_from_text", new { report_name, report_data });

    [McpServerTool(Name = "delete_report")]
    [Description("Permanently delete a report")]
    public Task<string> DeleteReport(
        [Description("Report name")] string report_name)
        => channel.CallAsync("delete_report", new { report_name });

    // ── New Interop Features ──────────────────────────────────────────────────────

    [McpServerTool(Name = "get_table_definitions")]
    [Description("Get detailed definitions of all tables with field information (Interop feature)")]
    public Task<string> GetTableDefinitions()
        => channel.CallAsync("get_table_definitions");

    [McpServerTool(Name = "get_table_definition")]
    [Description("Get detailed definition of a specific table with field information (Interop feature)")]
    public Task<string> GetTableDefinition(
        [Description("Table name")] string table_name)
        => channel.CallAsync("get_table_definition", new { table_name });

    [McpServerTool(Name = "get_objects_by_type")]
    [Description("Get objects filtered by type using enum-based type selection (Interop feature)")]
    public Task<string> GetObjectsByType(
        [Description("Object type enum: acTable (0), acQuery (1), acForm (2), acReport (3), acMacro (4), acModule (5)")] int object_type)
        => channel.CallAsync("get_objects_by_type", new { object_type });

    [McpServerTool(Name = "compile_vba_with_errors")]
    [Description("Compile VBA and return detailed error information (Interop feature)")]
    public Task<string> CompileVbaWithErrors()
        => channel.CallAsync("compile_vba_with_errors");

    [McpServerTool(Name = "get_database_summary")]
    [Description("Get summary of all database objects (tables, forms, reports, queries, modules, macros) (Interop feature)")]
    public Task<string> GetDatabaseObjectsSummary()
        => channel.CallAsync("get_database_summary");
}
