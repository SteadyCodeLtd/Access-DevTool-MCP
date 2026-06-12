// Access-DevTool-MCP — Access.Application COM Automation MCP Server
//
// Exposes forms, reports, modules, macros, queries (with SQL), VBA code
// read/write, form controls, and form/report SaveAsText/LoadFromText —
// everything the ADODB server cannot reach.
//
// All COM calls are dispatched on a dedicated STA thread as required by
// Access.Application. Pass the database path as the first CLI argument
// for automatic connection at startup.

using System.Text.Json;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace AccessDevToolAgent
{

    internal class Program
    {
        static readonly AccessComService _svc = new AccessComService();
        static string _engineValidationError;
        static List<string> _providerWarnings = new List<string>();

        // ── Known 32-bit OLEDB providers that must have their ProgID in WOW6432Node ──
        static readonly (string ProgId, string Clsid, string Description)[] _knownProviders = new[]
        {
            ("Microsoft.ACE.OLEDB.16.0", "{3BE786A2-0366-4F5C-9434-25CF162E475E}", "Access Database Engine 2016 (required for .accdb/.xlsx linked tables)"),
            ("MSDASQL.1",               "{c8b522cb-5cf3-11ce-ade5-00aa0044773d}", "OLE DB Provider for ODBC (required for ODBC-bridged connections)"),
            ("SQLNCLI11.1",             "{397C2819-8272-4532-AD3A-FB5E43BEAA39}", "SQL Server Native Client 11.0"),
        };

        static async Task<int> Main(string[] args)
        {
            var tcs = new TaskCompletionSource<int>();
            var sta = new Thread(() =>
            {
                try { tcs.SetResult(RunServer(args).GetAwaiter().GetResult()); }
                catch (Exception ex) { Console.Error.WriteLine($"Fatal: {ex.Message}"); tcs.SetResult(1); }
            });
            sta.SetApartmentState(ApartmentState.STA);
            sta.Start();
            return await tcs.Task;
        }

        static async Task<int> RunServer(string[] args)
        {
            var accessBitness = GetAccessBitness();
            _engineValidationError = ValidateDatabaseEngine(accessBitness);
            if (!string.IsNullOrWhiteSpace(_engineValidationError))
                Console.Error.WriteLine(_engineValidationError);

            // Named pipe worker mode: Access-DevTool-Agent.exe --pipe <name> [optional db path]
            var pipeIdx = Array.IndexOf(args, "--pipe");
            if (pipeIdx >= 0 && pipeIdx + 1 < args.Length)
            {
                var pipeName = args[pipeIdx + 1];
                foreach (var arg in args)
                {
                    if (arg != "--pipe" && arg != pipeName && File.Exists(arg))
                    {
                        try
                        {
                            EnsureEngineReadyForConnect();
                            _svc.Connect(arg);
                            Console.Error.WriteLine($"[Worker] Auto-connected to {arg}");
                        }
                        catch (Exception ex) { Console.Error.WriteLine($"[Worker] Auto-connect failed: {ex.Message}"); }
                        break;
                    }
                }
                return RunPipeWorkerSync(pipeName);
            }

            if (args.Length > 0 && File.Exists(args[0]))
            {
                try
                {
                    EnsureEngineReadyForConnect();
                    _svc.Connect(args[0]);
                    Console.Error.WriteLine($"Auto-connected to {args[0]}");
                }
                catch (Exception ex) { Console.Error.WriteLine($"Auto-connect failed: {ex.Message}"); }
            }

            Console.Error.WriteLine($"Detected Access bitness: {accessBitness}");

            _providerWarnings = CheckProviderRegistrations();
            foreach (var w in _providerWarnings)
                Console.Error.WriteLine("[provider-check] " + w);

            string line;
            while ((line = await Console.In.ReadLineAsync()) != null)
            {
                try
                {
                    using (var doc = JsonDocument.Parse(line))
                    {
                        var root = doc.RootElement;

                        JsonElement methodEl;
                        if (!root.TryGetProperty("method", out methodEl)) continue;
                        var method = methodEl.GetString();
                        if (string.IsNullOrEmpty(method)) continue;

                        JsonElement idEl;
                        if (!root.TryGetProperty("id", out idEl)) continue; // notification — no response

                        JsonElement paramsEl;
                        root.TryGetProperty("params", out paramsEl);

                        object result;
                        switch (method)
                        {
                            case "initialize":
                                result = HandleInitialize();
                                break;
                            case "tools/list":
                                result = HandleToolsList();
                                break;
                            case "tools/call":
                                result = HandleToolsCall(paramsEl);
                                break;
                            default:
                                result = new { error = $"Unknown method: {method}" };
                                break;
                        }

                        Console.WriteLine(JsonSerializer.Serialize(new { jsonrpc = "2.0", id = idEl, result }));
                    }
                }
                catch (Exception ex) { Console.Error.WriteLine($"Request error: {ex.Message}"); }
            }

            _svc.Dispose();
            return 0;
        }

        // ── MCP Protocol ──────────────────────────────────────────────────────────

        static object HandleInitialize() => new
        {
            protocolVersion = "2024-11-05",
            capabilities = new { },
            serverInfo = new { name = "MS.Access.MCP.COM", version = "1.0.0" }
        };

        static object HandleToolsList() => new
        {
            tools = new object[]
            {
            T("connect_access",           "Connect to an Access database via COM automation (required before all other tools)",
              P("database_path", "Full path to the .accdb or .mdb file")),
            T("get_access_bitness",       "Get the bitness of the local Microsoft Access installation without opening Access"),
            T("disconnect_access",        "Close the COM connection and quit the Access instance"),
            T("is_connected",             "Check whether a database is currently connected"),
            T("get_application_info",     "Get basic Access application metadata (name, version, db path, visibility)"),
            T("get_option",               "Get an Access application option value",
              P("option_name", "Access option name")),
            T("set_option",               "Set an Access application option value",
              P("option_name", "Access option name"), P("value", "Option value as string/number/bool")),
            T("eval_expression",          "Evaluate an Access expression via Application.Eval",
              P("expression", "Expression text")),
            T("run_procedure",            "Run a VBA procedure/function via Application.Run",
              P("procedure_name", "Procedure/function name"), P("arguments_json", "JSON array of arguments, optional")),
            T("run_command",              "Run Access DoCmd.RunCommand by command id",
              P("command_id", "Integer command id")),
            T("invoke_application_method", "Invoke any Application method by name with positional args",
              P("method_name", "Application method name"), P("arguments_json", "JSON array of arguments, optional")),
            T("invoke_docmd_method",      "Invoke any DoCmd method by name with positional args",
              P("method_name", "DoCmd method name"), P("arguments_json", "JSON array of arguments, optional")),
            T("export_object_to_text",    "Export any Access object definition via SaveAsText",
              P("object_type", "Access object type integer (e.g., 2 form, 3 report, 5 query, -32761 module, -32766 macro)"),
              P("object_name", "Object name")),
            T("import_object_from_text",  "Import/replace any Access object via LoadFromText",
              P("object_type", "Access object type integer"), P("object_name", "Object name"), P("object_data", "Object definition text")),
            T("delete_object",            "Delete any Access object by type and name",
              P("object_type", "Access object type integer"), P("object_name", "Object name")),
            T("get_forms",                "List all forms in the database"),
            T("get_reports",              "List all reports in the database"),
            T("get_modules",              "List all standard and class VBA modules"),
            T("get_macros",               "List all macros"),
            T("get_queries",              "List all queries including their SQL text"),
            T("get_vba_projects",         "List VBA projects and their components"),
            T("get_vba_code",             "Get the complete VBA source of a module",
              P("module_name", "Module name")),
            T("set_vba_code",             "Replace the entire source of a VBA module",
              P("module_name", "Module name"), P("code", "Complete replacement source")),
            T("add_vba_procedure",        "Append a VBA procedure to an existing module",
              P("module_name", "Module name"), P("procedure_name", "Procedure name (reference)"), P("code", "Procedure source to append")),
            T("compile_vba",              "Compile all VBA modules"),
            T("open_form",                "Open a form in design view",
              P("form_name", "Form name")),
            T("close_form",               "Close an open form without saving",
              P("form_name", "Form name")),
            T("get_form_controls",        "List all controls on a form with type and key properties",
              P("form_name", "Form name")),
            T("get_control_properties",   "Get all properties of a specific control",
              P("form_name", "Form name"), P("control_name", "Control name")),
            T("set_control_property",     "Set a property on a form control and save the form",
              P("form_name", "Form name"), P("control_name", "Control name"),
              P("property_name", "Property name"), P("value", "New value")),
            T("export_form_to_text",      "Export a form definition via SaveAsText",
              P("form_name", "Form name")),
            T("import_form_from_text",    "Import/replace a form via LoadFromText",
              P("form_name", "Form name"), P("form_data", "Form definition text")),
            T("export_report_to_text",    "Export a report definition via SaveAsText",
              P("report_name", "Report name")),
            T("import_report_from_text",  "Import/replace a report via LoadFromText",
              P("report_name", "Report name"), P("report_data", "Report definition text")),
            T("delete_form",              "Permanently delete a form",
              P("form_name", "Form name")),
            T("delete_report",            "Permanently delete a report",
              P("report_name", "Report name")),
            T("export_database_objects",   "Export forms, reports, queries, and modules (any combination) for backup",
              P("object_types", "JSON object with keys forms/reports/queries/modules. Each value is [] for all or [\"Name1\",\"Name2\"] for specific objects")),
            }
        };

        static object HandleToolsCall(JsonElement p)
        {
            var name = p.GetProperty("name").GetString() ?? "";
            var args = p.TryGetProperty("arguments", out var a) ? a : default(JsonElement);

            object result;
            switch (name)
            {
                case "connect_access":
                    result = Wrap(() =>
                    {
                        EnsureEngineReadyForConnect();
                        _svc.Connect(S(args, "database_path"));
                        if (_providerWarnings.Count > 0)
                            return new { success = true, message = "Connected", connected = true, warnings = _providerWarnings.ToArray() };
                        return (object)new { success = true, message = "Connected", connected = true };
                    });
                    break;
                case "get_access_bitness":
                    result = new { success = true, bitness = GetAccessBitness() };
                    break;
                case "disconnect_access":
                    result = Wrap(() => { _svc.Disconnect(); return new { success = true, message = "Disconnected" }; });
                    break;
                case "is_connected":
                    result = new { success = true, connected = _svc.IsConnected };
                    break;
                case "get_application_info":
                    result = Wrap(() => new { success = true, info = _svc.GetApplicationInfo() });
                    break;
                case "get_option":
                    result = Wrap(() => new { success = true, value = _svc.GetOption(S(args, "option_name")) });
                    break;
                case "set_option":
                    result = Wrap(() => { _svc.SetOption(S(args, "option_name"), S(args, "value")); return new { success = true, message = "Option set" }; });
                    break;
                case "eval_expression":
                    result = Wrap(() => new { success = true, value = _svc.EvalExpression(S(args, "expression")) });
                    break;
                case "run_procedure":
                    result = Wrap(() => new { success = true, value = _svc.RunProcedure(S(args, "procedure_name"), ParseArgsArray(args, "arguments_json")) });
                    break;
                case "run_command":
                    result = Wrap(() => { _svc.RunCommand(ParseIntArg(args, "command_id")); return new { success = true, message = "Command executed" }; });
                    break;
                case "invoke_application_method":
                    result = Wrap(() => new { success = true, value = _svc.InvokeApplicationMethod(S(args, "method_name"), ParseArgsArray(args, "arguments_json")) });
                    break;
                case "invoke_docmd_method":
                    result = Wrap(() => new { success = true, value = _svc.InvokeDoCmdMethod(S(args, "method_name"), ParseArgsArray(args, "arguments_json")) });
                    break;
                case "export_object_to_text":
                    result = Wrap(() => new { success = true, object_data = _svc.ExportObjectToText(ParseIntArg(args, "object_type"), S(args, "object_name")) });
                    break;
                case "import_object_from_text":
                    result = Wrap(() => { _svc.ImportObjectFromText(ParseIntArg(args, "object_type"), S(args, "object_name"), S(args, "object_data")); return new { success = true, message = "Imported" }; });
                    break;
                case "delete_object":
                    result = Wrap(() => { _svc.DeleteObject(ParseIntArg(args, "object_type"), S(args, "object_name")); return new { success = true, message = "Deleted" }; });
                    break;
                case "get_forms":
                    result = Wrap(() => new { success = true, forms = _svc.GetForms() });
                    break;
                case "get_reports":
                    result = Wrap(() => new { success = true, reports = _svc.GetReports() });
                    break;
                case "get_modules":
                    result = Wrap(() => new { success = true, modules = _svc.GetModules() });
                    break;
                case "get_macros":
                    result = Wrap(() => new { success = true, macros = _svc.GetMacros() });
                    break;
                case "get_queries":
                    result = Wrap(() => new { success = true, queries = _svc.GetQueries() });
                    break;
                case "get_vba_projects":
                    result = Wrap(() => new { success = true, projects = _svc.GetVbaProjects() });
                    break;
                case "get_vba_code":
                    result = Wrap(() => new { success = true, code = _svc.GetVbaCode(S(args, "module_name")) });
                    break;
                case "set_vba_code":
                    result = Wrap(() => { _svc.SetVbaCode(S(args, "module_name"), S(args, "code")); return new { success = true, message = "Code written" }; });
                    break;
                case "add_vba_procedure":
                    result = Wrap(() => { _svc.AddVbaProcedure(S(args, "module_name"), S(args, "code")); return new { success = true, message = "Procedure appended" }; });
                    break;
                case "compile_vba":
                    result = Wrap(() => { _svc.CompileVba(); return new { success = true, message = "Compiled" }; });
                    break;
                case "open_form":
                    result = Wrap(() => { _svc.OpenForm(S(args, "form_name")); return new { success = true, message = "Opened" }; });
                    break;
                case "close_form":
                    result = Wrap(() => { _svc.CloseForm(S(args, "form_name")); return new { success = true, message = "Closed" }; });
                    break;
                case "get_form_controls":
                    result = Wrap(() => new { success = true, controls = _svc.GetFormControls(S(args, "form_name")) });
                    break;
                case "get_control_properties":
                    result = Wrap(() => new { success = true, properties = _svc.GetControlProperties(S(args, "form_name"), S(args, "control_name")) });
                    break;
                case "set_control_property":
                    result = Wrap(() => { _svc.SetControlProperty(S(args, "form_name"), S(args, "control_name"), S(args, "property_name"), S(args, "value")); return new { success = true, message = "Property set" }; });
                    break;
                case "export_form_to_text":
                    result = Wrap(() => new { success = true, form_data = _svc.ExportFormToText(S(args, "form_name")) });
                    break;
                case "import_form_from_text":
                    result = Wrap(() => { _svc.ImportFormFromText(S(args, "form_name"), S(args, "form_data")); return new { success = true, message = "Imported" }; });
                    break;
                case "export_report_to_text":
                    result = Wrap(() => new { success = true, report_data = _svc.ExportReportToText(S(args, "report_name")) });
                    break;
                case "import_report_from_text":
                    result = Wrap(() => { _svc.ImportReportFromText(S(args, "report_name"), S(args, "report_data")); return new { success = true, message = "Imported" }; });
                    break;
                case "delete_form":
                    result = Wrap(() => { _svc.DeleteForm(S(args, "form_name")); return new { success = true, message = "Deleted" }; });
                    break;
                case "delete_report":
                    result = Wrap(() => { _svc.DeleteReport(S(args, "report_name")); return new { success = true, message = "Deleted" }; });
                    break;
                case "export_database_objects":
                    {
                        JsonElement objectTypesEl;
                        if (!args.TryGetProperty("object_types", out objectTypesEl) || objectTypesEl.ValueKind == JsonValueKind.Null || objectTypesEl.ValueKind == JsonValueKind.Undefined)
                        {
                            result = new { success = false, error = "Required parameter 'object_types' missing" };
                            break;
                        }

                        var objectTypesDict = new Dictionary<int, List<string>>();

                        JsonElement formsArg;
                        if (objectTypesEl.TryGetProperty("forms", out formsArg) && formsArg.ValueKind != JsonValueKind.Null && formsArg.ValueKind != JsonValueKind.Undefined)
                            objectTypesDict[2] = JsonToList(formsArg); // 2 = acForm

                        JsonElement reportsArg;
                        if (objectTypesEl.TryGetProperty("reports", out reportsArg) && reportsArg.ValueKind != JsonValueKind.Null && reportsArg.ValueKind != JsonValueKind.Undefined)
                            objectTypesDict[3] = JsonToList(reportsArg); // 3 = acReport

                        JsonElement queriesArg;
                        if (objectTypesEl.TryGetProperty("queries", out queriesArg) && queriesArg.ValueKind != JsonValueKind.Null && queriesArg.ValueKind != JsonValueKind.Undefined)
                            objectTypesDict[1] = JsonToList(queriesArg); // 1 = acQuery

                        JsonElement modulesArg;
                        if (objectTypesEl.TryGetProperty("modules", out modulesArg) && modulesArg.ValueKind != JsonValueKind.Null && modulesArg.ValueKind != JsonValueKind.Undefined)
                            objectTypesDict[5] = JsonToList(modulesArg); // 5 = acModule

                        if (objectTypesDict.Count == 0)
                        {
                            result = new { success = false, error = "At least one object type must be specified (forms, reports, queries, modules)." };
                            break;
                        }

                        result = Wrap(() =>
                        {
                            var exportResult = _svc.ExportDatabaseObjects(objectTypesDict);
                            return new
                            {
                                success = exportResult.Success,
                                message = exportResult.Message,
                                exported_objects = exportResult.ExportedObjects.ConvertAll(o => new { type = o.Type, name = o.Name, code = o.Code }),
                                errors = exportResult.Errors
                            };
                        });
                    }
                    break;
                default:
                    result = (object)new { success = false, error = $"Unknown tool: {name}" };
                    break;
            }

            return new { content = new[] { new { type = "text", text = JsonSerializer.Serialize(result) } } };
        }

        // ── Helpers ───────────────────────────────────────────────────────────────

        static void EnsureEngineReadyForConnect()
        {
            if (!string.IsNullOrWhiteSpace(_engineValidationError))
                throw new InvalidOperationException(_engineValidationError);
        }

        static string ValidateDatabaseEngine(string accessBitness)
        {
            if (string.Equals(accessBitness, "x86", StringComparison.OrdinalIgnoreCase))
            {
                if (!HasAceProvider(RegistryView.Registry32))
                    return "Access is x86, but the 32-bit Access Database Engine (Microsoft.ACE.OLEDB.12.0/16.0) is not installed.";

                return null;
            }

            return null;
        }

        /// <summary>
        /// Checks each known OLEDB provider for the 32-bit ProgID split caused by
        /// Click-to-Run Office: the CLSID + DLL are registered but the WOW6432Node
        /// ProgID entry is missing, which causes error 3706 in any 32-bit ADO/DAO code
        /// that references the provider by name.
        /// </summary>
        static List<string> CheckProviderRegistrations()
        {
            var issues = new List<string>();
            try
            {
                using (var wow32 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32))
                {
                    foreach (var (progId, clsid, description) in _knownProviders)
                    {
                        // Is the ProgID already registered for 32-bit processes?
                        using (var progIdKey = wow32.OpenSubKey($@"SOFTWARE\Classes\{progId}"))
                        {
                            if (progIdKey != null) continue; // already fine
                        }

                        // ProgID is missing — check whether the CLSID + DLL are present
                        // so we can give a precise, actionable fix.
                        string dllPath = null;
                        using (var clsidKey = wow32.OpenSubKey($@"SOFTWARE\Classes\CLSID\{clsid}\InprocServer32"))
                        {
                            if (clsidKey != null)
                                dllPath = Environment.ExpandEnvironmentVariables(clsidKey.GetValue(null)?.ToString() ?? "");
                        }

                        if (!string.IsNullOrEmpty(dllPath) && File.Exists(dllPath))
                        {
                            // DLL present but ProgID missing — the two-key registry fix applies.
                            issues.Add(
                                $"OLEDB provider '{progId}' ({description}) is not registered for 32-bit processes. " +
                                $"If your database uses this provider you will see error 3706 dialogs. " +
                                $"Fix (run as Administrator):\n" +
                                $"  reg add \"HKLM\\SOFTWARE\\WOW6432Node\\Classes\\{progId}\" /ve /d \"{progId}\" /f\n" +
                                $"  reg add \"HKLM\\SOFTWARE\\WOW6432Node\\Classes\\{progId}\\CLSID\" /ve /d \"{clsid}\" /f");
                        }
                        else
                        {
                            // Neither ProgID nor CLSID/DLL — provider not installed at all.
                            issues.Add(
                                $"OLEDB provider '{progId}' ({description}) is not installed. " +
                                $"Install the provider package to resolve potential error 3706 dialogs.");
                        }
                    }
                }
            }
            catch { /* best-effort */ }
            return issues;
        }

        static bool HasAceProvider(RegistryView view)
        {
            try
            {
                using (var root = RegistryKey.OpenBaseKey(RegistryHive.ClassesRoot, view))
                {
                    return root.OpenSubKey(@"Microsoft.ACE.OLEDB.16.0") != null
                        || root.OpenSubKey(@"Microsoft.ACE.OLEDB.12.0") != null;
                }
            }
            catch
            {
                return false;
            }
        }

        static string S(JsonElement el, string key) =>
            el.TryGetProperty(key, out var v) ? v.GetString() ?? "" : throw new ArgumentException($"'{key}' is required");

        static object T(string name, string desc, params (string name, string desc)[] props)
        {
            var propSchema = new Dictionary<string, object>();
            foreach (var prop in props)
                propSchema[prop.name] = new { type = "string", description = prop.desc };
            var req = Array.ConvertAll(props, p => p.name);
            return new { name, description = desc, inputSchema = new { type = "object", properties = propSchema, required = req } };
        }

        static (string name, string desc) P(string name, string desc) => (name, desc);

        static int ParseIntArg(JsonElement el, string key)
        {
            JsonElement v;
            if (!el.TryGetProperty(key, out v)) throw new ArgumentException($"'{key}' is required");
            if (v.ValueKind == JsonValueKind.Number)
            {
                int n;
                if (v.TryGetInt32(out n)) return n;
                throw new ArgumentException($"'{key}' must be an integer");
            }
            string s = v.GetString() ?? "";
            int i;
            if (!int.TryParse(s, out i)) throw new ArgumentException($"'{key}' must be an integer");
            return i;
        }

        static object[] ParseArgsArray(JsonElement el, string key)
        {
            JsonElement arr;
            if (!el.TryGetProperty(key, out arr) || arr.ValueKind == JsonValueKind.Null || arr.ValueKind == JsonValueKind.Undefined)
                return new object[0];

            if (arr.ValueKind == JsonValueKind.String)
            {
                var text = arr.GetString();
                if (string.IsNullOrWhiteSpace(text)) return new object[0];
                using (var doc = JsonDocument.Parse(text))
                {
                    if (doc.RootElement.ValueKind != JsonValueKind.Array)
                        throw new ArgumentException($"'{key}' must be a JSON array");
                    return ToObjectArray(doc.RootElement);
                }
            }

            if (arr.ValueKind != JsonValueKind.Array)
                throw new ArgumentException($"'{key}' must be a JSON array");

            return ToObjectArray(arr);
        }

        static List<string> JsonToList(JsonElement el)
        {
            if (el.ValueKind == JsonValueKind.Array)
            {
                var list = new List<string>();
                foreach (var item in el.EnumerateArray())
                    list.Add(item.GetString() ?? "");
                return list;
            }

            if (el.ValueKind == JsonValueKind.String)
            {
                var text = el.GetString();
                if (string.IsNullOrWhiteSpace(text)) return new List<string>();
                using (var doc = JsonDocument.Parse(text))
                {
                    if (doc.RootElement.ValueKind != JsonValueKind.Array)
                        throw new ArgumentException("Object type filter value must be a JSON array or array string");

                    var list = new List<string>();
                    foreach (var item in doc.RootElement.EnumerateArray())
                        list.Add(item.GetString() ?? "");
                    return list;
                }
            }

            throw new ArgumentException("Object type filter value must be an array or stringified JSON array");
        }

        static object[] ToObjectArray(JsonElement arr)
        {
            var list = new List<object>();
            foreach (var item in arr.EnumerateArray())
                list.Add(J(item));
            return list.ToArray();
        }

        static object J(JsonElement e)
        {
            switch (e.ValueKind)
            {
                case JsonValueKind.String:
                    return e.GetString() ?? "";
                case JsonValueKind.Number:
                    int i;
                    if (e.TryGetInt32(out i)) return i;
                    long l;
                    if (e.TryGetInt64(out l)) return l;
                    double d;
                    if (e.TryGetDouble(out d)) return d;
                    return e.GetRawText();
                case JsonValueKind.True:
                case JsonValueKind.False:
                    return e.GetBoolean();
                case JsonValueKind.Null:
                case JsonValueKind.Undefined:
                    return null;
                default:
                    return e.GetRawText();
            }
        }

        public static string GetAccessBitness(string officeVersion = "16.0")
        {
            // Check the 64-bit registry view first
            using (var hklm64 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
            {
                var outlookKey = hklm64.OpenSubKey($@"SOFTWARE\Microsoft\Office\{officeVersion}\Outlook");
                if (outlookKey != null)
                {
                    var bitness = outlookKey.GetValue("Bitness")?.ToString();
                    if (!string.IsNullOrWhiteSpace(bitness))
                        return bitness; // Returns "x86" or "x64"
                }

                var foundAccessKey = false;
                var found32BitAccessKey = false;
                var accessJet = hklm64.OpenSubKey($@"SOFTWARE\Microsoft\Office\{officeVersion}\Access\InstallRoot");
                if (accessJet != null)
                {
                    foundAccessKey = true;
                }
                var access32BitJet = hklm64.OpenSubKey($@"SOFTWARE\WOW6432Node\Microsoft\Office\{officeVersion}\Access\InstallRoot");
                if (access32BitJet != null)
                {
                    found32BitAccessKey = true;
                }

                if (foundAccessKey && !found32BitAccessKey)
                {
                    return "x64";
                }
                else if (!foundAccessKey && found32BitAccessKey)
                {
                    return "x86";
                }
            }
            return "Not Found";
        }

        static object Wrap(Func<object> fn)
        {
            try { return fn(); }
            catch (Exception ex) { return new { success = false, error = ex.Message }; }
        }

        // ── Named Pipe Worker Mode ────────────────────────────────────────────────

        static int RunPipeWorkerSync(string pipeName)
        {
            Console.Error.WriteLine($"[Worker] Starting pipe server: {pipeName}");
            var running = true;
            Console.CancelKeyPress += (s, e) => { e.Cancel = true; running = false; };

            while (running)
            {
                NamedPipeServerStream pipe = null;
                try
                {
                    pipe = new NamedPipeServerStream(pipeName, PipeDirection.InOut, 1,
                        PipeTransmissionMode.Byte, PipeOptions.None);
                    pipe.WaitForConnection();
                    Console.Error.WriteLine("[Worker] MCP client connected.");

                    using (var reader = new StreamReader(pipe, new UTF8Encoding(false), false, 65536, true))
                    using (var writer = new StreamWriter(pipe, new UTF8Encoding(false), 65536, true) { AutoFlush = true })
                    {
                        string line;
                        while ((line = reader.ReadLine()) != null)
                        {
                            string response;
                            try { response = DispatchPipeRequest(line); }
                            catch (Exception ex) { response = JsonSerializer.Serialize(new { success = false, error = ex.Message }); }
                            writer.WriteLine(response);
                        }
                    }
                    Console.Error.WriteLine("[Worker] MCP client disconnected.");
                }
                catch (Exception ex)
                {
                    if (running) Console.Error.WriteLine($"[Worker] Pipe error: {ex.Message}");
                }
                finally
                {
                    pipe?.Dispose();
                }
            }

            _svc.Dispose();
            return 0;
        }

        static string DispatchPipeRequest(string requestJson)
        {
            string toolName;
            string paramsRaw;

            using (var doc = JsonDocument.Parse(requestJson))
            {
                var root = doc.RootElement;
                toolName = root.GetProperty("tool").GetString() ?? "";
                JsonElement paramsEl;
                paramsRaw = (root.TryGetProperty("params", out paramsEl) && paramsEl.ValueKind == JsonValueKind.Object)
                    ? paramsEl.GetRawText()
                    : "{}";
            }

            var mcpParamsJson = "{\"name\":" + JsonSerializer.Serialize(toolName) + ",\"arguments\":" + paramsRaw + "}";
            using (var mcpDoc = JsonDocument.Parse(mcpParamsJson))
            {
                var mcpResult = HandleToolsCall(mcpDoc.RootElement);
                var envelope = JsonSerializer.Serialize(mcpResult);
                using (var envDoc = JsonDocument.Parse(envelope))
                {
                    return envDoc.RootElement
                        .GetProperty("content")[0]
                        .GetProperty("text")
                        .GetString() ?? "{}";
                }
            }
        }
    }
}