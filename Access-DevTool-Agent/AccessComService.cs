using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace AccessDevToolAgent
{
    // All public methods must be called from an STA thread.
    // In the MCP server the entire message loop runs on a dedicated STA thread.
    // In test projects use [StaFact] from Xunit.StaFact.
    public sealed class AccessComService : IDisposable
    {
        private dynamic _app;
        private Process _accessProcess;
        private string  _databasePath = "";

        public bool IsConnected => _app != null;

        public static readonly Dictionary<int, string> ControlTypeNames = new Dictionary<int, string>()
        {
            [100] = "Label",            [101] = "Rectangle",        [102] = "Line",
            [103] = "Image",            [104] = "CommandButton",    [105] = "OptionButton",
            [106] = "CheckBox",         [107] = "OptionGroup",      [108] = "BoundObjectFrame",
            [109] = "TextBox",          [110] = "ListBox",          [111] = "ComboBox",
            [112] = "Subform",          [114] = "UnboundObjectFrame",
            [118] = "PageBreak",        [119] = "CustomControl",    [122] = "ToggleButton",
            [123] = "TabControl",       [124] = "Page",             [129] = "NavigationControl",
            [130] = "NavigationButton"
        };

        // ── Connection ────────────────────────────────────────────────────────

        public void Connect(string databasePath)
        {
            if (!File.Exists(databasePath))
                throw new FileNotFoundException("Database file not found.", databasePath);

            Disconnect();

            var type = Type.GetTypeFromProgID("Access.Application")
                ?? throw new InvalidOperationException("Microsoft Access is not installed on this machine.");

            _databasePath = databasePath;

            // Snapshot existing MSACCESS pids before creating so we can identify the new one.
            var before = new HashSet<int>(Process.GetProcessesByName("MSACCESS").Select(p => p.Id));

            _app = Activator.CreateInstance(type);
            _app.UserControl = false;
            _app.Visible = false;
            TryForceDisableAutomationMacros(_app);

            _app.OpenCurrentDatabase(databasePath, false); // shared mode
            WaitForDatabaseReady(_app);

            // Identify the newly spawned Access process.
            try
            {
                var after = Process.GetProcessesByName("MSACCESS");
                _accessProcess = after.FirstOrDefault(p => !before.Contains(p.Id));
            }
            catch { _accessProcess = null; }
        }

        public void Disconnect() => DisconnectInternal();

        public void Dispose() => Disconnect();

        // ── Object enumeration ────────────────────────────────────────────────

        public List<AccessObjectInfo> GetForms()
        {
            var app  = EnsureApp();
            var list = new List<AccessObjectInfo>();
            dynamic all = app.CurrentProject.AllForms;
            for (int i = 0; i < (int)all.Count; i++)
            {
                dynamic f = null;
                try { f = all.Item(i); }
                catch
                {
                    try { f = all.Item(i + 1); } catch { }
                }
                if (f == null) continue;
                list.Add(new AccessObjectInfo { Name = (string)f.Name, IsLoaded = (bool)f.IsLoaded });
            }
            return list;
        }

        public List<AccessObjectInfo> GetReports()
        {
            var app  = EnsureApp();
            var list = new List<AccessObjectInfo>();
            dynamic all = app.CurrentProject.AllReports;
            for (int i = 0; i < (int)all.Count; i++)
            {
                dynamic r = null;
                try { r = all.Item(i); }
                catch
                {
                    try { r = all.Item(i + 1); } catch { }
                }
                if (r == null) continue;
                list.Add(new AccessObjectInfo { Name = (string)r.Name, IsLoaded = (bool)r.IsLoaded });
            }
            return list;
        }

        public List<AccessObjectInfo> GetModules()
        {
            var app  = EnsureApp();
            var list = new List<AccessObjectInfo>();
            dynamic all = app.CurrentProject.AllModules;
            for (int i = 0; i < (int)all.Count; i++)
            {
                dynamic m = null;
                try { m = all.Item(i); }
                catch
                {
                    try { m = all.Item(i + 1); } catch { }
                }
                if (m == null) continue;
                list.Add(new AccessObjectInfo { Name = (string)m.Name, IsLoaded = (bool)m.IsLoaded });
            }
            return list;
        }

        public List<AccessObjectInfo> GetMacros()
        {
            var app  = EnsureApp();
            var list = new List<AccessObjectInfo>();
            dynamic all = app.CurrentProject.AllMacros;
            for (int i = 0; i < (int)all.Count; i++)
            {
                dynamic m = null;
                try { m = all.Item(i); }
                catch
                {
                    try { m = all.Item(i + 1); } catch { }
                }
                if (m == null) continue;
                list.Add(new AccessObjectInfo { Name = (string)m.Name, IsLoaded = false });
            }
            return list;
        }

        public List<QueryInfo> GetQueries()
        {
            var app  = EnsureApp();
            var list = new List<QueryInfo>();
            dynamic db   = app.CurrentDb();
            dynamic defs = db.QueryDefs;
            for (int i = 0; i < (int)defs.Count; i++)
            {
                dynamic qd = null;
                try { qd = defs.Item(i); }
                catch
                {
                    try { qd = defs.Item(i + 1); } catch { }
                }

                if (qd == null) continue;

                try
                {
                    string qName = (string)qd.Name;
                    if (!qName.StartsWith("~"))
                    {
                        string sql = "";
                        try { sql = (string)qd.SQL; } catch { }
                        list.Add(new QueryInfo { Name = qName, Sql = sql });
                    }
                }
                finally
                {
                    if (Marshal.IsComObject(qd))
                    {
                        try { Marshal.ReleaseComObject(qd); } catch { }
                    }
                }
            }
            Marshal.ReleaseComObject(defs);
            Marshal.ReleaseComObject(db);
            return list;
        }

        // ── VBA ───────────────────────────────────────────────────────────────

        public List<VbaProjectInfo> GetVbaProjects()
        {
            var app      = EnsureApp();
            dynamic vbe  = app.VBE;
            var projects = new List<VbaProjectInfo>();
            dynamic vbps = vbe.VBProjects;
            for (int i = 1; i <= (int)vbps.Count; i++)
            {
                dynamic proj  = vbps.Item(i);
                var components = new List<VbaComponentInfo>();
                dynamic comps  = proj.VBComponents;
                for (int j = 1; j <= (int)comps.Count; j++)
                {
                    dynamic comp = comps.Item(j);
                    int     kind = (int)comp.Type;
                    components.Add(new VbaComponentInfo
                    {
                        Name = (string)comp.Name,
                        Type = GetVbaComponentTypeName(kind)
                    });
                }
                projects.Add(new VbaProjectInfo { Name = (string)proj.Name, Components = components });
            }
            return projects;
        }

        public string GetVbaCode(string moduleName)
        {
            var app       = EnsureApp();
            dynamic vbe   = app.VBE;
            dynamic proj  = vbe.VBProjects.Item(1);
            dynamic comps = proj.VBComponents;
            for (int i = 1; i <= (int)comps.Count; i++)
            {
                dynamic comp = comps.Item(i);
                if (!string.Equals((string)comp.Name, moduleName, StringComparison.OrdinalIgnoreCase)) continue;
                dynamic cm    = comp.CodeModule;
                int     lines = (int)cm.CountOfLines;
                return lines > 0 ? (string)cm.Lines(1, lines) : "";
            }
            throw new ArgumentException($"Module '{moduleName}' not found.");
        }

        public void SetVbaCode(string moduleName, string code)
        {
            var app       = EnsureApp();
            dynamic vbe   = app.VBE;
            dynamic proj  = vbe.VBProjects.Item(1);
            dynamic comps = proj.VBComponents;
            for (int i = 1; i <= (int)comps.Count; i++)
            {
                dynamic comp = comps.Item(i);
                if (!string.Equals((string)comp.Name, moduleName, StringComparison.OrdinalIgnoreCase)) continue;
                dynamic cm    = comp.CodeModule;
                int     lines = (int)cm.CountOfLines;
                if (lines > 0) cm.DeleteLines(1, lines);
                cm.InsertLines(1, code);
                return;
            }
            throw new ArgumentException($"Module '{moduleName}' not found.");
        }

        public void AddVbaProcedure(string moduleName, string code)
        {
            var app       = EnsureApp();
            dynamic vbe   = app.VBE;
            dynamic proj  = vbe.VBProjects.Item(1);
            dynamic comps = proj.VBComponents;
            for (int i = 1; i <= (int)comps.Count; i++)
            {
                dynamic comp = comps.Item(i);
                if (!string.Equals((string)comp.Name, moduleName, StringComparison.OrdinalIgnoreCase)) continue;
                dynamic cm       = comp.CodeModule;
                int     insertAt = (int)cm.CountOfLines + 1;
                cm.InsertLines(insertAt, "\r\n" + code);
                return;
            }
            throw new ArgumentException($"Module '{moduleName}' not found.");
        }

        public void CompileVba() =>
            EnsureApp().DoCmd.RunCommand(584); // acCmdCompileAllModules

        // ── Forms ─────────────────────────────────────────────────────────────

        public void OpenForm(string formName) =>
            EnsureApp().DoCmd.OpenForm(formName, 1); // 1 = acDesign

        public void CloseForm(string formName) =>
            EnsureApp().DoCmd.Close(2, formName, 2); // acForm=2, acSaveNo=2

        public List<ControlInfo> GetFormControls(string formName)
        {
            var app      = EnsureApp();
            bool wasOpen = IsFormLoaded(app, formName);
            if (!wasOpen) app.DoCmd.OpenForm(formName, 1);

            var list    = new List<ControlInfo>();
            dynamic form = app.Forms.Item(formName);
            dynamic ctls = form.Controls;
            for (int i = 0; i < (int)ctls.Count; i++)
            {
                dynamic c        = ctls.Item(i);
                int     typeCode = (int)c.ControlType;
                list.Add(new ControlInfo
                {
                    Name            = (string)c.Name,
                    ControlType     = typeCode,
                    ControlTypeName = ControlTypeNames.TryGetValue(typeCode, out var typeName) ? typeName : $"Unknown({typeCode})",
                    Caption         = SafeProp(c, "Caption"),
                    ControlSource   = SafeProp(c, "ControlSource"),
                    Left            = SafePropInt(c, "Left"),
                    Top             = SafePropInt(c, "Top"),
                    Width           = SafePropInt(c, "Width"),
                    Height          = SafePropInt(c, "Height"),
                    Visible         = SafePropBool(c, "Visible"),
                    Enabled         = SafePropBool(c, "Enabled"),
                });
            }

            if (!wasOpen) app.DoCmd.Close(2, formName, 2);
            return list;
        }

        public Dictionary<string, string> GetControlProperties(string formName, string controlName)
        {
            var app         = EnsureApp();
            bool wasOpen    = IsFormLoaded(app, formName);
            if (!wasOpen) app.DoCmd.OpenForm(formName, 1);

            dynamic form  = app.Forms.Item(formName);
            dynamic ctrl  = form.Controls.Item(controlName);
            dynamic props = ctrl.Properties;
            var result    = new Dictionary<string, string>();
            for (int i = 0; i < (int)props.Count; i++)
            {
                try
                {
                    dynamic p     = props.Item(i);
                    string  pName = (string)p.Name;
                    string  pVal  = null;
                    try { pVal = p.Value?.ToString(); } catch { }
                    result[pName] = pVal;
                }
                catch { }
            }

            if (!wasOpen) app.DoCmd.Close(2, formName, 2);
            return result;
        }

        public void SetControlProperty(string formName, string controlName, string propertyName, string value)
        {
            var app      = EnsureApp();
            bool wasOpen = IsFormLoaded(app, formName);
            if (!wasOpen) app.DoCmd.OpenForm(formName, 1);

            dynamic form  = app.Forms.Item(formName);
            dynamic ctrl  = form.Controls.Item(controlName);
            dynamic props = ctrl.Properties;
            bool found    = false;
            for (int i = 0; i < (int)props.Count; i++)
            {
                try
                {
                    dynamic p = props.Item(i);
                    if (!string.Equals((string)p.Name, propertyName, StringComparison.OrdinalIgnoreCase)) continue;
                    p.Value = value;
                    found   = true;
                    break;
                }
                catch { }
            }

            if (!found)
                throw new ArgumentException($"Property '{propertyName}' not found on control '{controlName}'.");

            app.DoCmd.Save(2, formName); // acForm=2
            if (!wasOpen) app.DoCmd.Close(2, formName, 1); // acSaveYes
        }

        // ── Generic automation surface ────────────────────────────────────────

        public Dictionary<string, string> GetApplicationInfo()
        {
            var app = EnsureApp();
            var info = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            info["Name"] = SafeAppProp(app, "Name") ?? "";
            info["Version"] = SafeAppProp(app, "Version") ?? "";
            info["CurrentDb"] = _databasePath;
            info["Visible"] = SafeAppProp(app, "Visible") ?? "";
            info["UserControl"] = SafeAppProp(app, "UserControl") ?? "";
            return info;
        }

        public string GetOption(string optionName)
        {
            var app = EnsureApp();
            var value = app.GetOption(optionName);
            return NormalizeOptionValue(value);
        }

        public void SetOption(string optionName, string value)
        {
            var app = EnsureApp();
            app.SetOption(optionName, ParseValue(value));
        }

        public string EvalExpression(string expression)
        {
            var app = EnsureApp();
            var value = app.Eval(expression);
            return value == null ? "" : value.ToString();
        }

        public object RunProcedure(string procedureName, object[] arguments)
        {
            try
            {
                var app = EnsureApp();
                var args = new List<object> { procedureName };
                if (arguments != null && arguments.Length > 0)
                    args.AddRange(arguments);
                return app.GetType().InvokeMember("Run",
                    System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance,
                    null, app, args.ToArray());
            }
            catch (Exception ex)
            {
                throw new Exception(UnwrapMessage(ex));
            }
        }

        public void RunCommand(int commandId)
        {
            try
            {
                EnsureApp().DoCmd.RunCommand(commandId);
            }
            catch (Exception ex)
            {
                throw new Exception(UnwrapMessage(ex));
            }
        }

        public object InvokeApplicationMethod(string methodName, object[] arguments)
        {
            var app = EnsureApp();
            return app.GetType().InvokeMember(methodName,
                System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance,
                null, app, arguments ?? new object[0]);
        }

        public object InvokeDoCmdMethod(string methodName, object[] arguments)
        {
            var doCmd = EnsureApp().DoCmd;
            return doCmd.GetType().InvokeMember(methodName,
                System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance,
                null, doCmd, arguments ?? new object[0]);
        }

        public string ExportObjectToText(int objectType, string objectName)
        {
            var tmp = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".txt");
            try
            {
                EnsureApp().SaveAsText(objectType, objectName, tmp);
                return File.ReadAllText(tmp, Encoding.Default);
            }
            finally { if (File.Exists(tmp)) File.Delete(tmp); }
        }

        public void ImportObjectFromText(int objectType, string objectName, string objectData)
        {
            ImportObjectFromTextExclusive(objectType, objectName, objectData);
        }

        public void DeleteObject(int objectType, string objectName)
        {
            try
            {
                EnsureApp().DoCmd.DeleteObject(objectType, objectName);
            }
            catch (Exception ex)
            {
                throw new Exception(UnwrapMessage(ex));
            }
        }

        // ── Form/report text export/import ────────────────────────────────────

        public string ExportFormToText(string formName)
        {
            return ExportObjectToText(2, formName); // 2 = acForm
        }

        public void ImportFormFromText(string formName, string formData)
        {
            ImportObjectFromTextExclusive(2, formName, formData);
        }

        public string ExportReportToText(string reportName)
        {
            return ExportObjectToText(3, reportName); // 3 = acReport
        }

        public void ImportReportFromText(string reportName, string reportData)
        {
            ImportObjectFromText(3, reportName, reportData); // 3 = acReport
        }

        // ── Destructive ───────────────────────────────────────────────────────

        public void DeleteForm(string formName) =>
            DeleteObject(2, formName); // 2 = acForm

        public void DeleteReport(string reportName) =>
            DeleteObject(3, reportName); // 3 = acReport

        // ── New Interop-Enabled Features ────────────────────────────────────────

        public List<TableDefinitionInfo> GetTableDefinitions()
        {
            var app = EnsureApp();
            var list = new List<TableDefinitionInfo>();

            try
            {
                dynamic db = app.CurrentDb();
                dynamic tableDefs = db.TableDefs;
                int count = (int)tableDefs.Count;

                for (int i = 0; i < count; i++)
                {
                    try
                    {
                        dynamic tdf = tableDefs.Item(i);
                        string tName = (string)tdf.Name;
                        if (tName.StartsWith("MSys")) continue;

                        var fieldList = new List<FieldDefinitionInfo>();
                        dynamic fields = tdf.Fields;
                        int fieldCount = (int)fields.Count;

                        for (int j = 0; j < fieldCount; j++)
                        {
                            try
                            {
                                dynamic fld = fields.Item(j);
                                fieldList.Add(new FieldDefinitionInfo
                                {
                                    Name = (string)fld.Name,
                                    Type = (int)fld.Type,
                                    TypeName = GetFieldTypeName((int)fld.Type),
                                    Size = (int)fld.Size,
                                    Required = (bool)fld.Required,
                                    AllowZeroLength = (bool)fld.AllowZeroLength
                                });
                            }
                            catch { }
                        }

                        list.Add(new TableDefinitionInfo
                        {
                            Name = tName,
                            FieldCount = fieldCount,
                            Fields = fieldList
                        });
                    }
                    catch { }
                }
            }
            catch { }

            return list;
        }

        public TableDefinitionInfo GetTableDefinition(string tableName)
        {
            var tables = GetTableDefinitions();
            return tables.FirstOrDefault(t => string.Equals(t.Name, tableName, StringComparison.OrdinalIgnoreCase))
                ?? throw new ArgumentException($"Table '{tableName}' not found.");
        }

        public List<AccessObjectInfo> GetObjectsByType(int objectType)
        {
            var app = EnsureApp();
            var list = new List<AccessObjectInfo>();

            try
            {
                dynamic objects = null;
                switch (objectType)
                {
                    case 0: objects = app.CurrentData.AllTables; break;         // acTable
                    case 1: objects = app.CurrentData.AllQueries; break;        // acQuery
                    case 2: objects = app.CurrentProject.AllForms; break;       // acForm
                    case 3: objects = app.CurrentProject.AllReports; break;     // acReport
                    case 4: objects = app.CurrentProject.AllMacros; break;      // acMacro
                    case 5: objects = app.CurrentProject.AllModules; break;     // acModule
                    default: throw new ArgumentException($"Unsupported object type: {objectType}");
                }

                for (int i = 0; i < (int)objects.Count; i++)
                {
                    try
                    {
                        dynamic obj = objects.Item(i);
                        list.Add(new AccessObjectInfo { Name = (string)obj.Name, IsLoaded = (bool)obj.IsLoaded });
                    }
                    catch { }
                }
            }
            catch (ArgumentException) { throw; }
            catch { }

            return list;
        }

        public CompileResultInfo CompileVbaWithErrors()
        {
            var app = EnsureApp();
            var result = new CompileResultInfo { Success = false, Errors = new List<string>() };

            try
            {
                try
                {
                    dynamic vbe = app.VBE;
                    app.DoCmd.RunCommand(584); // acCmdCompileAllModules
                    result.Success = true;
                    result.Message = "Compilation successful";
                }
                catch (Exception comEx)
                {
                    result.Success = false;
                    result.Message = comEx.Message;
                    result.Errors.Add($"Compilation Error: {comEx.Message}");
                }
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Message = ex.Message;
                result.Errors.Add(UnwrapMessage(ex));
            }

            return result;
        }

        public DatabaseObjectsSummary GetDatabaseObjectsSummary()
        {
            var app = EnsureApp();
            var summary = new DatabaseObjectsSummary();

            try
            {
                summary.Forms = GetForms().Count;
                summary.Reports = GetReports().Count;
                summary.Modules = GetModules().Count;
                summary.Macros = GetMacros().Count;
                summary.Queries = GetQueries().Count;

                var tables = GetTableDefinitions();
                summary.Tables = tables.Count;
                summary.TotalFields = tables.Sum(t => t.FieldCount);
            }
            catch { }

            return summary;
        }

        private static string GetFieldTypeName(int typeCode)
        {
            switch (typeCode)
            {
                case 1: return "Boolean";
                case 2: return "Byte";
                case 3: return "Integer";
                case 4: return "Long";
                case 5: return "Currency";
                case 6: return "Single";
                case 7: return "Double";
                case 8: return "Date/Time";
                case 9: return "Binary";
                case 10: return "Text";
                case 11: return "Long Binary";
                case 12: return "Memo";
                case 15: return "GUID";
                case 16: return "Numeric";
                case 17: return "Decimal";
                default: return $"Unknown({typeCode})";
            }
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        private void ImportObjectFromTextExclusive(int objectType, string objectName, string objectData)
        {
            EnsureApp();
            var tmp = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName() + ".txt");
            try
            {
                File.WriteAllText(tmp, objectData, Encoding.Default);

                string dbPath = _databasePath;
                DisconnectInternal();

                string lockFile = Path.ChangeExtension(dbPath, ".laccdb");
                for (int i = 0; i < 100 && File.Exists(lockFile); i++)
                    Thread.Sleep(100);

                var type = Type.GetTypeFromProgID("Access.Application")
                    ?? throw new InvalidOperationException("Microsoft Access is not installed on this machine.");
                dynamic exclusiveApp = Activator.CreateInstance(type);
                exclusiveApp.UserControl = false;
                exclusiveApp.Visible = false;
                exclusiveApp.OpenCurrentDatabase(dbPath, true);
                try
                {
                    WaitForDatabaseReady(exclusiveApp);
                    exclusiveApp.DoCmd.SetWarnings(false);
                    try
                    {
                        try { exclusiveApp.DoCmd.DeleteObject(objectType, objectName); } catch { }
                        exclusiveApp.LoadFromText(objectType, objectName, tmp);
                    }
                    finally
                    {
                        try { exclusiveApp.DoCmd.SetWarnings(true); } catch { }
                    }
                }
                finally
                {
                    try { exclusiveApp.CloseCurrentDatabase(); } catch { }
                    try { exclusiveApp.Quit(2); } catch { }
                    try { Marshal.FinalReleaseComObject(exclusiveApp); } catch { }
                    for (int i = 0; i < 100 && File.Exists(lockFile); i++)
                        Thread.Sleep(100);
                }

                Connect(dbPath);
            }
            finally
            {
                if (File.Exists(tmp)) File.Delete(tmp);
            }
        }

        private void DisconnectInternal()
        {
            if (_app == null) return;
            object appRef = _app;
            var dbPath = _databasePath;
            _app = null;
            try { ((dynamic)appRef).CloseCurrentDatabase(); } catch { }
            try { ((dynamic)appRef).Quit(2); } catch { }
            try { Marshal.FinalReleaseComObject(appRef); } catch { }
            try
            {
                if (_accessProcess != null)
                {
                    if (!_accessProcess.WaitForExit(10_000))
                        _accessProcess.Kill();
                    _accessProcess.WaitForExit(3_000);
                }
            }
            catch { }
            _accessProcess = null;
            // Delete the stale lock file Access sometimes leaves behind after exit.
            // Retry up to 5 s to allow the OS to fully release the file handle.
            try
            {
                if (!string.IsNullOrEmpty(dbPath))
                {
                    var lockFile = Path.ChangeExtension(dbPath, ".laccdb");
                    Console.Error.WriteLine($"[disconnect] Checking for lock file: {lockFile}");
                    for (int i = 0; i < 50 && File.Exists(lockFile); i++)
                    {
                        try
                        {
                            File.Delete(lockFile);
                            Console.Error.WriteLine($"[disconnect] Lock file deleted on attempt {i + 1}.");
                            break;
                        }
                        catch (Exception ex)
                        {
                            Console.Error.WriteLine($"[disconnect] Delete attempt {i + 1} failed ({ex.GetType().Name}): {ex.Message}");
                            Thread.Sleep(100);
                        }
                    }
                    if (File.Exists(lockFile))
                        Console.Error.WriteLine($"[disconnect] Lock file still present after retries: {lockFile}");
                }
            }
            catch (Exception ex) { Console.Error.WriteLine($"[disconnect] Lock file cleanup error: {ex.Message}"); }
        }

        private dynamic EnsureApp()
        {
            if (_app == null)
                throw new InvalidOperationException("Not connected. Call Connect() first.");

            if (!IsDatabaseSessionOpen(_app))
            {
                if (string.IsNullOrWhiteSpace(_databasePath) || !File.Exists(_databasePath))
                    throw new InvalidOperationException("Access connection is stale and no valid database path is available to reconnect.");

                Connect(_databasePath);
            }

            return _app;
        }

        private static bool IsDatabaseSessionOpen(dynamic app)
        {
            object db = null;
            try
            {
                db = app.CurrentDb();
                var dbName = db?.GetType().InvokeMember("Name",
                    System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance,
                    null, db, null) as string;
                if (string.IsNullOrWhiteSpace(dbName))
                    return false;

                var currentProject = app.CurrentProject;
                var projectName = SafeProp(currentProject, "Name");
                return !string.IsNullOrWhiteSpace(projectName);
            }
            catch
            {
                return false;
            }
            finally
            {
                if (db != null && Marshal.IsComObject(db))
                {
                    try { Marshal.ReleaseComObject(db); } catch { }
                }
            }
        }

        private static bool FormExists(dynamic app, string formName)
        {
            try
            {
                dynamic all = app.CurrentProject.AllForms;
                for (int i = 0; i < (int)all.Count; i++)
                {
                    dynamic f = all.Item(i);
                    if (string.Equals((string)f.Name, formName, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
            catch { }
            return false;
        }

        private static void WaitForDatabaseReady(dynamic app)
        {
            Exception lastError = null;
            var timeoutAt = DateTime.UtcNow.AddSeconds(15);

            while (DateTime.UtcNow < timeoutAt)
            {
                object db = null;
                try
                {
                    db = app.CurrentDb();
                    var dbName = db?.GetType().InvokeMember("Name",
                        System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance,
                        null, db, null) as string;
                    if (!string.IsNullOrWhiteSpace(dbName))
                    {
                        var currentProject = app.CurrentProject;
                        var projectName = SafeProp(currentProject, "Name");
                        if (!string.IsNullOrWhiteSpace(projectName))
                            return;
                    }
                }
                catch (Exception ex)
                {
                    lastError = ex;
                }
                finally
                {
                    if (db != null && Marshal.IsComObject(db))
                    {
                        try { Marshal.ReleaseComObject(db); } catch { }
                    }
                }

                Thread.Sleep(100);
            }

            throw new InvalidOperationException(
                "Access did not finish opening the database within the timeout window.",
                lastError);
        }

        private static void TryForceDisableAutomationMacros(dynamic app)
        {
            try
            {
                // 3 == msoAutomationSecurityForceDisable
                app.AutomationSecurity = 3;
            }
            catch
            {
            }
        }

        private static bool IsFormLoaded(dynamic app, string formName)
        {
            try
            {
                dynamic all = app.CurrentProject.AllForms;
                for (int i = 0; i < (int)all.Count; i++)
                {
                    dynamic f = all.Item(i);
                    if (string.Equals((string)f.Name, formName, StringComparison.OrdinalIgnoreCase))
                        return (bool)f.IsLoaded;
                }
            }
            catch { }
            return false;
        }

        private static string GetVbaComponentTypeName(int kind)
        {
            switch (kind)
            {
                case 1:
                    return "StandardModule";
                case 2:
                    return "ClassModule";
                case 3:
                    return "UserForm";
                case 100:
                    return "Document";
                default:
                    return "Type" + kind;
            }
        }

        private static string SafeProp(dynamic obj, string prop)
        {
            try { return obj.GetType().InvokeMember(prop,
                System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.Public,
                null, obj, null)?.ToString(); }
            catch { return null; }
        }

        private static int? SafePropInt(dynamic obj, string prop)
        {
            var s = SafeProp(obj, prop);
            if (s == null) return null;
            int v;
            return int.TryParse(s, out v) ? (int?)v : null;
        }

        private static bool? SafePropBool(dynamic obj, string prop)
        {
            var s = SafeProp(obj, prop);
            if (s == null) return null;
            bool bv;
            return bool.TryParse(s, out bv) ? (bool?)bv : null;
        }

        private static string SafeAppProp(dynamic app, string propName)
        {
            try
            {
                var value = app.GetType().InvokeMember(propName,
                    System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance,
                    null, app, null);
                return value == null ? null : value.ToString();
            }
            catch
            {
                return null;
            }
        }

        private static object ParseValue(string value)
        {
            if (value == null) return "";

            bool boolValue;
            if (bool.TryParse(value, out boolValue)) return boolValue;

            int intValue;
            if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out intValue)) return intValue;

            double doubleValue;
            if (double.TryParse(value, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out doubleValue)) return doubleValue;

            return value;
        }

        private static string NormalizeOptionValue(object value)
        {
            if (value == null) return "";

            bool boolValue;
            if (value is bool) return (bool)value ? "True" : "False";

            var text = value.ToString();
            if (string.Equals(text, "-1", StringComparison.OrdinalIgnoreCase)) return "True";
            if (string.Equals(text, "0", StringComparison.OrdinalIgnoreCase)) return "False";
            if (bool.TryParse(text, out boolValue)) return boolValue ? "True" : "False";

            return text;
        }

        private static string UnwrapMessage(Exception ex)
        {
            var tie = ex as System.Reflection.TargetInvocationException;
            if (tie != null && tie.InnerException != null && !string.IsNullOrWhiteSpace(tie.InnerException.Message))
                return tie.InnerException.Message;

            if (ex.InnerException != null && !string.IsNullOrWhiteSpace(ex.InnerException.Message))
                return ex.InnerException.Message;

            return ex.Message;
        }
    }

    // ── Data models ───────────────────────────────────────────────────────────

    public class AccessObjectInfo
    {
        public string Name     { get; set; } = "";
        public bool   IsLoaded { get; set; }
    }

    public class QueryInfo
    {
        public string Name { get; set; } = "";
        public string Sql  { get; set; } = "";
    }

    public class VbaProjectInfo
    {
        public string Name { get; set; } = "";
        public List<VbaComponentInfo> Components { get; set; } = new List<VbaComponentInfo>();
    }

    public class VbaComponentInfo
    {
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
    }

    public class ControlInfo
    {
        public string Name            { get; set; } = "";
        public int    ControlType     { get; set; }
        public string ControlTypeName { get; set; } = "";
        public string Caption         { get; set; }
        public string ControlSource   { get; set; }
        public int?   Left            { get; set; }
        public int?   Top             { get; set; }
        public int?   Width           { get; set; }
        public int?   Height          { get; set; }
        public bool?  Visible         { get; set; }
        public bool?  Enabled         { get; set; }
    }

    public class FieldDefinitionInfo
    {
        public string Name { get; set; }
        public int Type { get; set; }
        public string TypeName { get; set; }
        public int Size { get; set; }
        public bool Required { get; set; }
        public bool AllowZeroLength { get; set; }
    }

    public class TableDefinitionInfo
    {
        public string Name { get; set; }
        public int FieldCount { get; set; }
        public List<FieldDefinitionInfo> Fields { get; set; }
    }

    public class CompileResultInfo
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public List<string> Errors { get; set; }
    }

    public class DatabaseObjectsSummary
    {
        public int Tables { get; set; }
        public int Forms { get; set; }
        public int Reports { get; set; }
        public int Queries { get; set; }
        public int Modules { get; set; }
        public int Macros { get; set; }
        public int TotalFields { get; set; }
    }
}
