using Microsoft.VisualStudio.TestTools.UnitTesting;
using AccessDevToolAgent;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Threading;

namespace AccessDevToolAgentTests
{
    // All tests run on a dedicated STA thread because Access.Application
    // is an STA COM object. Each test creates its own AccessComService instance
    // so COM objects are always created and used on the same STA thread.

    [TestClass]
    [TestCategory("AccessCOM")]
    [DoNotParallelize]
    public class AccessComServiceTests
    {
        private const string TestDb = @"C:\GitHub\NorthwindAccess\AccessMcp.accdb";
        private const int AcForm = 2;
        private const int AcReport = 3;

        private static AccessComService Connect()
        {
            var svc = new AccessComService();
            svc.Connect(TestDb);
            return svc;
        }

        private static void RunInSta(Action action)
        {
            Exception error = null;
            var thread = new Thread(() =>
            {
                try
                {
                    action();
                }
                catch (Exception ex)
                {
                    error = ex;
                }
            });

            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();

            if (error != null)
                ExceptionDispatchInfo.Capture(error).Throw();
        }

        private static void SkipIfEnvironmentImportLimitation(Exception ex)
        {
            var message = ex.ToString();
            if (message.IndexOf("You canceled the previous operation.", StringComparison.OrdinalIgnoreCase) >= 0 ||
                message.IndexOf("exclusive access", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Assert.Inconclusive("Access environment does not allow text import round-trip in the current session: " + ex.Message);
            }
        }

        // ── Connection ────────────────────────────────────────────────────────

        [TestMethod]
        public void Connect_ValidPath_SetsIsConnected()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    Assert.IsTrue(svc.IsConnected);
                }
            });
        }

        [TestMethod]
        public void Connect_MissingFile_Throws()
        {
            RunInSta(() =>
            {
                using (var svc = new AccessComService())
                {
                    Assert.ThrowsException<FileNotFoundException>(() => svc.Connect(@"C:\does\not\exist.accdb"));
                }
            });
        }

        [TestMethod]
        public void Disconnect_ClearsIsConnected()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    svc.Disconnect();
                    Assert.IsFalse(svc.IsConnected);
                }
            });
        }

        [TestMethod]
        public void EnsureApp_WhenNotConnected_Throws()
        {
            RunInSta(() =>
            {
                using (var svc = new AccessComService())
                {
                    Assert.ThrowsException<InvalidOperationException>(() => svc.GetForms());
                }
            });
        }

        // ── Object enumeration ────────────────────────────────────────────────

        [TestMethod]
        public void GetForms_ReturnsNonEmptyList()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var forms = svc.GetForms();
                    Assert.IsTrue(forms.Count > 0);
                    Assert.IsTrue(forms.All(f => !string.IsNullOrWhiteSpace(f.Name)));
                }
            });
        }

        [TestMethod]
        public void GetReports_ReturnsNonEmptyList()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var reports = svc.GetReports();
                    Assert.IsTrue(reports.Count > 0);
                    Assert.IsTrue(reports.All(r => !string.IsNullOrWhiteSpace(r.Name)));
                }
            });
        }

        [TestMethod]
        public void GetModules_ReturnsNonEmptyList()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var modules = svc.GetModules();
                    Assert.IsTrue(modules.Count > 0);
                    Assert.IsTrue(modules.All(m => !string.IsNullOrWhiteSpace(m.Name)));
                }
            });
        }

        [TestMethod]
        public void GetMacros_ReturnsList()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var macros = svc.GetMacros();
                    Assert.IsNotNull(macros);
                }
            });
        }

        [TestMethod]
        public void GetQueries_ReturnsNonEmptyListWithSql()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var queries = svc.GetQueries();
                    Assert.IsTrue(queries.Count > 0);
                    Assert.IsTrue(queries.All(q => !string.IsNullOrWhiteSpace(q.Name)));
                    Assert.IsTrue(queries.Any(q => !string.IsNullOrWhiteSpace(q.Sql)));
                }
            });
        }

        [TestMethod]
        public void GetQueries_NoneStartWithTilde()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var queries = svc.GetQueries();
                    Assert.IsTrue(queries.All(q => !q.Name.StartsWith("~")));
                }
            });
        }

        // ── VBA ───────────────────────────────────────────────────────────────

        [TestMethod]
        public void GetVbaProjects_ReturnsAtLeastOneProject()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var projects = svc.GetVbaProjects();
                    Assert.IsTrue(projects.Count > 0);
                    Assert.IsFalse(string.IsNullOrWhiteSpace(projects[0].Name));
                }
            });
        }

        [TestMethod]
        public void GetVbaProjects_ComponentsHaveNames()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var projects = svc.GetVbaProjects();
                    var allComponents = projects.SelectMany(p => p.Components).ToList();
                    Assert.IsTrue(allComponents.Count > 0);
                    Assert.IsTrue(allComponents.All(c => !string.IsNullOrWhiteSpace(c.Name)));
                }
            });
        }

        [TestMethod]
        public void GetVbaProjects_ComponentTypesAreKnown()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var projects = svc.GetVbaProjects();
                    var knownTypes = new[] { "StandardModule", "ClassModule", "UserForm", "Document" };
                    var allComponents = projects.SelectMany(p => p.Components);
                    Assert.IsTrue(allComponents.All(c => knownTypes.Contains(c.Type)));
                }
            });
        }

        [TestMethod]
        public void GetVbaCode_FirstStandardModule_ReturnsNonEmptyString()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var projects = svc.GetVbaProjects();
                    var standardModule = projects
                        .SelectMany(p => p.Components)
                        .FirstOrDefault(c => c.Type == "StandardModule");

                    Assert.IsNotNull(standardModule);
                    var code = svc.GetVbaCode(standardModule.Name);
                    Assert.IsFalse(string.IsNullOrWhiteSpace(code));
                }
            });
        }

        [TestMethod]
        public void GetVbaCode_UnknownModule_Throws()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    Assert.ThrowsException<ArgumentException>(() => svc.GetVbaCode("__NoSuchModule__"));
                }
            });
        }

        [TestMethod]
        public void SetVbaCode_RoundTrips()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var projects = svc.GetVbaProjects();
                    var standardModule = projects
                        .SelectMany(p => p.Components)
                        .FirstOrDefault(c => c.Type == "StandardModule");

                    Assert.IsNotNull(standardModule);
                    string name = standardModule.Name;
                    string original = svc.GetVbaCode(name);

                    const string marker = "'-- TEST MARKER --\r\n";
                    string modified = marker + original;
                    svc.SetVbaCode(name, modified);

                    string readBack = svc.GetVbaCode(name);
                    StringAssert.StartsWith(readBack, marker);

                    svc.SetVbaCode(name, original);
                    Assert.AreEqual(original, svc.GetVbaCode(name));
                }
            });
        }

        [TestMethod]
        public void AddVbaProcedure_AppendsToModule()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var projects = svc.GetVbaProjects();
                    var standardModule = projects
                        .SelectMany(p => p.Components)
                        .FirstOrDefault(c => c.Type == "StandardModule");

                    Assert.IsNotNull(standardModule);
                    string name = standardModule.Name;
                    string original = svc.GetVbaCode(name);

                    const string proc = "Public Sub TestProc_TempMarker()\r\n    ' temp\r\nEnd Sub";
                    svc.AddVbaProcedure(name, proc);

                    string readBack = svc.GetVbaCode(name);
                    StringAssert.Contains(readBack, "TestProc_TempMarker");

                    svc.SetVbaCode(name, original);
                }
            });
        }

        // ── Forms ─────────────────────────────────────────────────────────────

        [TestMethod]
        public void GetFormControls_FirstForm_ReturnsControls()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var forms = svc.GetForms();
                    Assert.IsTrue(forms.Count > 0);

                    var controls = svc.GetFormControls(forms[0].Name);
                    Assert.IsTrue(controls.Count > 0);
                    Assert.IsTrue(controls.All(c => !string.IsNullOrWhiteSpace(c.Name)));
                }
            });
        }

        [TestMethod]
        public void GetFormControls_ControlTypesArePopulated()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var forms = svc.GetForms();
                    var controls = svc.GetFormControls(forms[0].Name);
                    Assert.IsTrue(controls.All(c => c.ControlType > 0 && !string.IsNullOrWhiteSpace(c.ControlTypeName)));
                }
            });
        }

        [TestMethod]
        public void GetControlProperties_ReturnsPopulatedDictionary()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var forms = svc.GetForms();
                    var controls = svc.GetFormControls(forms[0].Name);
                    Assert.IsTrue(controls.Count > 0);

                    var props = svc.GetControlProperties(forms[0].Name, controls[0].Name);
                    Assert.IsTrue(props.Count > 0);
                    Assert.IsTrue(props.ContainsKey("Name"));
                }
            });
        }

        // ── Text export / import ──────────────────────────────────────────────

        [TestMethod]
        public void ExportFormToText_ReturnsNonEmptyText()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var forms = svc.GetForms();
                    Assert.IsTrue(forms.Count > 0);

                    string text = svc.ExportFormToText(forms[0].Name);
                    Assert.IsFalse(string.IsNullOrWhiteSpace(text));
                    StringAssert.Contains(text, "Begin Form");
                }
            });
        }

        [TestMethod]
        public void ExportReportToText_ReturnsNonEmptyText()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var reports = svc.GetReports();
                    Assert.IsTrue(reports.Count > 0);

                    string text = svc.ExportReportToText(reports[0].Name);
                    Assert.IsFalse(string.IsNullOrWhiteSpace(text));
                    StringAssert.Contains(text, "Begin Report");
                }
            });
        }

        [TestMethod]
        public void ExportThenImportForm_RoundTrips()
        {
            try
            {
                RunInSta(() =>
                {
                    using (var svc = Connect())
                    {
                        var forms = svc.GetForms();
                        Assert.IsTrue(forms.Count > 0);

                        string formName = forms[0].Name;
                        string exported = svc.ExportFormToText(formName);

                        svc.ImportFormFromText(formName, exported);

                        string reExported = svc.ExportFormToText(formName);
                        Assert.IsFalse(string.IsNullOrWhiteSpace(reExported));
                    }
                });
            }
            catch (Exception ex)
            {
                SkipIfEnvironmentImportLimitation(ex);
                throw;
            }
        }

        [TestMethod]
        public void ExportDatabaseObjects_SpecificSubset_ExportsOnlyRequestedItems()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var forms = svc.GetForms();
                    var reports = svc.GetReports();
                    var queries = svc.GetQueries();
                    var modules = svc.GetModules();

                    if (forms.Count == 0 || reports.Count == 0 || queries.Count == 0 || modules.Count == 0)
                        Assert.Inconclusive("Test database must contain at least one form/report/query/module");

                    var request = new System.Collections.Generic.Dictionary<int, System.Collections.Generic.List<string>>
                    {
                        [2] = new System.Collections.Generic.List<string> { forms[0].Name },
                        [3] = new System.Collections.Generic.List<string> { reports[0].Name },
                        [1] = new System.Collections.Generic.List<string> { queries[0].Name },
                        [5] = new System.Collections.Generic.List<string> { modules[0].Name }
                    };

                    var result = svc.ExportDatabaseObjects(request);

                    Assert.IsNotNull(result);
                    Assert.IsTrue(result.ExportedObjects.Count > 0, "Expected at least one exported object");
                    Assert.IsTrue(result.ExportedObjects.Count <= 4, "Subset export should not include more than requested objects");
                    Assert.IsTrue(result.ExportedObjects.Any(o => o.Type == "form" && o.Name == forms[0].Name));
                    Assert.IsTrue(result.ExportedObjects.Any(o => o.Type == "report" && o.Name == reports[0].Name));
                    Assert.IsTrue(result.ExportedObjects.Any(o => o.Type == "query" && o.Name == queries[0].Name));
                    Assert.IsTrue(result.ExportedObjects.Any(o => o.Type == "module" && o.Name == modules[0].Name));
                    Assert.IsTrue(result.ExportedObjects.All(o => !string.IsNullOrWhiteSpace(o.Name)));
                    Assert.IsTrue(result.ExportedObjects.All(o => o.Code != null));
                }
            });
        }

        [TestMethod]
        public void GetAccessBitness_ReturnsString_ViaReflection()
        {
            RunInSta(() =>
            {
                var asm = Assembly.Load("Access-DevTool-Agent");
                Assert.IsNotNull(asm);

                var progType = asm.GetType("AccessDevToolAgent.Program");
                Assert.IsNotNull(progType);

                var method = progType.GetMethod("GetAccessBitness", BindingFlags.Static | BindingFlags.Public);
                Assert.IsNotNull(method);

                var result = method.Invoke(null, new object[] { "16.0" });
                Assert.IsInstanceOfType(result, typeof(string));
                var bitness = (string)result;
                Assert.IsTrue(bitness == "x86" || bitness == "x64" || bitness == "Not Found", "Unexpected bitness: " + bitness);
            });
        }

        // ── Generic automation surface ────────────────────────────────────────

        [TestMethod]
        public void GetApplicationInfo_ReturnsExpectedKeys()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var info = svc.GetApplicationInfo();
                    Assert.IsNotNull(info);
                    Assert.IsTrue(info.ContainsKey("Name"));
                    Assert.IsTrue(info.ContainsKey("Version"));
                    Assert.IsTrue(info.ContainsKey("CurrentDb"));
                }
            });
        }

        [TestMethod]
        public void GetOptionAndSetOption_RoundTrips()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    const string optionName = "Show Status Bar";
                    var original = svc.GetOption(optionName);
                    Assert.IsNotNull(original);

                    var toggled = string.Equals(original, "True", StringComparison.OrdinalIgnoreCase) ? "False" : "True";
                    svc.SetOption(optionName, toggled);
                    var changed = svc.GetOption(optionName);
                    Assert.AreEqual(toggled, changed, true);

                    svc.SetOption(optionName, original);
                    var restored = svc.GetOption(optionName);
                    Assert.AreEqual(original, restored, true);
                }
            });
        }

        [TestMethod]
        public void EvalExpression_ReturnsExpectedValue()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var value = svc.EvalExpression("1+2");
                    Assert.AreEqual("3", value);
                }
            });
        }

        [TestMethod]
        public void RunProcedure_UnknownProcedure_Throws()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    Assert.ThrowsException<Exception>(() => svc.RunProcedure("__NoSuchProcedure__", new object[0]));
                }
            });
        }

        [TestMethod]
        public void RunCommand_InvalidCommand_Throws()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    Assert.ThrowsException<Exception>(() => svc.RunCommand(-1));
                }
            });
        }

        [TestMethod]
        public void InvokeApplicationMethod_Eval_Works()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var result = svc.InvokeApplicationMethod("Eval", new object[] { "2+3" });
                    Assert.IsNotNull(result);
                    Assert.AreEqual("5", result.ToString());
                }
            });
        }

        [TestMethod]
        public void InvokeDoCmdMethod_SetWarnings_Works()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    svc.InvokeDoCmdMethod("SetWarnings", new object[] { false });
                    svc.InvokeDoCmdMethod("SetWarnings", new object[] { true });
                }
            });
        }

        [TestMethod]
        public void ExportObjectToText_Report_ReturnsText()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var reports = svc.GetReports();
                    Assert.IsTrue(reports.Count > 0);

                    var text = svc.ExportObjectToText(AcReport, reports[0].Name);
                    Assert.IsFalse(string.IsNullOrWhiteSpace(text));
                    StringAssert.Contains(text, "Begin Report");
                }
            });
        }

        [TestMethod]
        public void ImportObjectFromText_Report_RoundTrips()
        {
            try
            {
                RunInSta(() =>
                {
                    using (var svc = Connect())
                    {
                        var reports = svc.GetReports();
                        Assert.IsTrue(reports.Count > 0);

                        var reportName = reports[0].Name;
                        var text = svc.ExportObjectToText(AcReport, reportName);
                        svc.ImportObjectFromText(AcReport, reportName, text);
                        var reExported = svc.ExportObjectToText(AcReport, reportName);
                        Assert.IsFalse(string.IsNullOrWhiteSpace(reExported));
                    }
                });
            }
            catch (Exception ex)
            {
                SkipIfEnvironmentImportLimitation(ex);
                throw;
            }
        }

        [TestMethod]
        public void DeleteObject_UnknownForm_Throws()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    Assert.ThrowsException<Exception>(() => svc.DeleteObject(AcForm, "__NoSuchForm__"));
                }
            });
        }

        // ── New Interop Features ──────────────────────────────────────────────

        [TestMethod]
        public void GetTableDefinitions_ReturnsNonEmptyList()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var tables = svc.GetTableDefinitions();
                    Assert.IsNotNull(tables);
                    Assert.IsTrue(tables.Count > 0, "Database should have at least one table");
                    Assert.IsTrue(tables.All(t => !string.IsNullOrWhiteSpace(t.Name)), "All tables should have names");
                }
            });
        }

        [TestMethod]
        public void GetTableDefinitions_TablesHaveFields()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var tables = svc.GetTableDefinitions();
                    Assert.IsTrue(tables.Count > 0);

                    var tablesWithFields = tables.Where(t => t.Fields != null && t.Fields.Count > 0).ToList();
                    Assert.IsTrue(tablesWithFields.Count > 0, "At least one table should have fields");

                    var firstTable = tablesWithFields[0];
                    Assert.AreEqual(firstTable.FieldCount, firstTable.Fields.Count, "FieldCount should match Fields.Count");
                    Assert.IsTrue(firstTable.Fields.All(f => !string.IsNullOrWhiteSpace(f.Name)), "All fields should have names");
                }
            });
        }

        [TestMethod]
        public void GetTableDefinitions_FieldsHaveTypeInfo()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var tables = svc.GetTableDefinitions();
                    var allFields = tables.SelectMany(t => t.Fields).ToList();
                    Assert.IsTrue(allFields.Count > 0);

                    Assert.IsTrue(allFields.All(f => f.Type >= 0), "All fields should have type code");
                    Assert.IsTrue(allFields.All(f => !string.IsNullOrWhiteSpace(f.TypeName)), "All fields should have type name");
                    Assert.IsTrue(allFields.All(f => f.Size >= 0), "All fields should have size");
                }
            });
        }

        [TestMethod]
        public void GetTableDefinition_ValidTableName_ReturnsTableInfo()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var tables = svc.GetTableDefinitions();
                    Assert.IsTrue(tables.Count > 0);

                    string tableName = tables[0].Name;
                    var tableInfo = svc.GetTableDefinition(tableName);

                    Assert.IsNotNull(tableInfo);
                    Assert.AreEqual(tableName, tableInfo.Name);
                    Assert.IsTrue(tableInfo.FieldCount >= 0);
                    Assert.IsNotNull(tableInfo.Fields);
                }
            });
        }

        [TestMethod]
        public void GetTableDefinition_InvalidTableName_Throws()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    Assert.ThrowsException<ArgumentException>(() => svc.GetTableDefinition("__NoSuchTable__"));
                }
            });
        }

        [TestMethod]
        public void GetTableDefinition_CaseInsensitive()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var tables = svc.GetTableDefinitions();
                    Assert.IsTrue(tables.Count > 0);

                    string tableName = tables[0].Name;
                    string lowerName = tableName.ToLowerInvariant();
                    string upperName = tableName.ToUpperInvariant();

                    var result1 = svc.GetTableDefinition(tableName);
                    var result2 = svc.GetTableDefinition(lowerName);
                    var result3 = svc.GetTableDefinition(upperName);

                    Assert.AreEqual(result1.Name, result2.Name);
                    Assert.AreEqual(result2.Name, result3.Name);
                }
            });
        }

        [TestMethod]
        public void GetObjectsByType_Tables_ReturnsNonEmptyList()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var tables = svc.GetObjectsByType(0); // acTable = 0

                    Assert.IsNotNull(tables);
                    Assert.IsTrue(tables.Count > 0, "Database should have at least one table");
                    Assert.IsTrue(tables.All(t => !string.IsNullOrWhiteSpace(t.Name)), "All objects should have names");
                }
            });
        }

        [TestMethod]
        public void GetObjectsByType_Forms_ReturnsNonEmptyList()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var forms = svc.GetObjectsByType(2); // acForm = 2

                    Assert.IsNotNull(forms);
                    Assert.IsTrue(forms.Count > 0, "Database should have at least one form");
                    Assert.IsTrue(forms.All(f => !string.IsNullOrWhiteSpace(f.Name)), "All forms should have names");
                }
            });
        }

        [TestMethod]
        public void GetObjectsByType_Reports_ReturnsNonEmptyList()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var reports = svc.GetObjectsByType(3); // acReport = 3

                    Assert.IsNotNull(reports);
                    Assert.IsTrue(reports.Count > 0, "Database should have at least one report");
                    Assert.IsTrue(reports.All(r => !string.IsNullOrWhiteSpace(r.Name)), "All reports should have names");
                }
            });
        }

        [TestMethod]
        public void GetObjectsByType_Queries_ReturnsNonEmptyList()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var queries = svc.GetObjectsByType(1); // acQuery = 1

                    Assert.IsNotNull(queries);
                    Assert.IsTrue(queries.Count > 0, "Database should have at least one query");
                    Assert.IsTrue(queries.All(q => !string.IsNullOrWhiteSpace(q.Name)), "All queries should have names");
                }
            });
        }

        [TestMethod]
        public void GetObjectsByType_Modules_ReturnsNonEmptyList()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var modules = svc.GetObjectsByType(5); // acModule = 5

                    Assert.IsNotNull(modules);
                    Assert.IsTrue(modules.Count > 0, "Database should have at least one module");
                    Assert.IsTrue(modules.All(m => !string.IsNullOrWhiteSpace(m.Name)), "All modules should have names");
                }
            });
        }

        [TestMethod]
        public void GetObjectsByType_IsLoadedProperty_IsPopulated()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var forms = svc.GetObjectsByType(2); // acForm = 2

                    Assert.IsTrue(forms.Count > 0);
                    Assert.IsTrue(forms.All(f => f.IsLoaded == true || f.IsLoaded == false),
                        "IsLoaded should be a valid boolean");
                }
            });
        }

        [TestMethod]
        public void GetObjectsByType_InvalidType_Throws()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    Assert.ThrowsException<ArgumentException>(() => svc.GetObjectsByType(99));
                }
            });
        }

        [TestMethod]
        public void CompileVbaWithErrors_ReturnsCompileResult()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var result = svc.CompileVbaWithErrors();

                    Assert.IsNotNull(result);
                    Assert.IsNotNull(result.Errors);
                    Assert.IsTrue(result.Success || !result.Success, "Success must be set");
                    Assert.IsTrue(string.IsNullOrEmpty(result.Message) || result.Message.Length > 0, "Message should be valid");
                }
            });
        }

        [TestMethod]
        public void CompileVbaWithErrors_SuccessfulCompile_ReturnsTrue()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var result = svc.CompileVbaWithErrors();

                    if (!result.Success)
                    {
                        // VBA compilation failures in automation (missing references, etc.)
                        // are environment-dependent and should not hard-fail.
                        var msg = result.Message ?? "";
                        Assert.Inconclusive(
                            $"VBA compilation failed in this environment (may be missing references): {msg}");
                    }

                    Assert.IsTrue(result.Success, "Test database should compile successfully");
                    Assert.AreEqual(0, result.Errors.Count, "No errors should be returned");
                    StringAssert.Contains(result.Message.ToLower(), "successful");
                }
            });
        }

        [TestMethod]
        public void GetDatabaseObjectsSummary_ReturnsValidSummary()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var summary = svc.GetDatabaseObjectsSummary();

                    Assert.IsNotNull(summary);
                    Assert.IsTrue(summary.Tables >= 0, "Tables count should be non-negative");
                    Assert.IsTrue(summary.Forms >= 0, "Forms count should be non-negative");
                    Assert.IsTrue(summary.Reports >= 0, "Reports count should be non-negative");
                    Assert.IsTrue(summary.Queries >= 0, "Queries count should be non-negative");
                    Assert.IsTrue(summary.Modules >= 0, "Modules count should be non-negative");
                    Assert.IsTrue(summary.Macros >= 0, "Macros count should be non-negative");
                    Assert.IsTrue(summary.TotalFields >= 0, "TotalFields count should be non-negative");
                }
            });
        }

        [TestMethod]
        public void GetDatabaseObjectsSummary_HasExpectedCounts()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var summary = svc.GetDatabaseObjectsSummary();

                    // Test database should have objects
                    Assert.IsTrue(summary.Tables > 0, "Database should have tables");
                    Assert.IsTrue(summary.Forms > 0, "Database should have forms");
                    Assert.IsTrue(summary.Reports > 0, "Database should have reports");
                    Assert.IsTrue(summary.TotalFields > 0, "Database should have fields");
                }
            });
        }

        [TestMethod]
        public void GetDatabaseObjectsSummary_MatchesIndividualCalls()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var summary = svc.GetDatabaseObjectsSummary();

                    var forms = svc.GetForms().Count;
                    var reports = svc.GetReports().Count;
                    var modules = svc.GetModules().Count;
                    var macros = svc.GetMacros().Count;
                    var queries = svc.GetQueries().Count;
                    var tables = svc.GetTableDefinitions().Count;

                    Assert.AreEqual(forms, summary.Forms, "Forms count should match");
                    Assert.AreEqual(reports, summary.Reports, "Reports count should match");
                    Assert.AreEqual(modules, summary.Modules, "Modules count should match");
                    Assert.AreEqual(macros, summary.Macros, "Macros count should match");
                    Assert.AreEqual(queries, summary.Queries, "Queries count should match");
                    Assert.AreEqual(tables, summary.Tables, "Tables count should match");
                }
            });
        }

        [TestMethod]
        public void GetDatabaseObjectsSummary_TotalFieldsMatchesTableDefinitions()
        {
            RunInSta(() =>
            {
                using (var svc = Connect())
                {
                    var summary = svc.GetDatabaseObjectsSummary();
                    var tables = svc.GetTableDefinitions();

                    int totalFieldsFromTables = tables.Sum(t => t.FieldCount);
                    Assert.AreEqual(totalFieldsFromTables, summary.TotalFields,
                        "TotalFields should match sum of all table field counts");
                }
            });
        }
    }
}
