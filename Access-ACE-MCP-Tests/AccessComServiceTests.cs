using Microsoft.VisualStudio.TestTools.UnitTesting;
using Access_ACE_MCP;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Threading;

namespace Access_ACE_MCP_Tests
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
        public void GetAccessBitness_ReturnsString_ViaReflection()
        {
            RunInSta(() =>
            {
                var asm = Assembly.Load("Access-ACE-Agent");
                Assert.IsNotNull(asm);

                var progType = asm.GetType("Access_ACE_MCP.Program");
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
    }
}
