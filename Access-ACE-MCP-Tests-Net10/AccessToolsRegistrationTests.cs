using System.Reflection;
using AccessAceMcp;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ModelContextProtocol.Server;
using SysDescription = System.ComponentModel.DescriptionAttribute;

namespace AccessAceMcpTests;

// Reflection-based tests that verify the MCP tool registrations on AccessTools
// are complete and correct. No running process or Access installation needed.

[TestClass]
[TestCategory("Registration")]
public class AccessToolsRegistrationTests
{
    private static readonly MethodInfo[] ToolMethods =
        typeof(AccessTools)
            .GetMethods(BindingFlags.Instance | BindingFlags.Public | BindingFlags.DeclaredOnly)
            .Where(m => m.GetCustomAttribute<McpServerToolAttribute>() != null)
            .ToArray();

    private static readonly string[] ExpectedToolNames =
    [
        "connect_access", "disconnect_access", "is_connected",
        "get_application_info", "get_option", "set_option", "eval_expression",
        "run_procedure", "run_command",
        "invoke_application_method", "invoke_docmd_method",
        "export_object_to_text", "import_object_from_text", "delete_object",
        "get_forms", "get_reports", "get_modules", "get_macros",
        "get_queries", "get_vba_projects",
        "get_vba_code", "set_vba_code", "add_vba_procedure", "compile_vba",
        "open_form", "close_form",
        "get_form_controls", "get_control_properties", "set_control_property",
        "export_form_to_text", "import_form_from_text",
        "export_report_to_text", "import_report_from_text",
        "delete_form", "delete_report",
        "get_table_definitions", "get_table_definition",
        "get_objects_by_type", "compile_vba_with_errors", "get_database_summary",
    ];

    // ── Count & completeness ──────────────────────────────────────────────────

    [TestMethod]
    public void ToolCount_Is40()
    {
        Assert.AreEqual(40, ToolMethods.Length,
            $"Expected 40 tools. Found: {string.Join(", ", ToolMethods.Select(m => m.Name))}");
    }

    [TestMethod]
    public void AllExpectedToolNames_ArePresent()
    {
        var registeredNames = ToolMethods
            .Select(m => m.GetCustomAttribute<McpServerToolAttribute>()!.Name)
            .ToHashSet();

        foreach (var name in ExpectedToolNames)
            Assert.IsTrue(registeredNames.Contains(name), $"Tool '{name}' is missing from AccessTools");
    }

    [TestMethod]
    public void NoToolNames_AreDuplicated()
    {
        var names = ToolMethods
            .Select(m => m.GetCustomAttribute<McpServerToolAttribute>()!.Name)
            .ToList();

        var duplicates = names.GroupBy(n => n).Where(g => g.Count() > 1).Select(g => g.Key).ToList();
        Assert.AreEqual(0, duplicates.Count,
            $"Duplicate tool names: {string.Join(", ", duplicates)}");
    }

    // ── Naming conventions ────────────────────────────────────────────────────

    [TestMethod]
    public void AllToolNames_AreSnakeCase()
    {
        foreach (var m in ToolMethods)
        {
            var name = m.GetCustomAttribute<McpServerToolAttribute>()!.Name;
            Assert.IsNotNull(name, $"{m.Name} has null tool Name");
            Assert.IsFalse(string.IsNullOrWhiteSpace(name), $"{m.Name} has empty tool Name");
            Assert.IsFalse(name.Any(char.IsUpper),
                $"Tool '{name}' on {m.Name} contains uppercase — expected snake_case");
        }
    }

    // ── Description attributes ────────────────────────────────────────────────

    [TestMethod]
    public void AllToolMethods_HaveNonEmptyDescription()
    {
        foreach (var m in ToolMethods)
        {
            var desc = m.GetCustomAttribute<SysDescription>();
            Assert.IsNotNull(desc, $"Tool method {m.Name} is missing [Description]");
            Assert.IsFalse(string.IsNullOrWhiteSpace(desc.Description),
                $"Tool method {m.Name} has empty description text");
        }
    }

    [TestMethod]
    public void AllStringParameters_HaveDescriptions()
    {
        foreach (var m in ToolMethods)
        {
            foreach (var p in m.GetParameters())
            {
                if (p.HasDefaultValue) continue; // optional params may omit description
                var desc = p.GetCustomAttribute<SysDescription>();
                Assert.IsNotNull(desc,
                    $"Parameter '{p.Name}' on tool {m.Name} is missing [Description]");
            }
        }
    }

    // ── Type contracts ────────────────────────────────────────────────────────

    [TestMethod]
    public void RunCommand_CommandId_IsTypedAsInt()
    {
        var m = ToolMethods.Single(x => x.GetCustomAttribute<McpServerToolAttribute>()!.Name == "run_command");
        var p = m.GetParameters().Single(x => x.Name == "command_id");
        Assert.AreEqual(typeof(int), p.ParameterType,
            "command_id should be int so the SDK emits an integer JSON schema, not string");
    }

    [TestMethod]
    public void ExportObjectToText_ObjectType_IsTypedAsInt()
    {
        var m = ToolMethods.Single(x => x.GetCustomAttribute<McpServerToolAttribute>()!.Name == "export_object_to_text");
        var p = m.GetParameters().Single(x => x.Name == "object_type");
        Assert.AreEqual(typeof(int), p.ParameterType);
    }

    [TestMethod]
    public void ImportObjectFromText_ObjectType_IsTypedAsInt()
    {
        var m = ToolMethods.Single(x => x.GetCustomAttribute<McpServerToolAttribute>()!.Name == "import_object_from_text");
        var p = m.GetParameters().Single(x => x.Name == "object_type");
        Assert.AreEqual(typeof(int), p.ParameterType);
    }

    [TestMethod]
    public void DeleteObject_ObjectType_IsTypedAsInt()
    {
        var m = ToolMethods.Single(x => x.GetCustomAttribute<McpServerToolAttribute>()!.Name == "delete_object");
        var p = m.GetParameters().Single(x => x.Name == "object_type");
        Assert.AreEqual(typeof(int), p.ParameterType);
    }

    // ── Optional parameters ───────────────────────────────────────────────────

    [TestMethod]
    public void RunProcedure_ArgumentsJson_IsNullableWithNullDefault()
    {
        var m = ToolMethods.Single(x => x.GetCustomAttribute<McpServerToolAttribute>()!.Name == "run_procedure");
        var p = m.GetParameters().Single(x => x.Name == "arguments_json");
        Assert.IsTrue(p.HasDefaultValue, "arguments_json should be optional (default = null)");
        Assert.IsNull(p.DefaultValue, "arguments_json default should be null");
        Assert.AreEqual(typeof(string), p.ParameterType.GetGenericArguments().FirstOrDefault() ?? p.ParameterType,
            "arguments_json should be nullable string");
    }

    [TestMethod]
    public void InvokeApplicationMethod_ArgumentsJson_IsOptional()
    {
        var m = ToolMethods.Single(x => x.GetCustomAttribute<McpServerToolAttribute>()!.Name == "invoke_application_method");
        var p = m.GetParameters().Single(x => x.Name == "arguments_json");
        Assert.IsTrue(p.HasDefaultValue);
    }

    [TestMethod]
    public void InvokeDoCmdMethod_ArgumentsJson_IsOptional()
    {
        var m = ToolMethods.Single(x => x.GetCustomAttribute<McpServerToolAttribute>()!.Name == "invoke_docmd_method");
        var p = m.GetParameters().Single(x => x.Name == "arguments_json");
        Assert.IsTrue(p.HasDefaultValue);
    }

    // ── Return types ──────────────────────────────────────────────────────────

    [TestMethod]
    public void AllToolMethods_ReturnTaskOfString()
    {
        foreach (var m in ToolMethods)
        {
            Assert.AreEqual(typeof(Task<string>), m.ReturnType,
                $"Tool method {m.Name} should return Task<string>");
        }
    }

    // ── Class-level attribute ─────────────────────────────────────────────────

    [TestMethod]
    public void AccessTools_HasMcpServerToolTypeAttribute()
    {
        Assert.IsNotNull(typeof(AccessTools).GetCustomAttribute<McpServerToolTypeAttribute>(),
            "AccessTools class must be decorated with [McpServerToolType]");
    }
}
