using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using ExcelAccel.Application.Commands;
using ExcelAccel.Core.Commands;
using Xunit;

namespace ExcelAccel.Core.Tests;

public sealed class ArchitectureBoundaryTests
{
    private static readonly string[] ForbiddenAssemblyPrefixes =
    {
        "ExcelDna",
        "Microsoft.Office.Interop",
        "Microsoft.Vbe.Interop",
    };

    [Fact]
    public void PureAssembliesReferenceNoHostOrOfficeInteropAssembly()
    {
        foreach (var assembly in PureAssemblies())
        {
            var references = assembly.GetReferencedAssemblies();
            Assert.DoesNotContain(references, reference =>
                ForbiddenAssemblyPrefixes.Any(prefix =>
                    reference.Name?.StartsWith(prefix, StringComparison.Ordinal) == true));
        }
    }

    [Fact]
    public void CorePublicApiExposesNoHostOfficeOrComType()
    {
        var violations = new List<string>();
        foreach (var assembly in PureAssemblies())
        {
            foreach (var type in assembly.GetExportedTypes())
            {
                InspectType(type, $"type {type.FullName}", violations);

                foreach (var constructor in type.GetConstructors())
                {
                    foreach (var parameter in constructor.GetParameters())
                    {
                        InspectType(parameter.ParameterType, $"{type.FullName} constructor parameter {parameter.Name}", violations);
                    }
                }

                foreach (var method in type.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly))
                {
                    InspectType(method.ReturnType, $"{type.FullName}.{method.Name} return", violations);
                    foreach (var parameter in method.GetParameters())
                    {
                        InspectType(parameter.ParameterType, $"{type.FullName}.{method.Name} parameter {parameter.Name}", violations);
                    }
                }

                foreach (var property in type.GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly))
                {
                    InspectType(property.PropertyType, $"{type.FullName}.{property.Name} property", violations);
                }

                foreach (var field in type.GetFields(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly))
                {
                    InspectType(field.FieldType, $"{type.FullName}.{field.Name} field", violations);
                }
            }
        }

        Assert.Empty(violations);
    }

    [Fact]
    public void ProductionProjectDependencyGraphMatchesApprovedDirection()
    {
        var root = FindRepositoryRoot();
        AssertProjectReferences(root, "src/ExcelAccel.Core/ExcelAccel.Core.csproj");
        AssertProjectReferences(
            root,
            "src/ExcelAccel.Application/ExcelAccel.Application.csproj",
            "src/ExcelAccel.Core/ExcelAccel.Core.csproj");
        AssertProjectReferences(
            root,
            "src/ExcelAccel.ExcelInterop/ExcelAccel.ExcelInterop.csproj",
            "src/ExcelAccel.Application/ExcelAccel.Application.csproj",
            "src/ExcelAccel.Core/ExcelAccel.Core.csproj");
        AssertProjectReferences(
            root,
            "src/ExcelAccel.ExcelAddIn/ExcelAccel.ExcelAddIn.csproj",
            "src/ExcelAccel.Application/ExcelAccel.Application.csproj",
            "src/ExcelAccel.Core/ExcelAccel.Core.csproj",
            "src/ExcelAccel.ExcelInterop/ExcelAccel.ExcelInterop.csproj");

        Assert.False(Directory.Exists(Path.Combine(root, "src", "ExcelAccel.PresentationInterop")));
    }

    private static IEnumerable<Assembly> PureAssemblies()
    {
        yield return typeof(SelectionContext).Assembly;
        yield return typeof(CommandPlan).Assembly;
    }

    private static string FindRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "ExcelAccel.sln")))
        {
            directory = directory.Parent;
        }

        return Assert.IsType<DirectoryInfo>(directory).FullName;
    }

    private static void AssertProjectReferences(string root, string projectPath, params string[] expectedReferences)
    {
        var absoluteProjectPath = Path.Combine(root, projectPath.Replace('/', Path.DirectorySeparatorChar));
        var projectDirectory = Path.GetDirectoryName(absoluteProjectPath)!;
        var document = XDocument.Load(absoluteProjectPath);
        var actual = document
            .Descendants("ProjectReference")
            .Select(element => element.Attribute("Include")?.Value)
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => Path.GetFullPath(Path.Combine(projectDirectory, value!)))
            .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
            .ToArray();
        var expected = expectedReferences
            .Select(value => Path.GetFullPath(Path.Combine(root, value.Replace('/', Path.DirectorySeparatorChar))))
            .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        Assert.Equal(expected.Length, actual.Length);
        for (var index = 0; index < expected.Length; index++)
        {
            Assert.True(
                string.Equals(expected[index], actual[index], StringComparison.OrdinalIgnoreCase),
                $"Expected project reference '{expected[index]}' but found '{actual[index]}'.");
        }
    }

    private static void InspectType(Type type, string location, ICollection<string> violations)
    {
        if (type.IsByRef || type.IsPointer || type.IsArray)
        {
            InspectType(type.GetElementType()!, location, violations);
            return;
        }

        if (type.IsGenericType)
        {
            foreach (var argument in type.GetGenericArguments())
            {
                InspectType(argument, location, violations);
            }
        }

        var assemblyName = type.Assembly.GetName().Name ?? string.Empty;
        var namespaceName = type.Namespace ?? string.Empty;
        var forbidden =
            ForbiddenAssemblyPrefixes.Any(prefix => assemblyName.StartsWith(prefix, StringComparison.Ordinal)) ||
            namespaceName.StartsWith("ExcelDna", StringComparison.Ordinal) ||
            namespaceName.StartsWith("Microsoft.Office.Interop", StringComparison.Ordinal) ||
            type.FullName == "System.__ComObject" ||
            type.IsImport ||
            type.GetCustomAttributes(typeof(System.Runtime.InteropServices.ComImportAttribute), inherit: false).Length > 0;

        if (forbidden)
        {
            violations.Add($"{location} exposes {type.FullName ?? type.Name}");
        }
    }
}
