using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
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
    public void CoreAssemblyReferencesNoHostOrOfficeInteropAssembly()
    {
        var references = typeof(CommandPlan).Assembly.GetReferencedAssemblies();

        Assert.DoesNotContain(references, reference =>
            ForbiddenAssemblyPrefixes.Any(prefix =>
                reference.Name?.StartsWith(prefix, StringComparison.Ordinal) == true));
    }

    [Fact]
    public void CorePublicApiExposesNoHostOfficeOrComType()
    {
        var violations = new List<string>();
        var assembly = typeof(CommandPlan).Assembly;

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

        Assert.Empty(violations);
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
