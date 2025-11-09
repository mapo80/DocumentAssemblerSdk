#!/usr/bin/env python3
"""
Create partial class files from RevisionProcessor.cs
"""
import re

# Read the source file
with open("../OpenXmlPowerTools/RevisionProcessor.cs", 'r', encoding='utf-8') as f:
    content = f.read()

# Replace namespace
content = content.replace("Codeuctivity.OpenXmlPowerTools", "DocumentAssembler.Core")
lines = content.split('\n')

# Header
header = """using DocumentAssembler.Core.Exceptions;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace DocumentAssembler.Core
{"""

# Find class boundaries
class_start = None
for i, line in enumerate(lines):
    if 'public class RevisionProcessor' in line:
        class_start = i
        break

# Extract lines by method name patterns
def extract_method(name_pattern, start_line=class_start):
    """Extract a complete method by searching for its signature."""
    for i in range(start_line, len(lines)):
        if re.search(name_pattern, lines[i]):
            # Found method start, now find end
            method_lines = []
            brace_count = 0
            started = False
            for j in range(i, len(lines)):
                method_lines.append(lines[j])
                if '{' in lines[j]:
                    started = True
                if started:
                    brace_count += lines[j].count('{') - lines[j].count('}')
                    if brace_count == 0:
                        return '\n'.join(method_lines) + '\n\n'
            break
    return ""

# 1. Main file - public API + constants + PT class
main_content = header + "\n    public partial class RevisionProcessor\n    {\n"

# Public RejectRevisions methods
main_content += extract_method(r'public static WmlDocument RejectRevisions\(WmlDocument')
main_content += extract_method(r'public static void RejectRevisions\(WordprocessingDocument')

# Public AcceptRevisions methods
main_content += extract_method(r'public static WmlDocument AcceptRevisions\(WmlDocument')
main_content += extract_method(r'public static void AcceptRevisions\(WordprocessingDocument')
main_content += extract_method(r'public static XElement\? AcceptRevisionsForElement')

# TrackedRevisionsElements array
for i in range(class_start, len(lines)):
    if 'public static readonly XName[] TrackedRevisionsElements' in lines[i]:
        j = i
        while j < len(lines) and '];' not in lines[j]:
            main_content += lines[j] + '\n'
            j += 1
        if j < len(lines):
            main_content += lines[j] + '\n\n'
        break

# HasTrackedRevisions methods
main_content += extract_method(r'public static bool PartHasTrackedRevisions')
main_content += extract_method(r'public static bool HasTrackedRevisions\(WmlDocument')
main_content += extract_method(r'public static bool HasTrackedRevisions\(WordprocessingDocument')

# PT class
for i in range(class_start, len(lines)):
    if 'public static class PT' in lines[i]:
        j = i
        brace_count = 0
        started = False
        while True:
            main_content += lines[j] + '\n'
            if '{' in lines[j]:
                started = True
            if started:
                brace_count += lines[j].count('{') - lines[j].count('}')
                if brace_count == 0:
                    break
            j += 1
        break

main_content += "    }\n\n"

# Add WmlDocument partial class
for i in range(class_start, len(lines)):
    if 'public partial class WmlDocument' in lines[i]:
        j = i
        brace_count = 0
        started = False
        while True:
            if j >= len(lines):
                break
            line = lines[j]
            # Stop before documentation comments
            if '// Markup that this code processes' in line or j > i + 20:
                # Find closing brace
                if started and brace_count > 0:
                    main_content += "    }\n"
                break
            main_content += line + '\n'
            if '{' in line:
                started = True
            if started:
                brace_count += line.count('{') - line.count('}')
                if brace_count == 0:
                    break
            j += 1
        break

main_content += "}\n"

with open("DocumentAssemblerSdk/Utilities/RevisionProcessor.cs", 'w', encoding='utf-8') as f:
    f.write(main_content)
print(f"Created RevisionProcessor.cs ({len(main_content.splitlines())} lines)")

# 2. Reject file
reject_content = header + "\n    public partial class RevisionProcessor\n    {\n"
reject_content += extract_method(r'private static void RejectRevisionsForPart')
reject_content += extract_method(r'private static object\? RejectRevisionsForPartTransform')
reject_content += extract_method(r'private static void RejectRevisionsForStylesDefinitionPart')
reject_content += extract_method(r'private static object RejectRevisionsForStylesTransform')
reject_content += "    }\n}\n"

with open("DocumentAssemblerSdk/Utilities/RevisionProcessor.Reject.cs", 'w', encoding='utf-8') as f:
    f.write(reject_content)
print(f"Created RevisionProcessor.Reject.cs ({len(reject_content.splitlines())} lines)")

# 3. Reverse file (with ReverseRevisionsInfo class)
reverse_content = header + "\n"

# Add ReverseRevisionsInfo class
for i in range(0, class_start):
    if 'internal class ReverseRevisionsInfo' in lines[i]:
        j = i
        while True:
            reverse_content += lines[j] + '\n'
            if lines[j].strip() == '}':
                break
            j += 1
        reverse_content += '\n'
        break

reverse_content += "    public partial class RevisionProcessor\n    {\n"
reverse_content += extract_method(r'private static void ReverseRevisions\(WordprocessingDocument')
reverse_content += extract_method(r'private static void ReverseRevisionsForPart')
reverse_content += extract_method(r'private static object\? RemoveRsidTransform')
reverse_content += extract_method(r'private static object MergeAdjacentTablesTransform')
reverse_content += extract_method(r'private static object ReverseRevisionsTransform')

# Add Order_tcPr dictionary
for i in range(class_start, len(lines)):
    if 'private static readonly Dictionary<XName, int> Order_tcPr' in lines[i]:
        j = i
        while j < len(lines) and '};' not in lines[j]:
            reverse_content += lines[j] + '\n'
            j += 1
        if j < len(lines):
            reverse_content += lines[j] + '\n\n'
        break

reverse_content += extract_method(r'private static XElement FixWidths')
reverse_content += "    }\n}\n"

with open("DocumentAssemblerSdk/Utilities/RevisionProcessor.Reverse.cs", 'w', encoding='utf-8') as f:
    f.write(reverse_content)
print(f"Created RevisionProcessor.Reverse.cs ({len(reverse_content.splitlines())} lines)")

# 4. Accept file - all accept-related methods
accept_content = header + "\n    public partial class RevisionProcessor\n    {\n"

accept_patterns = [
    r'private static void AcceptRevisionsForStylesDefinitionPart',
    r'private static object\? AcceptRevisionsForStylesTransform',
    r'public static void AcceptRevisionsForPart',
    r'private static object FixUpDeletedOrInsertedFieldCodesTransform',
    r'private static object TransformInstrTextToDelInstrText',
    r'private static object AddEmptyParagraphToAnyEmptyCells',
    r'private static object\? AcceptMoveFromMoveToTransform',
    r'private static XElement\? AcceptMoveFromRanges',
    r'private static object AcceptParagraphEndTagsInMoveFromTransform',
    r'private static object\? AcceptAllOtherRevisionsTransform',
    r'private static object CollapseParagraphTransform',
    r'private static void AnnotateBlockContentElements',
    r'private static void AnnotateRunElementsWithId',
    r'private static void AnnotateContentControlsWithRunIds',
    r'private static XElement AddBlockLevelContentControls',
    r'private static XElement\? AcceptDeletedAndMoveFromParagraphMarks',
    r'private static object AcceptDeletedAndMoveFromParagraphMarksTransform',
    r'private static bool AllParaContentIsDeleted',
    r'private static object\? CollapseTransform',
    r'private static bool\? IsRunContent',
    r'private static object\? AcceptDeletedAndMovedFromContentControlsTransform',
    r'private static XElement\? AcceptDeletedAndMovedFromContentControls',
    r'private static object\? AcceptMoveFromRangesTransform',
    r'private static object CoalesqueParagraphEndTagsInMoveFromTransform',
    r'private static object\? AcceptDeletedCellsTransform',
    r'private static object\? RemoveRowsLeftEmptyByMoveFrom',
]

for pattern in accept_patterns:
    method = extract_method(pattern)
    if method:
        accept_content += method

# Add helper enums and classes used in Accept methods
for i in range(class_start, len(lines)):
    if 'private enum MoveFromCollectionType' in lines[i]:
        j = i
        while j < len(lines) and lines[j].strip() != '};':
            accept_content += lines[j] + '\n'
            j += 1
        if j < len(lines):
            accept_content += lines[j] + '\n\n'
        break

for i in range(class_start, len(lines)):
    if 'private enum GroupingType' in lines[i]:
        j = i
        while j < len(lines) and lines[j].strip() != '};':
            accept_content += lines[j] + '\n'
            j += 1
        if j < len(lines):
            accept_content += lines[j] + '\n\n'
        break

for i in range(class_start, len(lines)):
    if 'private class GroupingInfo' in lines[i]:
        j = i
        brace_count = 0
        started = False
        while True:
            accept_content += lines[j] + '\n'
            if '{' in lines[j]:
                started = True
            if started:
                brace_count += lines[j].count('{') - lines[j].count('}')
                if brace_count == 0:
                    break
            j += 1
        accept_content += '\n'
        break

for i in range(class_start, len(lines)):
    if 'private enum DeletedCellCollectionType' in lines[i]:
        j = i
        while j < len(lines) and lines[j].strip() != '};':
            accept_content += lines[j] + '\n'
            j += 1
        if j < len(lines):
            accept_content += lines[j] + '\n\n'
        break

for i in range(class_start, len(lines)):
    if 'private static readonly XName[] BlockLevelElements' in lines[i]:
        j = i
        while j < len(lines) and '};' not in lines[j]:
            accept_content += lines[j] + '\n'
            j += 1
        if j < len(lines):
            accept_content += lines[j] + '\n\n'
        break

for i in range(class_start, len(lines)):
    if 'private static readonly Dictionary<XName, int> Order_sdt' in lines[i]:
        j = i
        while j < len(lines) and '};' not in lines[j]:
            accept_content += lines[j] + '\n'
            j += 1
        if j < len(lines):
            accept_content += lines[j] + '\n\n'
        break

accept_content += "    }\n}\n"

with open("DocumentAssemblerSdk/Utilities/RevisionProcessor.Accept.cs", 'w', encoding='utf-8') as f:
    f.write(accept_content)
print(f"Created RevisionProcessor.Accept.cs ({len(accept_content.splitlines())} lines)")

# 5. Helpers file - helper classes and extension methods
helpers_content = header + "\n"

# Add all helper classes
for i in range(class_start + 1, len(lines)):
    if 'public class BlockContentInfo' in lines[i] or 'internal class BlockContentInfo' in lines[i]:
        j = i
        brace_count = 0
        started = False
        while True:
            if j >= len(lines):
                break
            line = lines[j]
            if '// Markup that this code processes' in line:
                break
            helpers_content += line + '\n'
            if '{' in line:
                started = True
            if started:
                brace_count += line.count('{') - line.count('}')
                if brace_count == 0 and j > i:
                    helpers_content += '\n'
                    break
            j += 1
        break

# Add Tag-related classes and enums
for i in range(class_start, len(lines)):
    if 'private enum TagTypeEnum' in lines[i]:
        j = i
        while j < len(lines) and lines[j].strip() != '}':
            helpers_content += lines[j] + '\n'
            j += 1
        if j < len(lines):
            helpers_content += lines[j] + '\n\n'
        break

for i in range(class_start, len(lines)):
    if 'private class Tag' in lines[i]:
        j = i
        brace_count = 0
        started = False
        while True:
            helpers_content += lines[j] + '\n'
            if '{' in lines[j]:
                started = True
            if started:
                brace_count += lines[j].count('{') - lines[j].count('}')
                if brace_count == 0:
                    break
            j += 1
        helpers_content += '\n'
        break

for i in range(class_start, len(lines)):
    if 'private class PotentialInRangeElements' in lines[i]:
        j = i
        brace_count = 0
        started = False
        while True:
            helpers_content += lines[j] + '\n'
            if '{' in lines[j]:
                started = True
            if started:
                brace_count += lines[j].count('{') - lines[j].count('}')
                if brace_count == 0:
                    break
            j += 1
        helpers_content += '\n'
        break

# Add partial class with helper methods
helpers_content += "    public partial class RevisionProcessor\n    {\n"
helpers_content += extract_method(r'private static IEnumerable<Tag> DescendantAndSelfTags')
helpers_content += extract_method(r'private static IEnumerable<BlockContentInfo> IterateBlockContentElements')
helpers_content += "    }\n\n"

# Add extension class
for i in range(class_start, len(lines)):
    if 'internal static class RevisionAccepterExtensions' in lines[i]:
        j = i
        brace_count = 0
        started = False
        while True:
            if j >= len(lines):
                break
            line = lines[j]
            if '// Markup that this code processes' in line:
                break
            helpers_content += line + '\n'
            if '{' in line:
                started = True
            if started:
                brace_count += line.count('{') - line.count('}')
                if brace_count == 0:
                    break
            j += 1
        break

helpers_content += "}\n"

with open("DocumentAssemblerSdk/Utilities/RevisionProcessor.Helpers.cs", 'w', encoding='utf-8') as f:
    f.write(helpers_content)
print(f"Created RevisionProcessor.Helpers.cs ({len(helpers_content.splitlines())} lines)")

print("\nFile splitting complete!")
