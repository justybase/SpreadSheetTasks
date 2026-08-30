using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace SpreadSheetTasks;

internal static class UpdaterXmlUtils
{
    internal const string SpreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    internal const string RelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";

    internal static XmlDocument LoadDocument(byte[] bytes)
    {
        var document = new XmlDocument
        {
            PreserveWhitespace = true,
            XmlResolver = null
        };
        using var stream = new MemoryStream(bytes, writable: false);
        using var reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = false,
            IgnoreWhitespace = false
        });
        document.Load(reader);
        return document;
    }

    internal static byte[] SaveDocument(XmlDocument document)
    {
        var settings = new XmlWriterSettings
        {
            Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
            OmitXmlDeclaration = document.FirstChild is not XmlDeclaration,
            Indent = false,
            NewLineHandling = NewLineHandling.None
        };
        using var stream = new MemoryStream();
        using (var writer = XmlWriter.Create(stream, settings))
        {
            document.Save(writer);
        }
        return stream.ToArray();
    }

    internal static IEnumerable<XmlElement> Descendants(XmlNode node, string localName)
    {
        if (node is XmlElement element && element.LocalName == localName)
            yield return element;

        foreach (XmlNode child in node.ChildNodes)
        {
            foreach (var descendant in Descendants(child, localName))
                yield return descendant;
        }
    }

    internal static XmlElement? FirstDescendant(XmlNode node, string localName)
    {
        foreach (var element in Descendants(node, localName))
            return element;
        return null;
    }

    internal static string GetAttribute(XmlElement element, string localName, string? namespaceUri = null)
    {
        if (namespaceUri is not null)
            return element.GetAttribute(localName, namespaceUri);

        if (element.HasAttribute(localName))
            return element.GetAttribute(localName);

        foreach (XmlAttribute attribute in element.Attributes)
        {
            if (attribute.LocalName == localName)
                return attribute.Value;
        }
        return string.Empty;
    }

    internal static void SetAttribute(XmlElement element, string localName, string value, string? namespaceUri = null)
    {
        if (namespaceUri is not null)
        {
            element.SetAttribute(localName, namespaceUri, value);
            return;
        }

        foreach (XmlAttribute attribute in element.Attributes)
        {
            if (attribute.LocalName == localName)
            {
                attribute.Value = value;
                return;
            }
        }
        element.SetAttribute(localName, value);
    }

    internal static Dictionary<string, string> ReadRelationships(byte[] bytes, string relationshipPartPath)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        var document = LoadDocument(bytes);
        string baseDirectory = GetDirectoryName(relationshipPartPath);
        foreach (var relationship in Descendants(document, "Relationship"))
        {
            string id = GetAttribute(relationship, "Id");
            string target = GetAttribute(relationship, "Target");
            if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(target))
                continue;
            result[id] = ResolvePartPath(baseDirectory, target);
        }
        return result;
    }

    internal static string ResolvePartPath(string baseDirectory, string target)
    {
        target = Uri.UnescapeDataString(target.Replace('\\', '/'));
        if (target.StartsWith('/'))
            target = target.TrimStart('/');
        else if (!target.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
            target = string.IsNullOrEmpty(baseDirectory) ? target : $"{baseDirectory}/{target}";

        var segments = new List<string>();
        foreach (string segment in target.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (segment == ".")
                continue;
            if (segment == "..")
            {
                if (segments.Count > 0)
                    segments.RemoveAt(segments.Count - 1);
                continue;
            }
            segments.Add(segment);
        }
        return string.Join('/', segments);
    }

    internal static string GetDirectoryName(string path)
    {
        int separator = path.LastIndexOf('/');
        return separator < 0 ? string.Empty : path[..separator];
    }

    internal static string ColumnToLetters(int columnIndex)
    {
        if (columnIndex < 0)
            throw new ArgumentOutOfRangeException(nameof(columnIndex));

        int number = columnIndex + 1;
        var builder = new StringBuilder();
        while (number > 0)
        {
            int remainder = (number - 1) % 26;
            builder.Insert(0, (char)('A' + remainder));
            number = (number - 1) / 26;
        }
        return builder.ToString();
    }

    internal static int ColumnFromLetters(string letters)
    {
        if (string.IsNullOrWhiteSpace(letters))
            throw new FormatException("Excel cell reference has no column.");

        int result = 0;
        foreach (char character in letters)
        {
            if (character is < 'A' or > 'Z')
                throw new FormatException($"Invalid Excel column '{letters}'.");
            result = checked(result * 26 + character - 'A' + 1);
        }
        return result - 1;
    }

    internal static bool TryParseCellReference(string reference, out int column, out int row)
    {
        column = -1;
        row = -1;
        if (string.IsNullOrWhiteSpace(reference))
            return false;

        int position = 0;
        while (position < reference.Length && char.IsLetter(reference[position]))
            position++;
        if (position == 0 || position == reference.Length)
            return false;

        if (!int.TryParse(reference[position..], out int oneBasedRow) || oneBasedRow <= 0)
            return false;
        try
        {
            column = ColumnFromLetters(reference[..position].ToUpperInvariant());
            row = oneBasedRow;
            return true;
        }
        catch (FormatException)
        {
            return false;
        }
    }

    internal static string EscapeXmlText(string value)
    {
        ArgumentNullException.ThrowIfNull(value);
        var builder = new StringBuilder(value.Length);
        for (int i = 0; i < value.Length; i++)
        {
            char character = value[i];
            if (char.IsHighSurrogate(character))
            {
                if (i + 1 < value.Length && char.IsLowSurrogate(value[i + 1]))
                {
                    builder.Append(character);
                    builder.Append(value[++i]);
                }
                continue;
            }
            if (char.IsLowSurrogate(character))
                continue;
            if (character < 0x20 && character is not ('\t' or '\n' or '\r'))
                continue;

            switch (character)
            {
                case '&': builder.Append("&amp;"); break;
                case '<': builder.Append("&lt;"); break;
                case '>': builder.Append("&gt;"); break;
                case '"': builder.Append("&quot;"); break;
                case '\'': builder.Append("&apos;"); break;
                default: builder.Append(character); break;
            }
        }
        return builder.ToString();
    }

    internal static bool ShouldPreserveWhitespace(string value)
    {
        return value.Length > 0 && (char.IsWhiteSpace(value[0]) || char.IsWhiteSpace(value[^1]));
    }

    internal static List<string> ParseSharedStringsXml(byte[] bytes)
    {
        var result = new List<string>();
        var document = LoadDocument(bytes);
        foreach (var sharedString in Descendants(document, "si"))
        {
            var builder = new StringBuilder();
            foreach (var text in Descendants(sharedString, "t"))
                builder.Append(text.InnerText);
            result.Add(builder.ToString());
        }
        return result;
    }

    internal static string BuildSharedStringItem(string value)
    {
        string escaped = EscapeXmlText(value);
        string preserve = ShouldPreserveWhitespace(value) ? " xml:space=\"preserve\"" : string.Empty;
        return $"<si><t{preserve}>{escaped}</t></si>";
    }

    internal static string ReplaceTagAttribute(string tag, string attributeName, string value)
    {
        string escaped = EscapeXmlText(value).Replace("&quot;", "&quot;", StringComparison.Ordinal);
        var regex = new Regex($"(?<prefix>\\b{Regex.Escape(attributeName)}\\s*=\\s*)(?<quote>[\\\"'])(?<value>.*?)(\\k<quote>)",
            RegexOptions.CultureInvariant | RegexOptions.Singleline);
        if (regex.IsMatch(tag))
            return regex.Replace(tag, match => match.Groups["prefix"].Value + "\"" + escaped + "\"", 1);

        int insertAt = tag.EndsWith("/>", StringComparison.Ordinal) ? tag.Length - 2 : tag.Length - 1;
        return tag.Insert(insertAt, $" {attributeName}=\"{escaped}\"");
    }

    internal static Match FindStartTag(string xml, string localName, int startAt = 0)
    {
        var regex = new Regex($"<(?<prefix>[A-Za-z_][A-Za-z0-9_.-]*:)?{Regex.Escape(localName)}\\b[^>]*>",
            RegexOptions.CultureInvariant | RegexOptions.Singleline);
        Match match = regex.Match(xml, startAt);
        if (!match.Success)
            throw new InvalidDataException($"XML element '{localName}' was not found.");
        return match;
    }

    internal static string PrefixFromTag(Match match)
    {
        return match.Groups["prefix"].Value;
    }

    internal static string FindClosingTag(string xml, string localName, string prefix, int startAt)
    {
        string escapedPrefix = Regex.Escape(prefix);
        var regex = new Regex($"</{escapedPrefix}{Regex.Escape(localName)}\\s*>",
            RegexOptions.CultureInvariant | RegexOptions.Singleline);
        Match match = regex.Match(xml, startAt);
        if (!match.Success)
            throw new InvalidDataException($"XML element '{localName}' is not closed.");
        return match.Value;
    }
}
