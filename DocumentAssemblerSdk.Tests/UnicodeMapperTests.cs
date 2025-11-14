using DocumentAssembler.Core;
using System;
using System.Linq;
using System.Xml.Linq;
using Xunit;

namespace DocumentAssembler.Tests;

public class UnicodeMapperTests
{
    [Fact]
    public void RunToString_ShouldTranslateSpecialElements()
    {
        var symChar = new XElement(W.sym,
            new XAttribute(W.font, UniqueFont()),
            new XAttribute(W._char, "F321"));

        var run = new XElement(W.r,
            new XElement(W.rPr),
            new XElement(W.t, "Hello"),
            new XElement(W.br),
            new XElement(W.br, new XAttribute(W.type, "page")),
            new XElement(W.cr),
            new XElement(W.tab),
            new XElement(W.noBreakHyphen),
            new XElement(W.softHyphen),
            symChar,
            new XElement(W.fldChar, new XAttribute(W.fldCharType, "begin")),
            new XElement(W.fldChar, new XAttribute(W.fldCharType, "end")),
            new XElement(W.fldChar, new XAttribute(W.fldCharType, "separate")),
            new XElement(W.instrText, "FIELD"),
            new XElement(W.customXml, new XElement(W.t, "Ignored")));

        var text = UnicodeMapper.RunToString(run);

        Assert.Contains("Hello", text, StringComparison.Ordinal);
        Assert.Contains(UnicodeMapper.CarriageReturn, text);
        Assert.Contains(UnicodeMapper.FormFeed, text);
        Assert.Contains(UnicodeMapper.HorizontalTabulation, text);
        Assert.Contains(UnicodeMapper.NonBreakingHyphen, text);
        Assert.Contains(UnicodeMapper.SoftHyphen, text);
        Assert.Contains(UnicodeMapper.StartOfHeading, text);
        Assert.Contains("{", text, StringComparison.Ordinal);
        Assert.Contains("}", text, StringComparison.Ordinal);
        Assert.Contains("_", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SymToChar_ShiftsLowCodePointsIntoPrivateUse()
    {
        var fontName = UniqueFont();
        var mapped = UnicodeMapper.SymToChar(fontName, (int)0x210);
        Assert.Equal((char)(0xF000 + 0x210), mapped);
    }

    [Fact]
    public void SymToChar_UsesPrivateRangeWhenConflictingFonts()
    {
        var symbol = "F345";
        var fontOne = UniqueFont();
        var fontTwo = UniqueFont();

        var first = UnicodeMapper.SymToChar(fontOne, symbol);
        var second = UnicodeMapper.SymToChar(fontTwo, symbol);

        Assert.NotEqual(first, second);
        Assert.True(second >= UnicodeMapper.StartOfPrivateUseArea);
    }

    [Fact]
    public void SymToChar_StringOverloads_ValidateInput()
    {
        Assert.Throws<ArgumentException>(() => UnicodeMapper.SymToChar(string.Empty, "F000"));
        Assert.Throws<ArgumentException>(() => UnicodeMapper.SymToChar("Font", string.Empty));
        Assert.Throws<ArgumentNullException>(() => UnicodeMapper.SymToChar((XElement)null!));

        var notSym = new XElement(W.t, "text");
        Assert.Throws<ArgumentException>(() => UnicodeMapper.SymToChar(notSym));

        var missingFont = new XElement(W.sym, new XAttribute(W._char, "F010"));
        Assert.Throws<ArgumentException>(() => UnicodeMapper.SymToChar(missingFont));

        var missingChar = new XElement(W.sym, new XAttribute(W.font, "Font"));
        Assert.Throws<ArgumentException>(() => UnicodeMapper.SymToChar(missingChar));
    }

    [Fact]
    public void StringToCoalescedRunList_PreservesWhitespace()
    {
        var runProps = new XElement(W.rPr, new XElement(W.b));
        var runs = UnicodeMapper.StringToCoalescedRunList("  padded  ", runProps);
        Assert.Single(runs);

        var textElement = runs.Single().Element(W.t);
        Assert.NotNull(textElement);
        Assert.Equal("  padded  ", textElement!.Value);
        Assert.Equal("preserve", textElement.Attribute(XNamespace.Xml + "space")?.Value);
    }

    [Fact]
    public void StringToRunList_TranslatesSpecialCharacters()
    {
        var buffer = new string(new[]
        {
            UnicodeMapper.CarriageReturn,
            UnicodeMapper.FormFeed,
            UnicodeMapper.HorizontalTabulation,
            UnicodeMapper.NonBreakingHyphen,
            UnicodeMapper.SoftHyphen
        });
        var runProps = new XElement(W.rPr);

        var runs = UnicodeMapper.StringToRunList(buffer, runProps);
        Assert.Equal(5, runs.Count);
        Assert.Equal(W.br, runs[0].Element(W.br)?.Name);
        Assert.Equal("page", runs[1].Element(W.br)?.Attribute(W.type)?.Value);
        Assert.Equal(W.tab, runs[2].Element(W.tab)?.Name);
        Assert.Equal(W.noBreakHyphen, runs[3].Element(W.noBreakHyphen)?.Name);
        Assert.Equal(W.softHyphen, runs[4].Element(W.softHyphen)?.Name);
    }

    [Fact]
    public void CharToRunChild_IgnoresStartOfHeading()
    {
        var element = UnicodeMapper.CharToRunChild(UnicodeMapper.StartOfHeading);
        Assert.Null(element);
    }

    private static string UniqueFont()
    {
        return "Font-" + Guid.NewGuid().ToString("N");
    }
}
