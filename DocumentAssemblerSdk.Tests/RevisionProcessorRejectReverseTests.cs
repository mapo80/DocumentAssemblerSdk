using DocumentAssembler.Core;
using System;
using System.Linq;
using System.Reflection;
using System.Xml.Linq;
using Xunit;

namespace DocumentAssembler.Tests;

public class RevisionProcessorRejectReverseTests
{
    [Fact]
    public void RejectRevisionsForPartTransform_RemovesInsertedCells()
    {
        var tableCell = new XElement(W.tc,
            new XElement(W.tcPr, new XElement(W.cellIns)),
            new XElement(W.p));
        var result = Invoke<object?>("RejectRevisionsForPartTransform", tableCell);
        Assert.Null(result);
    }

    [Fact]
    public void RejectRevisionsForPartTransform_UnwrapsParagraphPropertiesChange()
    {
        var paragraphProperty = new XElement(W.pPr,
            new XElement(W.pPrChange,
                new XElement(W.pPr,
                    new XElement(W.rPr))));

        var result = Invoke<object?>("RejectRevisionsForPartTransform", paragraphProperty);
        var element = Assert.IsType<XElement>(result);
        Assert.Equal(W.pPr, element.Name);
    }

    [Fact]
    public void RejectRevisionsForPartTransform_RemovesNumberingChange()
    {
        var numberingChange = new XElement(W.numberingChange);
        var result = Invoke<object?>("RejectRevisionsForPartTransform", numberingChange);
        Assert.Null(result);
    }

    [Fact]
    public void RejectRevisionsForPartTransform_HandlesTableGridChange()
    {
        var grid = new XElement(W.tblGrid,
            new XElement(W.tblGridChange,
                new XElement(W.tblGrid,
                    new XElement(W.gridCol, new XAttribute(W._w, 100)))));
        var result = Invoke<object?>("RejectRevisionsForPartTransform", grid);
        var element = Assert.IsType<XElement>(result);
        Assert.Equal(W.tblGrid, element.Name);
    }

    [Fact]
    public void RejectRevisionsForPartTransform_HandlesTablePropertyChanges()
    {
        var tcPr = new XElement(W.tcPr,
            new XElement(W.tcPrChange,
                new XElement(W.tcPr,
                    new XElement(W.tcBorders))));
        var result = Invoke<object?>("RejectRevisionsForPartTransform", tcPr);
        var element = Assert.IsType<XElement>(result);
        Assert.Equal(W.tcPr, element.Name);
    }

    [Fact]
    public void RejectRevisionsForStylesTransform_UnwrapsRunPropertiesChange()
    {
        var runProperties = new XElement(W.rPr,
            new XElement(W.rPrChange,
                new XElement(W.rPr,
                    new XElement(W.b))));

        var result = Invoke<object?>("RejectRevisionsForStylesTransform", runProperties);
        var element = Assert.IsType<XElement>(result);
        Assert.Equal(W.rPr, element.Name);
    }

    [Fact]
    public void ReverseRevisionsTransform_SwapsDeletedRuns()
    {
        var del = new XElement(W.del, new XElement(W.r, new XElement(W.t, "removed")));
        var paragraph = new XElement(W.p, del);
        var info = CreateReverseInfo();
        var result = Invoke<object?>("ReverseRevisionsTransform", del, info);
        var element = Assert.IsType<XElement>(result);
        Assert.Equal(W.ins, element.Name);
    }

    [Fact]
    public void ReverseRevisionsTransform_SwapsMoveRanges()
    {
        var move = new XElement(W.moveFrom, new XElement(W.r, new XElement(W.t, "text")));
        var paragraph = new XElement(W.p, move);
        var info = CreateReverseInfo();
        var result = Invoke<object?>("ReverseRevisionsTransform", move, info);
        var element = Assert.IsType<XElement>(result);
        Assert.Equal(W.moveTo, element.Name);
    }

    [Fact]
    public void ReverseRevisionsTransform_InsertedRunBecomesDeleted()
    {
        var ins = new XElement(W.ins, new XElement(W.r, new XElement(W.t, "new")));
        var paragraph = new XElement(W.p, ins);
        var info = CreateReverseInfo();
        var result = Invoke<object?>("ReverseRevisionsTransform", ins, info);
        var element = Assert.IsType<XElement>(result);
        Assert.Equal(W.del, element.Name);
    }

    [Fact]
    public void ReverseRevisionsTransform_CustomXmlMovesSwap()
    {
        var moveStart = new XElement(W.customXmlMoveFromRangeStart);
        var info = CreateReverseInfo();
        var result = Invoke<object?>("ReverseRevisionsTransform", moveStart, info);
        var element = Assert.IsType<XElement>(result);
        Assert.Equal(W.customXmlMoveToRangeStart, element.Name);
    }

    [Fact]
    public void ReverseRevisionsTransform_MoveToRangeStartBecomesMoveFrom()
    {
        var moveStart = new XElement(W.moveToRangeStart);
        var info = CreateReverseInfo();
        var result = Invoke<object?>("ReverseRevisionsTransform", moveStart, info);
        var element = Assert.IsType<XElement>(result);
        Assert.Equal(W.moveFromRangeStart, element.Name);
    }

    [Fact]
    public void ReverseRevisionsTransform_MathRunsSwap()
    {
        var deleted = new XElement(W.del, new XElement(W.t, "m"));
        var math = new XElement(M.r, deleted);
        var wrapper = new XElement(W.p, math);
        var info = CreateReverseInfo();
        var result = Invoke<object?>("ReverseRevisionsTransform", deleted, info);
        Assert.Equal(W.ins, Assert.IsType<XElement>(result).Name);
    }

    [Fact]
    public void ReverseRevisionsTransform_InsertedParagraphMarksBecomeDelete()
    {
        var ins = new XElement(W.ins);
        var rPr = new XElement(W.rPr, ins);
        var pPr = new XElement(W.pPr, rPr);
        var info = CreateReverseInfo();
        var result = Invoke<object?>("ReverseRevisionsTransform", ins, info);
        Assert.Equal(W.del, Assert.IsType<XElement>(result).Name);
    }

    [Fact]
    public void ReverseRevisionsTransform_DeletedParagraphMarksBecomeInsert()
    {
        var del = new XElement(W.del);
        var rPr = new XElement(W.rPr, del);
        var pPr = new XElement(W.pPr, rPr);
        var info = CreateReverseInfo();
        var result = Invoke<object?>("ReverseRevisionsTransform", del, info);
        Assert.Equal(W.ins, Assert.IsType<XElement>(result).Name);
    }

    [Fact]
    public void ReverseRevisionsTransform_InInsertTextBecomesDelText()
    {
        var text = new XElement(W.t, new XText("value"));
        var info = CreateReverseInfo();
        var result = Invoke<object?>("ReverseRevisionsTransform", text, info);
        Assert.Equal(W.delText, Assert.IsType<XElement>(result).Name);
    }

    [Fact]
    public void ReverseRevisionsTransform_ConvertsCustomXmlRanges()
    {
        var delRange = new XElement(W.customXmlDelRangeStart);
        var info = CreateReverseInfo();
        var result = Invoke<object?>("ReverseRevisionsTransform", delRange, info);
        var element = Assert.IsType<XElement>(result);
        Assert.Equal(W.customXmlInsRangeStart, element.Name);
    }

    [Fact]
    public void ReverseRevisionsTransform_DeletedFieldCodeBecomesInstrText()
    {
        var element = new XElement(W.delInstrText, new XText("MERGEFIELD"));
        var info = CreateReverseInfo();
        var result = Invoke<object?>("ReverseRevisionsTransform", element, info);
        var instr = Assert.IsType<XElement>(result);
        Assert.Equal(W.instrText, instr.Name);
    }

    [Fact]
    public void ReverseRevisionsTransform_DelTextBecomesText()
    {
        var delText = new XElement(W.delText, new XText("value"));
        var info = CreateReverseInfo();
        var result = Invoke<object?>("ReverseRevisionsTransform", delText, info);
        var element = Assert.IsType<XElement>(result);
        Assert.Equal(W.t, element.Name);
    }

    [Fact]
    public void ReverseRevisionsTransform_InInsertInstrTextBecomesDeleted()
    {
        var instr = new XElement(W.instrText, new XText("CODE"));
        var info = CreateReverseInfo();
        var result = Invoke<object?>("ReverseRevisionsTransform", instr, info);
        var element = Assert.IsType<XElement>(result);
        Assert.Equal(W.delInstrText, element.Name);
    }

    [Fact]
    public void ReverseRevisionsTransform_MoveToBecomesMoveFrom()
    {
        var move = new XElement(W.moveTo);
        var info = CreateReverseInfo();
        var result = Invoke<object?>("ReverseRevisionsTransform", move, info);
        var element = Assert.IsType<XElement>(result);
        Assert.Equal(W.moveFrom, element.Name);
    }

    [Fact]
    public void ReverseRevisionsTransform_CustomXmlInsRangeEndBecomesDel()
    {
        var ins = new XElement(W.customXmlInsRangeEnd);
        var info = CreateReverseInfo();
        var result = Invoke<object?>("ReverseRevisionsTransform", ins, info);
        Assert.Equal(W.customXmlDelRangeEnd, Assert.IsType<XElement>(result).Name);
    }

    [Fact]
    public void ReverseRevisionsTransform_TableRowMarkersSwap()
    {
        var delRow = new XElement(W.del);
        var trPr = new XElement(W.trPr, delRow);
        var result = Invoke<object?>("ReverseRevisionsTransform", delRow, CreateReverseInfo());
        Assert.Equal(W.ins, Assert.IsType<XElement>(result).Name);

        var insRow = new XElement(W.ins);
        var trPr2 = new XElement(W.trPr, insRow);
        var result2 = Invoke<object?>("ReverseRevisionsTransform", insRow, CreateReverseInfo());
        Assert.Equal(W.del, Assert.IsType<XElement>(result2).Name);
    }

    [Fact]
    public void ReverseRevisionsTransform_DelInstrTextToInstrTextAndBack()
    {
        var delInstr = new XElement(W.delInstrText, new XText("FIELD"));
        var info = CreateReverseInfo();
        var result = Invoke<object?>("ReverseRevisionsTransform", delInstr, info);
        Assert.Equal(W.instrText, Assert.IsType<XElement>(result).Name);
    }

    [Fact]
    public void ReverseRevisionsTransform_MoveFromRangeEndBecomesMoveTo()
    {
        var moveEnd = new XElement(W.moveFromRangeEnd);
        var result = Invoke<object?>("ReverseRevisionsTransform", moveEnd, CreateReverseInfo());
        Assert.Equal(W.moveToRangeEnd, Assert.IsType<XElement>(result).Name);
    }

    [Fact]
    public void ReverseRevisionsTransform_CustomXmlMoveToRangeEndBecomesMoveFrom()
    {
        var moveEnd = new XElement(W.customXmlMoveToRangeEnd);
        var result = Invoke<object?>("ReverseRevisionsTransform", moveEnd, CreateReverseInfo());
        Assert.Equal(W.customXmlMoveFromRangeEnd, Assert.IsType<XElement>(result).Name);
    }

    [Fact]
    public void ReverseRevisionsTransform_CustomXmlDelRangeEndBecomesIns()
    {
        var delEnd = new XElement(W.customXmlDelRangeEnd);
        var result = Invoke<object?>("ReverseRevisionsTransform", delEnd, CreateReverseInfo());
        Assert.Equal(W.customXmlInsRangeEnd, Assert.IsType<XElement>(result).Name);
    }

    [Fact]
    public void RemoveRsidTransform_StripsAttributes()
    {
        var element = new XElement(W.p,
            new XAttribute(W.rsid, "00112233"),
            new XElement(W.r,
                new XAttribute(W.rsidR, "0044")));
        var result = Invoke<object?>("RemoveRsidTransform", element);
        var processed = Assert.IsType<XElement>(result);
        Assert.DoesNotContain(processed.Attributes(), a => a.Name == W.rsid);
    }

    [Fact]
    public void RemoveRsidTransform_RemovesRsidElement()
    {
        var rsidElement = new XElement(W.rsid, "value");
        var result = Invoke<object?>("RemoveRsidTransform", rsidElement);
        Assert.Null(result);
    }

    [Fact]
    public void MergeAdjacentTablesTransform_MergesTables()
    {
        var table1 = CreateSampleTable();
        var table2 = CreateSampleTable();
        var body = new XElement(W.body, table1, table2);
        var result = Invoke<object?>("MergeAdjacentTablesTransform", body);
        var processed = Assert.IsType<XElement>(result);
        Assert.Single(processed.Elements(W.tbl));
    }

    private static XElement CreateSampleTable()
    {
        return new XElement(W.tbl,
            new XElement(W.tblPr,
                new XElement(W.tblStyle, new XAttribute(W.val, "TableStyle"))),
            new XElement(W.tblGrid,
                new XElement(W.gridCol, new XAttribute(W._w, 1200)),
                new XElement(W.gridCol, new XAttribute(W._w, 1800))),
            new XElement(W.tr,
                new XElement(W.tc,
                    new XElement(W.tcPr,
                        new XElement(W.tcW, new XAttribute(W._w, 1200))),
                    new XElement(W.p, new XElement(W.r, new XElement(W.t, "A")))),
                new XElement(W.tc,
                    new XElement(W.tcPr,
                        new XElement(W.tcW, new XAttribute(W._w, 1800))),
                    new XElement(W.p, new XElement(W.r, new XElement(W.t, "B"))))));
    }

    private static T Invoke<T>(string methodName, params object?[] parameters)
    {
        var method = typeof(RevisionProcessor).GetMethod(methodName,
            BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);
        return (T)method!.Invoke(null, parameters)!;
    }

    private static object CreateReverseInfo()
    {
        var infoType = typeof(RevisionProcessor).Assembly.GetType("DocumentAssembler.Core.ReverseRevisionsInfo", throwOnError: true)!;
        var instance = Activator.CreateInstance(infoType)!;
        var field = infoType.GetField("InInsert", BindingFlags.Public | BindingFlags.Instance);
        field?.SetValue(instance, true);
        return instance;
    }
}
