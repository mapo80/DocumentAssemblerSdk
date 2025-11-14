using DocumentAssembler.Core.Exceptions;
using System;
using Xunit;

namespace DocumentAssembler.Tests;

public class ExceptionTests
{
    [Fact]
    public void PowerToolsDocumentException_Constructors_Work()
    {
        var ex1 = new PowerToolsDocumentException();
        Assert.NotNull(ex1);

        var ex2 = new PowerToolsDocumentException("message");
        Assert.Equal("message", ex2.Message);

        var inner = new InvalidOperationException("inner");
        var ex3 = new PowerToolsDocumentException("outer", inner);
        Assert.Same(inner, ex3.InnerException);
    }

    [Fact]
    public void PowerToolsInvalidDataException_Constructors_Work()
    {
        var ex1 = new PowerToolsInvalidDataException();
        Assert.NotNull(ex1);

        var ex2 = new PowerToolsInvalidDataException("invalid");
        Assert.Equal("invalid", ex2.Message);

        var inner = new ArgumentException("inner");
        var ex3 = new PowerToolsInvalidDataException("invalid", inner);
        Assert.Same(inner, ex3.InnerException);
    }
}
