using DocumentAssembler.Core.Exceptions;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace DocumentAssembler.Core
{
    /// <summary>
    /// Represents a WordprocessingML document part with LINQ-to-XML access
    /// </summary>
    public class PtMainDocumentPart : XElement
    {
        private readonly WmlDocument ParentWmlDocument;

        public PtWordprocessingCommentsPart? WordprocessingCommentsPart
        {
            get
            {
                using var ms = new MemoryStream(ParentWmlDocument.DocumentByteArray);
                using var wDoc = WordprocessingDocument.Open(ms, false);
                if (wDoc.MainDocumentPart == null)
                {
                    return null;
                }
                var commentsPart = wDoc.MainDocumentPart.WordprocessingCommentsPart;
                if (commentsPart == null)
                {
                    return null;
                }

                var partElement = commentsPart.GetXDocument().Root;
                if (partElement == null)
                {
                    return null;
                }
                var childNodes = partElement.Nodes().ToList();
                foreach (var item in childNodes)
                {
                    item.Remove();
                }

                return new PtWordprocessingCommentsPart(ParentWmlDocument, commentsPart.Uri, partElement.Name, partElement.Attributes(), childNodes);
            }
        }

        public PtMainDocumentPart(WmlDocument wmlDocument, Uri uri, XName name, params object[] values)
            : base(name, values)
        {
            ParentWmlDocument = wmlDocument;
            Add(
                new XAttribute(PtOpenXml.Uri, uri),
                new XAttribute(XNamespace.Xmlns + "pt", PtOpenXml.pt)
            );
        }
    }

    /// <summary>
    /// Represents a WordprocessingML comments part with LINQ-to-XML access
    /// </summary>
    public class PtWordprocessingCommentsPart : XElement
    {
        public PtWordprocessingCommentsPart(WmlDocument wmlDocument, Uri uri, XName name, params object[] values)
            : base(name, values)
        {
            Add(new XAttribute(PtOpenXml.Uri, uri), new XAttribute(XNamespace.Xmlns + "pt", PtOpenXml.pt));
        }
    }

    /// <summary>
    /// Represents a WordprocessingML document (.docx)
    /// </summary>
    public partial class WmlDocument : OpenXmlPowerToolsDocument
    {
        #region Constructors

        public WmlDocument(OpenXmlPowerToolsDocument original)
            : base(original)
        {
            if (GetDocumentType() != typeof(WordprocessingDocument))
            {
                throw new PowerToolsDocumentException("Not a Wordprocessing document.");
            }
        }

        public WmlDocument(OpenXmlPowerToolsDocument original, bool convertToTransitional)
            : base(original, convertToTransitional)
        {
            if (GetDocumentType() != typeof(WordprocessingDocument))
            {
                throw new PowerToolsDocumentException("Not a Wordprocessing document.");
            }
        }

        public WmlDocument(string fileName)
            : base(fileName)
        {
            if (GetDocumentType() != typeof(WordprocessingDocument))
            {
                throw new PowerToolsDocumentException("Not a Wordprocessing document.");
            }
        }

        public WmlDocument(string fileName, bool convertToTransitional)
            : base(fileName, convertToTransitional)
        {
            if (GetDocumentType() != typeof(WordprocessingDocument))
            {
                throw new PowerToolsDocumentException("Not a Wordprocessing document.");
            }
        }

        public WmlDocument(string fileName, byte[] byteArray)
            : base(byteArray)
        {
            FileName = fileName;
            if (GetDocumentType() != typeof(WordprocessingDocument))
            {
                throw new PowerToolsDocumentException("Not a Wordprocessing document.");
            }
        }

        public WmlDocument(string fileName, byte[] byteArray, bool convertToTransitional)
            : base(byteArray, convertToTransitional)
        {
            FileName = fileName;
            if (GetDocumentType() != typeof(WordprocessingDocument))
            {
                throw new PowerToolsDocumentException("Not a Wordprocessing document.");
            }
        }

        public WmlDocument(string fileName, MemoryStream memStream)
            : base(fileName, memStream)
        {
        }

        public WmlDocument(string fileName, MemoryStream memStream, bool convertToTransitional)
            : base(fileName, memStream, convertToTransitional)
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// Gets the main document part with LINQ-to-XML access
        /// </summary>
        public PtMainDocumentPart MainDocumentPart
        {
            get
            {
                using var ms = new MemoryStream(DocumentByteArray);
                using var wDoc = WordprocessingDocument.Open(ms, false);
                if (wDoc.MainDocumentPart == null)
                {
                    throw new OpenXmlPowerToolsException("Document does not have a MainDocumentPart.");
                }
                var partElement = wDoc.MainDocumentPart.GetXDocument().Root;
                if (partElement == null)
                {
                    throw new OpenXmlPowerToolsException("MainDocumentPart does not have a root element.");
                }
                var childNodes = partElement.Nodes().ToList();
                foreach (var item in childNodes)
                {
                    item.Remove();
                }

                return new PtMainDocumentPart(this, wDoc.MainDocumentPart.Uri, partElement.Name, partElement.Attributes(), childNodes);
            }
        }

        #endregion
    }
}
