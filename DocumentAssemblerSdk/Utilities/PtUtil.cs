using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Schema;

namespace DocumentAssembler.Core
{
    public static class PtExtensions
    {
        public static XElement GetXElement(this XmlNode node)
        {
            var xDoc = new XDocument();
            using (var xmlWriter = xDoc.CreateWriter())
            {
                node.WriteTo(xmlWriter);
            }

            if (xDoc.Root == null)
            {
                throw new InvalidOperationException("XDocument does not have a root element.");
            }
            return xDoc.Root;
        }

        public static XmlNode GetXmlNode(this XElement element)
        {
            var xmlDoc = new XmlDocument();
            using (var xmlReader = element.CreateReader())
            {
                xmlDoc.Load(xmlReader);
            }

            return xmlDoc;
        }

        public static XDocument GetXDocument(this XmlDocument document)
        {
            var xDoc = new XDocument();
            using (var xmlWriter = xDoc.CreateWriter())
            {
                document.WriteTo(xmlWriter);
            }

            var decl = document.ChildNodes.OfType<XmlDeclaration>().FirstOrDefault();
            if (decl != null)
            {
                xDoc.Declaration = new XDeclaration(decl.Version, decl.Encoding, decl.Standalone);
            }

            return xDoc;
        }

        public static XmlDocument GetXmlDocument(this XDocument document)
        {
            var xmlDoc = new XmlDocument();
            using (var xmlReader = document.CreateReader())
            {
                xmlDoc.Load(xmlReader);
                if (document.Declaration != null)
                {
                    var dec = xmlDoc.CreateXmlDeclaration(document.Declaration.Version ?? "1.0",
                        document.Declaration.Encoding, document.Declaration.Standalone);
                    xmlDoc.InsertBefore(dec, xmlDoc.FirstChild);
                }
            }
            return xmlDoc;
        }

        public static string StringConcatenate(this IEnumerable<string> source)
        {
            return source.Aggregate(
                new StringBuilder(),
                (sb, s) => sb.Append(s),
                sb => sb.ToString());
        }

        public static string StringConcatenate<T>(this IEnumerable<T> source, Func<T, string> projectionFunc)
        {
            return source.Aggregate(
                new StringBuilder(),
                (sb, i) => sb.Append(projectionFunc(i)),
                sb => sb.ToString());
        }

        public static IEnumerable<IGrouping<TKey, TSource>> GroupAdjacent<TSource, TKey>(
            this IEnumerable<TSource> source,
            Func<TSource, TKey> keySelector)
        {
            TKey? last = default;
            var haveLast = false;
            var list = new List<TSource>();

            foreach (var s in source)
            {
                var k = keySelector(s);
                if (haveLast)
                {
                    if (k != null && !k.Equals(last))
                    {
                        yield return new GroupOfAdjacent<TSource, TKey>(list, last!);

                        list = new List<TSource> { s };
                        last = k;
                    }
                    else
                    {
                        list.Add(s);
                        last = k;
                    }
                }
                else
                {
                    list.Add(s);
                    last = k;
                    haveLast = true;
                }
            }
            if (haveLast && last != null)
            {
                yield return new GroupOfAdjacent<TSource, TKey>(list, last);
            }
        }

        private static void InitializeSiblingsReverseDocumentOrder(XElement element)
        {
            XElement? prev = null;
            foreach (var e in element.Elements())
            {
                e.AddAnnotation(new SiblingsReverseDocumentOrderInfo { PreviousSibling = prev });
                prev = e;
            }
        }

        public static IEnumerable<XElement> SiblingsBeforeSelfReverseDocumentOrder(
            this XElement element)
        {
            if (element.Annotation<SiblingsReverseDocumentOrderInfo>() == null && element.Parent != null)
            {
                InitializeSiblingsReverseDocumentOrder(element.Parent);
            }

            var current = element;
            while (true)
            {
                var annotation = current.Annotation<SiblingsReverseDocumentOrderInfo>();
                if (annotation == null)
                {
                    yield break;
                }

                var previousElement = annotation.PreviousSibling;
                if (previousElement == null)
                {
                    yield break;
                }

                yield return previousElement;

                current = previousElement;
            }
        }

        public static IEnumerable<XElement> DescendantsTrimmed(this XElement element,
            XName trimName)
        {
            return element.DescendantsTrimmed(e => e.Name == trimName);
        }

        public static IEnumerable<XElement> DescendantsTrimmed(this XElement element,
            Func<XElement, bool> predicate)
        {
            var iteratorStack = new Stack<IEnumerator<XElement>>();
            iteratorStack.Push(element.Elements().GetEnumerator());
            while (iteratorStack.Count > 0)
            {
                while (iteratorStack.Peek().MoveNext())
                {
                    var currentXElement = iteratorStack.Peek().Current;
                    if (predicate(currentXElement))
                    {
                        yield return currentXElement;
                        continue;
                    }
                    yield return currentXElement;
                    iteratorStack.Push(currentXElement.Elements().GetEnumerator());
                }
                iteratorStack.Pop();
            }
        }

        public static IEnumerable<TResult> Rollup<TSource, TResult>(
            this IEnumerable<TSource> source,
            TResult seed,
            Func<TSource, TResult, TResult> projection)
        {
            var nextSeed = seed;
            foreach (var src in source)
            {
                var projectedValue = projection(src, nextSeed);
                nextSeed = projectedValue;
                yield return projectedValue;
            }
        }

        public static IEnumerable<TResult> Rollup<TSource, TResult>(
            this IEnumerable<TSource> source,
            TResult seed,
            Func<TSource, TResult, int, TResult> projection)
        {
            var nextSeed = seed;
            var index = 0;
            foreach (var src in source)
            {
                var projectedValue = projection(src, nextSeed, index++);
                nextSeed = projectedValue;
                yield return projectedValue;
            }
        }

        public static bool? ToBoolean(this XAttribute a)
        {
            if (a == null)
            {
                return null;
            }

            var s = ((string)a).ToLower();
            switch (s)
            {
                case "1":
                    return true;

                case "0":
                    return false;

                case "true":
                    return true;

                case "false":
                    return false;

                case "on":
                    return true;

                case "off":
                    return false;

                default:
                    return (bool)a;
            }
        }
    }


    public class SiblingsReverseDocumentOrderInfo
    {
        public XElement? PreviousSibling { get; set; }
    }

    public class GroupOfAdjacent<TSource, TKey> : IGrouping<TKey, TSource>
    {
        public GroupOfAdjacent(List<TSource> source, TKey key)
        {
            GroupList = source;
            Key = key;
        }

        public TKey Key { get; set; }
        private List<TSource> GroupList { get; set; }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return ((IEnumerable<TSource>)this).GetEnumerator();
        }

        IEnumerator<TSource> IEnumerable<TSource>.GetEnumerator()
        {
            return ((IEnumerable<TSource>)GroupList).GetEnumerator();
        }
    }




}
