using DocumentAssembler.Core.Exceptions;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace DocumentAssembler.Core
{
    public class BlockContentInfo
    {
        public XElement? PreviousBlockContentElement { get; set; }
        public XElement? ThisBlockContentElement { get; set; }
        public XElement? NextBlockContentElement { get; set; }
    }

    internal enum TagTypeEnum
    {
        Element,
        EndElement,
        EmptyElement
    }

    internal class Tag
    {
        public XElement? Element;
        public TagTypeEnum TagType;
    }

    internal class PotentialInRangeElements
        {
            public List<XElement> PotentialStartElementTagsInRange;
            public List<XElement> PotentialEndElementTagsInRange;

            public PotentialInRangeElements()
            {
                PotentialStartElementTagsInRange = new List<XElement>();
                PotentialEndElementTagsInRange = new List<XElement>();
            }
        }

    public partial class RevisionProcessor
    {
        private static IEnumerable<Tag> DescendantAndSelfTags(XElement? element)
        {
            yield return new Tag
            {
                Element = element,
                TagType = TagTypeEnum.Element
            };
            var iteratorStack = new Stack<IEnumerator<XElement>>();
            iteratorStack.Push(element.Elements().GetEnumerator());
            while (iteratorStack.Count > 0)
            {
                if (iteratorStack.Peek().MoveNext())
                {
                    var currentXElement = iteratorStack.Peek().Current;
                    if (!currentXElement.Nodes().Any())
                    {
                        yield return new Tag()
                        {
                            Element = currentXElement,
                            TagType = TagTypeEnum.EmptyElement
                        };
                        continue;
                    }
                    yield return new Tag()
                    {
                        Element = currentXElement,
                        TagType = TagTypeEnum.Element
                    };
                    iteratorStack.Push(currentXElement.Elements().GetEnumerator());
                    continue;
                }
                iteratorStack.Pop();
                if (iteratorStack.Count > 0)
                {
                    yield return new Tag()
                    {
                        Element = iteratorStack.Peek().Current,
                        TagType = TagTypeEnum.EndElement
                    };
                }
            }
            yield return new Tag
            {
                Element = element,
                TagType = TagTypeEnum.EndElement
            };
        }

        private static IEnumerable<BlockContentInfo> IterateBlockContentElements(XElement element)
        {
            var current = element.Elements().FirstOrDefault();
            if (current == null)
            {
                yield break;
            }

            AnnotateBlockContentElements(element);
            var currentBlockContentInfo = element.Annotation<BlockContentInfo>();
            if (currentBlockContentInfo != null)
            {
                while (true)
                {
                    yield return currentBlockContentInfo;
                    if (currentBlockContentInfo.NextBlockContentElement == null)
                    {
                        yield break;
                    }

                    currentBlockContentInfo = currentBlockContentInfo.NextBlockContentElement.Annotation<BlockContentInfo>();
                }
            }
        }

    }

}
