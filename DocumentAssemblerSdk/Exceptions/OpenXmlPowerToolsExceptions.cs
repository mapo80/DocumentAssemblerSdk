using System;

namespace DocumentAssembler.Core
{
    /// <summary>
    /// General exception for OpenXML PowerTools operations.
    /// </summary>
    public class OpenXmlPowerToolsException : Exception
    {
        public OpenXmlPowerToolsException(string message) : base(message)
        {
        }
    }
}
