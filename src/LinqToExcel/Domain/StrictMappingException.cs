namespace LinqToExcel.Domain
{
    using System;

    [Serializable]
    public sealed class StrictMappingException : ApplicationException
    {
        public StrictMappingException(string message)
            : base(message)
        {
        }

        public StrictMappingException(string formatMessage, params object[] args)
            : base(string.Format(formatMessage, args))
        {
        }
    }
}