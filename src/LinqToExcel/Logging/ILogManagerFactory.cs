// copyright(c) 2016 Stephen Workman (workman.stephen@gmail.com)

using System;

namespace LinqToExcel.Logging {
    using NLog;

    public interface ILogManagerFactory {
        ILogger GetLogger(Type name);
      ILogger GetLogger(string name);
   }
}
