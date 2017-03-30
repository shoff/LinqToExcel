// copyright(c) 2016 Stephen Workman (workman.stephen@gmail.com)

using System;
using LinqToExcel.Logging;

namespace LinqToExcel.Tests {
    using NLog;

    public class LogManagerFactory : ILogManagerFactory {
      public ILogger GetLogger(string name) {
         return LogManager.GetLogger(name);
      }

      public ILogger GetLogger(Type type) {
         return LogManager.GetLogger(type.Name);
      }
   }
}
