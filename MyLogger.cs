using NLog.Config;
using NLog.Targets;
using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HelpMe
{
    internal class MyLogger
    {
        public static Logger Instance = GetLogger();
        public static Logger GetLogger() {
            var ftarget = new FileTarget
            {
                FileName = "${basedir}/Logger.log"
            };
            var rule = new LoggingRule("info", LogLevel.Info, ftarget);
            var config = new LoggingConfiguration();
            config.AddTarget("file", ftarget);    
            config.LoggingRules.Add(rule);
            LogManager.Configuration = config;
            return LogManager.GetLogger("info");
        }
    }
}
