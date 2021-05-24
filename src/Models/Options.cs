using CommandLine;
using System;
using System.Collections.Generic;
using System.Text;

namespace b2c_ms_graph.Models
{
    public class Options
    {
        [Option('a', "application", Required = false, HelpText = "Set application Id.")]
        public string Application { get; set; }

        [Option('s', "secret", Required = false, HelpText = "Set application secret.")]
        public string Secret { get; set; }

        [Option('d', "decision", Required = false, HelpText = "Set application decision.")]
        public string Decision { get; set; }

        [Option('r', "rateLimit", Required = false, HelpText = "Set request rate (per second).")]
        public int RateLimit { get; set; }

        [Option('u', "users", Required = false, HelpText = "Amount of users to generate when using the 'Create Random' test options")]
        public int UserGeneration { get; set; }

        [Option('t', "tenantId", Required = false, HelpText = "Amount of users to generate when using the 'Create Random' test options")]
        public string TenantId { get; set; }
    }
}
