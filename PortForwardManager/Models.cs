using System;
using System.Collections.Generic;

namespace PortForwardManager
{
    public enum PortProtocol
    {
        IPv4_to_IPv4,
        IPv4_to_IPv6,
        IPv6_to_IPv4,
        IPv6_to_IPv6
    }

    public class PortRule
    {
        public PortProtocol Protocol { get; set; }
        public int ListenPort { get; set; }
        public string TargetIP { get; set; }
        public int TargetPort { get; set; }
        public string Description { get; set; }
        public DateTime Created { get; set; } = DateTime.Now;
    }

    public class RuleProfile
    {
        public string Name { get; set; }
        public List<PortRule> Rules { get; set; } = new List<PortRule>();
        public bool IsActive { get; set; }
        public DateTime Created { get; set; } = DateTime.Now;
        public DateTime Modified { get; set; } = DateTime.Now;
    }

    public class AppSettings
    {
        public List<RuleProfile> Profiles { get; set; } = new List<RuleProfile>();
        public bool CreateBackups { get; set; } = true;
        public int MaxBackups { get; set; } = 10;
        public string LastActiveProfile { get; set; }
    }

    public class PortStatus
    {
        public int PortNumber { get; set; }
        public bool IsAvailable { get; set; }
        public string ServiceName { get; set; }
        public string ProcessInfo { get; set; }
    }
}