using System;

namespace RolePrivilegeAnalyzer.Models
{
    /// <summary>
    /// Represents a single privilege with its access level and source role.
    /// </summary>
    public class PrivilegeModel
    {
        public Guid PrivilegeId { get; set; }
        public string Name { get; set; } = string.Empty;
        public string AccessLevel { get; set; } = "None";
        public string SourceRoleName { get; set; } = string.Empty;
        public Guid SourceRoleId { get; set; }

        /// <summary>
        /// Numeric depth for comparison (higher = more access).
        /// </summary>
        public int AccessDepth
        {
            get
            {
                switch (AccessLevel)
                {
                    case "Global": return 4;
                    case "Deep": return 3;
                    case "Local": return 2;
                    case "Basic": return 1;
                    default: return 0;
                }
            }
        }

        /// <summary>
        /// Maps the Dataverse privilege depth mask integer to a human-readable access level.
        /// </summary>
        public static string MapDepthMask(int depthMask)
        {
            // PrivilegeDepth: 0=None, 1=Basic, 2=Local, 4=Deep, 8=Global
            switch (depthMask)
            {
                case 1: return "Basic";
                case 2: return "Local";
                case 4: return "Deep";
                case 8: return "Global";
                default: return "None";
            }
        }
    }
}
