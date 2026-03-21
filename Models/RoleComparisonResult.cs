using System.Collections.Generic;

namespace RolePrivilegeAnalyzer.Models
{
    /// <summary>
    /// Result of comparing two users' roles and privileges.
    /// </summary>
    public class RoleComparisonResult
    {
        public UserRoleModel UserA { get; set; }
        public UserRoleModel UserB { get; set; }
        public List<string> CommonRoles { get; set; } = new List<string>();
        public List<string> OnlyInA { get; set; } = new List<string>();
        public List<string> OnlyInB { get; set; } = new List<string>();
        public List<PrivilegeDifference> PrivilegeDifferences { get; set; } = new List<PrivilegeDifference>();
    }

    /// <summary>
    /// A single privilege difference between two users.
    /// </summary>
    public class PrivilegeDifference
    {
        public string PrivilegeName { get; set; } = string.Empty;
        public string UserALevel { get; set; } = "None";
        public string UserBLevel { get; set; } = "None";
        public string UserASource { get; set; } = string.Empty;
        public string UserBSource { get; set; } = string.Empty;
    }

    /// <summary>
    /// Dashboard analytics summary.
    /// </summary>
    public class AnalyticsSummary
    {
        public int TotalUsers { get; set; }
        public int TotalRoles { get; set; }
        public int HighRiskUsers { get; set; }
        public int OverPrivilegedUsers { get; set; }
        public int NoRoleUsers { get; set; }
        public int NormalUsers { get; set; }
        public List<RoleAssignmentCount> MostAssignedRoles { get; set; } = new List<RoleAssignmentCount>();
    }

    public class RoleAssignmentCount
    {
        public string RoleName { get; set; }
        public int Count { get; set; }
    }
}
