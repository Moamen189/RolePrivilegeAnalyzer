using System;
using System.Collections.Generic;

namespace RolePrivilegeAnalyzer.Models
{
    /// <summary>
    /// Represents a user with their assigned roles and computed risk assessment.
    /// </summary>
    public class UserRoleModel
    {
        public Guid UserId { get; set; }
        public string BusinessUnit { get; set; } = string.Empty;
        public string FullName { get; set; } = string.Empty;
        public string DomainName { get; set; } = string.Empty;
        public List<string> Roles { get; set; } = new List<string>();
        public List<Guid> RoleIds { get; set; } = new List<Guid>();
        public string RiskLevel { get; set; } = "Unknown";
        public string RiskStatus { get; set; } = "Unknown";
        public string OrganizationUrl { get; set; } = string.Empty;

        /// <summary>
        /// Formatted roles string for display (numbered list).
        /// </summary>
        public string RolesDisplay =>
            Roles.Count == 0
                ? "(No Roles)"
                : Roles.Count == 1
                    ? Roles[0]
                    : string.Join(Environment.NewLine, Roles.ConvertAll((r) =>
                        $"{Roles.IndexOf(r) + 1}. {r}"));

        /// <summary>
        /// URL to open this user record in CRM.
        /// </summary>
        public string CrmUrl =>
            string.IsNullOrEmpty(OrganizationUrl)
                ? string.Empty
                : $"{OrganizationUrl.TrimEnd('/')}/main.aspx?etn=systemuser&id={UserId}&pagetype=entityrecord";
    }
}
