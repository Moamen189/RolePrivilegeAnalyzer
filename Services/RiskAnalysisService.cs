using System;
using System.Collections.Generic;
using System.Linq;
using RolePrivilegeAnalyzer.Models;

namespace RolePrivilegeAnalyzer.Services
{
    /// <summary>
    /// Analyzes user roles and privileges to compute risk levels and status.
    /// Configurable thresholds for enterprise use.
    /// </summary>
    public class RiskAnalysisService
    {
        // ── Configurable Thresholds ──
        public int OverPrivilegedRoleThreshold { get; set; } = 5;

        private static readonly HashSet<string> HighRiskRoles = new HashSet<string>(
            StringComparer.OrdinalIgnoreCase)
        {
            "System Administrator",
            "System Customizer",
            "Support User"
        };

        private readonly DataService _dataService;

        public RiskAnalysisService(DataService dataService)
        {
            _dataService = dataService ?? throw new ArgumentNullException(nameof(dataService));
        }

        /// <summary>
        /// Compute risk for all users in-place. Call after data is loaded.
        /// </summary>
        public void AnalyzeAll(List<UserRoleModel> users)
        {
            foreach (var user in users)
            {
                AnalyzeUser(user);
            }
        }

        /// <summary>
        /// Compute risk for a single user.
        /// </summary>
        public void AnalyzeUser(UserRoleModel user)
        {
            user.RiskLevel = ComputeRiskLevel(user);
            user.RiskStatus = ComputeRiskStatus(user);
        }

        /// <summary>
        /// Privilege Risk Level — based on the highest access depth across all roles.
        /// </summary>
        private string ComputeRiskLevel(UserRoleModel user)
        {
            if (user.Roles == null || user.Roles.Count == 0)
                return "None";

            int maxDepth = _dataService.GetUserMaxAccessDepth(user.UserId);

            switch (maxDepth)
            {
                case 4: return "🔴 Global";
                case 3: return "🟠 Deep";
                case 2: return "🟡 Local";
                case 1: return "🟢 Basic";
                default: return "⚪ None";
            }
        }

        /// <summary>
        /// Risk Status — combining role count, role names, and privilege analysis.
        /// </summary>
        private string ComputeRiskStatus(UserRoleModel user)
        {
            if (user.Roles == null || user.Roles.Count == 0)
                return "❌ No Roles";

            // Check for high-risk roles (System Administrator, etc.)
            bool hasHighRiskRole = user.Roles.Any(r => HighRiskRoles.Contains(r));
            if (hasHighRiskRole)
                return "🔴 High Risk";

            // Check for over-privileged (too many roles)
            if (user.Roles.Count > OverPrivilegedRoleThreshold)
                return "⚠️ Over Privileged";

            // Check if user has any Global-level privilege
            int maxDepth = _dataService.GetUserMaxAccessDepth(user.UserId);
            if (maxDepth >= 4)
                return "⚠️ Over Privileged";

            return "✅ Normal";
        }

        /// <summary>
        /// Generate analytics summary for dashboard.
        /// </summary>
        public AnalyticsSummary GenerateAnalytics(List<UserRoleModel> users)
        {
            var allRoles = _dataService.GetRolesCache();

            // Count role assignments across all users
            var roleAssignmentCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var user in users)
            {
                foreach (var role in user.Roles)
                {
                    if (!roleAssignmentCounts.ContainsKey(role))
                        roleAssignmentCounts[role] = 0;
                    roleAssignmentCounts[role]++;
                }
            }

            return new AnalyticsSummary
            {
                TotalUsers = users.Count,
                TotalRoles = allRoles.Count,
                HighRiskUsers = users.Count(u => u.RiskStatus.Contains("High Risk")),
                OverPrivilegedUsers = users.Count(u => u.RiskStatus.Contains("Over Privileged")),
                NoRoleUsers = users.Count(u => u.RiskStatus.Contains("No Roles")),
                NormalUsers = users.Count(u => u.RiskStatus.Contains("Normal")),
                MostAssignedRoles = roleAssignmentCounts
                    .OrderByDescending(kvp => kvp.Value)
                    .Take(15)
                    .Select(kvp => new RoleAssignmentCount { RoleName = kvp.Key, Count = kvp.Value })
                    .ToList()
            };
        }

        /// <summary>
        /// Compare two users' roles and privileges.
        /// </summary>
        public RoleComparisonResult CompareUsers(UserRoleModel userA, UserRoleModel userB)
        {
            var rolesA = new HashSet<string>(userA.Roles, StringComparer.OrdinalIgnoreCase);
            var rolesB = new HashSet<string>(userB.Roles, StringComparer.OrdinalIgnoreCase);

            var result = new RoleComparisonResult
            {
                UserA = userA,
                UserB = userB,
                CommonRoles = rolesA.Intersect(rolesB, StringComparer.OrdinalIgnoreCase).OrderBy(r => r).ToList(),
                OnlyInA = rolesA.Except(rolesB, StringComparer.OrdinalIgnoreCase).OrderBy(r => r).ToList(),
                OnlyInB = rolesB.Except(rolesA, StringComparer.OrdinalIgnoreCase).OrderBy(r => r).ToList()
            };

            // Compare privileges at the access level
            var privsA = _dataService.GetUserPrivileges(userA.UserId);
            var privsB = _dataService.GetUserPrivileges(userB.UserId);

            // Build max access level per privilege for each user
            var maxA = BuildMaxPrivilegeMap(privsA);
            var maxB = BuildMaxPrivilegeMap(privsB);

            var allPrivNames = new HashSet<string>(maxA.Keys, StringComparer.OrdinalIgnoreCase);
            allPrivNames.UnionWith(maxB.Keys);

            foreach (var privName in allPrivNames.OrderBy(n => n))
            {
                PrivilegeModel pA = null, pB = null;
                maxA.TryGetValue(privName, out pA);
                maxB.TryGetValue(privName, out pB);

                int levelA = pA?.AccessDepth ?? 0;
                int levelB = pB?.AccessDepth ?? 0;

                if (levelA != levelB)
                {
                    result.PrivilegeDifferences.Add(new PrivilegeDifference
                    {
                        PrivilegeName = privName,
                        UserALevel = pA?.AccessLevel ?? "None",
                        UserBLevel = pB?.AccessLevel ?? "None",
                        UserASource = pA?.SourceRoleName ?? "-",
                        UserBSource = pB?.SourceRoleName ?? "-"
                    });
                }
            }

            return result;
        }

        /// <summary>
        /// Builds a map of privilege name → highest-level PrivilegeModel.
        /// </summary>
        private Dictionary<string, PrivilegeModel> BuildMaxPrivilegeMap(List<PrivilegeModel> privileges)
        {
            var map = new Dictionary<string, PrivilegeModel>(StringComparer.OrdinalIgnoreCase);

            foreach (var p in privileges)
            {
                PrivilegeModel existing;
                if (!map.TryGetValue(p.Name, out existing) || p.AccessDepth > existing.AccessDepth)
                {
                    map[p.Name] = p;
                }
            }

            return map;
        }
    }
}
