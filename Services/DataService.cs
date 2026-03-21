using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using RolePrivilegeAnalyzer.Models;

namespace RolePrivilegeAnalyzer.Services
{
    /// <summary>
    /// Central data retrieval service. Loads all users, roles, and privileges
    /// in minimal queries and builds in-memory relationships.
    /// </summary>
    public class DataService
    {
        private readonly IOrganizationService _service;

        // ── Caches ──
        private Dictionary<Guid, UserRoleModel> _usersCache;
        private Dictionary<Guid, string> _rolesCache;                           // roleId → roleName
        private Dictionary<Guid, List<Guid>> _userRolesCache;                   // userId → list of roleIds
        private Dictionary<Guid, List<PrivilegeModel>> _rolePrivilegesCache;    // roleId → privileges
        private Dictionary<Guid, string> _privilegeNamesCache;                  // privilegeId → name

        public bool IsLoaded { get; private set; }

        public DataService(IOrganizationService service)
        {
            _service = service ?? throw new ArgumentNullException(nameof(service));
            ClearCaches();
        }

        public void ClearCaches()
        {
            _usersCache = new Dictionary<Guid, UserRoleModel>();
            _rolesCache = new Dictionary<Guid, string>();
            _userRolesCache = new Dictionary<Guid, List<Guid>>();
            _rolePrivilegesCache = new Dictionary<Guid, List<PrivilegeModel>>();
            _privilegeNamesCache = new Dictionary<Guid, string>();
            IsLoaded = false;
        }

        /// <summary>
        /// Master load — retrieves all data in bulk and builds caches.
        /// Progress callback: (step description, percentage 0-100).
        /// </summary>
        public async Task<List<UserRoleModel>> LoadAllDataAsync(
            string orgUrl,
            Action<string, int> progressCallback = null)
        {
            ClearCaches();

            // Step 1: Load all roles (usually < 500)
            progressCallback?.Invoke("Loading roles...", 5);
            await Task.Run(() => LoadAllRoles());

            // Step 2: Load all privileges and role-privilege mappings
            progressCallback?.Invoke("Loading privileges...", 15);
            await Task.Run(() => LoadAllPrivileges());

            // Step 3: Load role-privilege relationships
            progressCallback?.Invoke("Mapping role privileges...", 30);
            await Task.Run(() => LoadRolePrivilegeMappings());

            // Step 4: Load all active users with business units
            progressCallback?.Invoke("Loading users...", 50);
            await Task.Run(() => LoadAllUsers(orgUrl));

            // Step 5: Load user-role assignments
            progressCallback?.Invoke("Loading user-role assignments...", 70);
            await Task.Run(() => LoadUserRoleAssignments());

            // Step 6: Build final models
            progressCallback?.Invoke("Analyzing risks...", 85);
            var result = await Task.Run(() => BuildUserModels());

            progressCallback?.Invoke("Complete", 100);
            IsLoaded = true;

            return result;
        }

        #region Data Loading Methods

        private void LoadAllRoles()
        {
            var query = new QueryExpression("role")
            {
                ColumnSet = new ColumnSet("roleid", "name"),
                PageInfo = new PagingInfo { Count = 5000, PageNumber = 1 }
            };

            EntityCollection results;
            do
            {
                results = _service.RetrieveMultiple(query);
                foreach (var entity in results.Entities)
                {
                    var id = entity.Id;
                    var name = entity.GetAttributeValue<string>("name") ?? "(Unnamed)";
                    _rolesCache[id] = name;
                }
                query.PageInfo.PageNumber++;
                query.PageInfo.PagingCookie = results.PagingCookie;
            }
            while (results.MoreRecords);
        }

        private void LoadAllPrivileges()
        {
            var query = new QueryExpression("privilege")
            {
                ColumnSet = new ColumnSet("privilegeid", "name"),
                PageInfo = new PagingInfo { Count = 5000, PageNumber = 1 }
            };

            EntityCollection results;
            do
            {
                results = _service.RetrieveMultiple(query);
                foreach (var entity in results.Entities)
                {
                    _privilegeNamesCache[entity.Id] =
                        entity.GetAttributeValue<string>("name") ?? "(Unknown)";
                }
                query.PageInfo.PageNumber++;
                query.PageInfo.PagingCookie = results.PagingCookie;
            }
            while (results.MoreRecords);
        }

        private void LoadRolePrivilegeMappings()
        {
            var query = new QueryExpression("roleprivileges")
            {
                ColumnSet = new ColumnSet("roleid", "privilegeid", "privilegedepthmask"),
                PageInfo = new PagingInfo { Count = 5000, PageNumber = 1 }
            };

            EntityCollection results;
            do
            {
                results = _service.RetrieveMultiple(query);
                foreach (var entity in results.Entities)
                {
                    var roleId = entity.GetAttributeValue<Guid>("roleid");
                    var privId = entity.GetAttributeValue<Guid>("privilegeid");
                    var depthMask = entity.GetAttributeValue<int>("privilegedepthmask");

                    if (!_rolePrivilegesCache.ContainsKey(roleId))
                        _rolePrivilegesCache[roleId] = new List<PrivilegeModel>();

                    string privName;
                    _privilegeNamesCache.TryGetValue(privId, out privName);

                    string roleName;
                    _rolesCache.TryGetValue(roleId, out roleName);

                    _rolePrivilegesCache[roleId].Add(new PrivilegeModel
                    {
                        PrivilegeId = privId,
                        Name = privName ?? "(Unknown)",
                        AccessLevel = PrivilegeModel.MapDepthMask(depthMask),
                        SourceRoleId = roleId,
                        SourceRoleName = roleName ?? "(Unknown)"
                    });
                }
                query.PageInfo.PageNumber++;
                query.PageInfo.PagingCookie = results.PagingCookie;
            }
            while (results.MoreRecords);
        }

        private void LoadAllUsers(string orgUrl)
        {
            var query = new QueryExpression("systemuser")
            {
                ColumnSet = new ColumnSet("systemuserid", "fullname", "domainname", "businessunitid"),
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression("isdisabled", ConditionOperator.Equal, false),
                        // Exclude SYSTEM and INTEGRATION users
                        new ConditionExpression("accessmode", ConditionOperator.NotEqual, 3),
                    }
                },
                PageInfo = new PagingInfo { Count = 5000, PageNumber = 1 }
            };

            // Link to businessunit for BU name
            var buLink = query.AddLink("businessunit", "businessunitid", "businessunitid", JoinOperator.LeftOuter);
            buLink.Columns.AddColumns("name");
            buLink.EntityAlias = "bu";

            EntityCollection results;
            do
            {
                results = _service.RetrieveMultiple(query);
                foreach (var entity in results.Entities)
                {
                    var userId = entity.Id;
                    var buName = entity.GetAttributeValue<AliasedValue>("bu.name")?.Value as string ?? "(Unknown)";

                    _usersCache[userId] = new UserRoleModel
                    {
                        UserId = userId,
                        FullName = entity.GetAttributeValue<string>("fullname") ?? "(No Name)",
                        DomainName = entity.GetAttributeValue<string>("domainname") ?? string.Empty,
                        BusinessUnit = buName,
                        OrganizationUrl = orgUrl ?? string.Empty
                    };
                }
                query.PageInfo.PageNumber++;
                query.PageInfo.PagingCookie = results.PagingCookie;
            }
            while (results.MoreRecords);
        }

        private void LoadUserRoleAssignments()
        {
            var query = new QueryExpression("systemuserroles")
            {
                ColumnSet = new ColumnSet("systemuserid", "roleid"),
                PageInfo = new PagingInfo { Count = 5000, PageNumber = 1 }
            };

            EntityCollection results;
            do
            {
                results = _service.RetrieveMultiple(query);
                foreach (var entity in results.Entities)
                {
                    var userId = entity.GetAttributeValue<Guid>("systemuserid");
                    var roleId = entity.GetAttributeValue<Guid>("roleid");

                    if (!_userRolesCache.ContainsKey(userId))
                        _userRolesCache[userId] = new List<Guid>();

                    _userRolesCache[userId].Add(roleId);
                }
                query.PageInfo.PageNumber++;
                query.PageInfo.PagingCookie = results.PagingCookie;
            }
            while (results.MoreRecords);
        }

        #endregion

        #region Model Building

        private List<UserRoleModel> BuildUserModels()
        {
            var users = new List<UserRoleModel>();

            foreach (var kvp in _usersCache)
            {
                var user = kvp.Value;

                // Assign role names
                List<Guid> roleIds;
                if (_userRolesCache.TryGetValue(user.UserId, out roleIds))
                {
                    user.RoleIds = roleIds;
                    user.Roles = roleIds
                        .Where(rid => _rolesCache.ContainsKey(rid))
                        .Select(rid => _rolesCache[rid])
                        .OrderBy(n => n)
                        .ToList();
                }

                users.Add(user);
            }

            return users;
        }

        #endregion

        #region Public Cache Accessors

        /// <summary>
        /// Gets all privileges for a specific user (aggregated from all their roles).
        /// </summary>
        public List<PrivilegeModel> GetUserPrivileges(Guid userId)
        {
            var privileges = new List<PrivilegeModel>();

            List<Guid> roleIds;
            if (!_userRolesCache.TryGetValue(userId, out roleIds))
                return privileges;

            foreach (var roleId in roleIds)
            {
                List<PrivilegeModel> rolePrivs;
                if (_rolePrivilegesCache.TryGetValue(roleId, out rolePrivs))
                {
                    privileges.AddRange(rolePrivs);
                }
            }

            return privileges;
        }

        /// <summary>
        /// Gets privileges grouped by role for a specific user (for drill-down display).
        /// </summary>
        public Dictionary<string, List<PrivilegeModel>> GetUserPrivilegesByRole(Guid userId)
        {
            var result = new Dictionary<string, List<PrivilegeModel>>();

            List<Guid> roleIds;
            if (!_userRolesCache.TryGetValue(userId, out roleIds))
                return result;

            foreach (var roleId in roleIds)
            {
                string roleName;
                _rolesCache.TryGetValue(roleId, out roleName);
                var key = roleName ?? roleId.ToString();

                List<PrivilegeModel> rolePrivs;
                if (_rolePrivilegesCache.TryGetValue(roleId, out rolePrivs))
                {
                    result[key] = rolePrivs;
                }
            }

            return result;
        }

        /// <summary>
        /// Returns the highest privilege access level a user has on any privilege.
        /// </summary>
        public int GetUserMaxAccessDepth(Guid userId)
        {
            var privileges = GetUserPrivileges(userId);
            return privileges.Count > 0 ? privileges.Max(p => p.AccessDepth) : 0;
        }

        public Dictionary<Guid, string> GetRolesCache() => _rolesCache;

        #endregion
    }
}
