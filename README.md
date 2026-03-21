# 🛡️ Role Privilege Analyzer — XrmToolBox Plugin

**Version:** 1.2026.4.1  
**Framework:** .NET Framework 4.8  
**Platform:** XrmToolBox

Enterprise-grade security audit tool for Microsoft Dynamics 365 / Dataverse. Analyzes user roles, privileges, and risk levels with a heatmap, role comparison, analytics dashboard, and Excel export.

---

## ✨ Features

### Core
- **User-Role Grid** — Virtual-mode DataGridView handling 20,000+ users without freezing
- **Privilege Risk Heatmap** — Color-coded access levels (🟢 Basic → 🔴 Global)
- **Security Risk Detection** — Automatic flagging: High Risk, Over Privileged, No Roles, Normal
- **Real-time Search** — Debounced 300ms filtering by user name, domain, business unit, or role
- **Column Sorting** — Click any header to sort ascending/descending

### Advanced
- **Privilege Drill-down** — Select any user to see Role → Privilege → Access Level breakdown
- **Role Comparison** — Compare two users: common roles, unique roles, and privilege-level differences
- **Analytics Dashboard** — Summary stats + bar chart of most-assigned roles
- **Excel Export** — Full export via ClosedXML with conditional formatting and analytics sheet
- **Open in CRM** — Double-click a user name to open their record in Dynamics 365

### Performance
- `VirtualMode = true` with `CellValueNeeded` — renders only visible rows
- All data loaded in **6 bulk queries** (no N+1)
- In-memory caching with `Dictionary<Guid, ...>` lookups
- Async/await throughout — UI never blocks
- Progress bar with step-by-step feedback

---

## 🏗️ Architecture

```
RolePrivilegeAnalyzer/
├── Models/
│   ├── UserRoleModel.cs          — User entity with roles and risk properties
│   ├── PrivilegeModel.cs         — Privilege with access level and depth mapping
│   └── RoleComparisonResult.cs   — Comparison output + analytics summary
├── Services/
│   ├── DataService.cs            — Bulk CRM data retrieval with caching
│   ├── RiskAnalysisService.cs    — Risk engine + comparison + analytics
│   └── ExportService.cs          — ClosedXML Excel export
├── UI/
│   └── RolePrivilegeAnalyzerControl.cs  — Main WinForms control (all UI)
├── Properties/
│   └── AssemblyInfo.cs
├── RolePrivilegeAnalyzerPlugin.cs       — XrmToolBox MEF entry point
└── RolePrivilegeAnalyzer.csproj
```

---


## 🎨 Risk Status Logic

| Status | Condition |
|--------|-----------|
| 🔴 High Risk | User has System Administrator, System Customizer, or Support User role |
| ⚠️ Over Privileged | User has > 5 roles OR any Global-level privilege |
| ✅ Normal | User has roles within safe thresholds |
| ❌ No Roles | User has zero role assignments |

## 🎨 Privilege Risk Level

| Level | Meaning |
|-------|---------|
| 🔴 Global | User can access all records in the organization |
| 🟠 Deep | User can access records in their BU and child BUs |
| 🟡 Local | User can access records in their business unit |
| 🟢 Basic | User can only access their own records |

---

## 📊 Data Queries

The plugin loads all data in 6 paged queries:

1. **Roles** — All `role` records (id, name)
2. **Privileges** — All `privilege` records (id, name)
3. **Role-Privilege Mappings** — All `roleprivileges` (roleid, privilegeid, depthmask)
4. **Users** — Active `systemuser` with linked `businessunit` name
5. **User-Role Assignments** — All `systemuserroles` (userid, roleid)
6. **Risk Analysis** — Computed in-memory from cached data

All queries use `PagingInfo` with `Count = 5000` to handle large environments.

---

## 📄 License (MIT)
MIT License

Copyright (c) 2026 Moamen Ashraf

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction...

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND.
