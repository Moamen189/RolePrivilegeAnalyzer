using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ClosedXML.Excel;
using RolePrivilegeAnalyzer.Models;

namespace RolePrivilegeAnalyzer.Services
{
    /// <summary>
    /// Exports analysis results to Excel using ClosedXML.
    /// </summary>
    public class ExportService
    {
        /// <summary>
        /// Export user role data to an Excel file.
        /// </summary>
        public async Task ExportToExcelAsync(
            List<UserRoleModel> users,
            AnalyticsSummary analytics,
            string filePath)
        {
            await Task.Run(() =>
            {
                using (var workbook = new XLWorkbook())
                {
                    // ── Sheet 1: User Roles ──
                    var wsUsers = workbook.Worksheets.Add("User Roles");

                    // Headers
                    wsUsers.Cell(1, 1).Value = "Business Unit";
                    wsUsers.Cell(1, 2).Value = "Full Name";
                    wsUsers.Cell(1, 3).Value = "Domain / Email";
                    wsUsers.Cell(1, 4).Value = "Roles";
                    wsUsers.Cell(1, 5).Value = "Role Count";
                    wsUsers.Cell(1, 6).Value = "Privilege Risk Level";
                    wsUsers.Cell(1, 7).Value = "Risk Status";

                    var headerRange = wsUsers.Range(1, 1, 1, 7);
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#1B3A5C");
                    headerRange.Style.Font.FontColor = XLColor.White;
                    headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    // Data rows
                    for (int i = 0; i < users.Count; i++)
                    {
                        var user = users[i];
                        int row = i + 2;

                        wsUsers.Cell(row, 1).Value = user.BusinessUnit;
                        wsUsers.Cell(row, 2).Value = user.FullName;
                        wsUsers.Cell(row, 3).Value = user.DomainName;
                        wsUsers.Cell(row, 4).Value = string.Join(", ", user.Roles);
                        wsUsers.Cell(row, 5).Value = user.Roles.Count;
                        wsUsers.Cell(row, 6).Value = user.RiskLevel;
                        wsUsers.Cell(row, 7).Value = user.RiskStatus;

                        // Conditional coloring on risk status
                        if (user.RiskStatus.Contains("High Risk"))
                        {
                            wsUsers.Cell(row, 7).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFD7D7");
                            wsUsers.Cell(row, 7).Style.Font.FontColor = XLColor.DarkRed;
                        }
                        else if (user.RiskStatus.Contains("Over Privileged"))
                        {
                            wsUsers.Cell(row, 7).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFF3CD");
                            wsUsers.Cell(row, 7).Style.Font.FontColor = XLColor.FromHtml("#856404");
                        }
                        else if (user.RiskStatus.Contains("No Roles"))
                        {
                            wsUsers.Cell(row, 7).Style.Fill.BackgroundColor = XLColor.FromHtml("#F0F0F0");
                            wsUsers.Cell(row, 7).Style.Font.FontColor = XLColor.Gray;
                        }
                    }

                    wsUsers.Columns().AdjustToContents(1, 100);
                    wsUsers.SheetView.FreezeRows(1);

                    // ── Sheet 2: Analytics Summary ──
                    if (analytics != null)
                    {
                        var wsAnalytics = workbook.Worksheets.Add("Analytics");

                        wsAnalytics.Cell(1, 1).Value = "Metric";
                        wsAnalytics.Cell(1, 2).Value = "Value";
                        var aHeader = wsAnalytics.Range(1, 1, 1, 2);
                        aHeader.Style.Font.Bold = true;
                        aHeader.Style.Fill.BackgroundColor = XLColor.FromHtml("#1B3A5C");
                        aHeader.Style.Font.FontColor = XLColor.White;

                        wsAnalytics.Cell(2, 1).Value = "Total Users";
                        wsAnalytics.Cell(2, 2).Value = analytics.TotalUsers;
                        wsAnalytics.Cell(3, 1).Value = "Total Roles";
                        wsAnalytics.Cell(3, 2).Value = analytics.TotalRoles;
                        wsAnalytics.Cell(4, 1).Value = "High Risk Users";
                        wsAnalytics.Cell(4, 2).Value = analytics.HighRiskUsers;
                        wsAnalytics.Cell(5, 1).Value = "Over Privileged Users";
                        wsAnalytics.Cell(5, 2).Value = analytics.OverPrivilegedUsers;
                        wsAnalytics.Cell(6, 1).Value = "No Role Users";
                        wsAnalytics.Cell(6, 2).Value = analytics.NoRoleUsers;
                        wsAnalytics.Cell(7, 1).Value = "Normal Users";
                        wsAnalytics.Cell(7, 2).Value = analytics.NormalUsers;

                        // Most assigned roles
                        int startRow = 9;
                        wsAnalytics.Cell(startRow, 1).Value = "Most Assigned Role";
                        wsAnalytics.Cell(startRow, 2).Value = "User Count";
                        var rHeader = wsAnalytics.Range(startRow, 1, startRow, 2);
                        rHeader.Style.Font.Bold = true;
                        rHeader.Style.Fill.BackgroundColor = XLColor.FromHtml("#2C5F8A");
                        rHeader.Style.Font.FontColor = XLColor.White;

                        for (int i = 0; i < analytics.MostAssignedRoles.Count; i++)
                        {
                            var r = analytics.MostAssignedRoles[i];
                            wsAnalytics.Cell(startRow + 1 + i, 1).Value = r.RoleName;
                            wsAnalytics.Cell(startRow + 1 + i, 2).Value = r.Count;
                        }

                        wsAnalytics.Columns().AdjustToContents();
                    }

                    workbook.SaveAs(filePath);
                }
            });
        }
    }
}
