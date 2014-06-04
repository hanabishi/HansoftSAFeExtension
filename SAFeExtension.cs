﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using HPMSdk;
using Hansoft.ObjectWrapper;

namespace Hansoft.Jean.Behavior.DeriveBehavior.Expressions
{

    public class TeamCollection
    {
        public string team = "";
        public int completedPoints = 0;
        public int completedStories = 0;
        public int totalPoints = 0;
        public int totalStories = 0;
        public HansoftEnumValue status = null;

        public void addTask(Task task)
        {
            // Once we can get aggregated functions to work this will do the trick!
            //HansoftEnumValue aggregatedStatus = task.AggregatedStatus
            HansoftEnumValue aggregatedStatus = GetAggregatedStatus(task);
            totalStories++;
            completedStories += ((aggregatedStatus.Equals(EHPMTaskStatus.Completed)) ? 1 : 0);
            status = CalculateNewStatus(task, status, aggregatedStatus);
            UpdatePoints(task);
        }

        private void UpdatePoints(Task task)
        {
            // Once we can get aggregated functions to work this will do the trick!
            //totalPoints += task.AggregatedPoints;
            //completedPoints += task.AggregatedPoints - (int)task.AggregatedWorkRemaining;
            if (task.DeepLeaves.Count > 0)
            {
                foreach (Task child in task.DeepLeaves)
                {
                    totalPoints += child.Points;
                    if (child.Status.Equals(EHPMTaskStatus.Completed))
                    {
                        completedPoints += child.Points;
                    }
                }
            }
            else
            {
                totalPoints += task.Points;
                if (task.Status.Equals(EHPMTaskStatus.Completed))
                {
                    completedPoints += task.Points;
                }
            }
        }

        private HansoftEnumValue CalculateNewStatus(Task task, HansoftEnumValue prevStatus, HansoftEnumValue newStatus)
        {
            HansoftEnumValue finalStatus = null;
            if (prevStatus == null || (newStatus.Equals(prevStatus)))
            {
                finalStatus = newStatus;
            }
            else if (newStatus.Equals(EHPMTaskStatus.Completed) || (prevStatus.Equals(EHPMTaskStatus.Blocked)))
            {
                finalStatus = prevStatus;
            }
            else if (newStatus.Equals(EHPMTaskStatus.Blocked))
            {
                finalStatus = newStatus;
            }
            else if ((prevStatus.Equals(EHPMTaskStatus.InProgress) && (newStatus.Equals(EHPMTaskStatus.NotDone))))
            {
                finalStatus = prevStatus;
            }
            else if ((!prevStatus.Equals(EHPMTaskStatus.NotDone) || (!newStatus.Equals(EHPMTaskStatus.NotDone))))
            {
                finalStatus = HansoftEnumValue.FromString(task.ProjectID, EHPMProjectDefaultColumn.ItemStatus, "In progress");
            }
            else
            {
                finalStatus = newStatus;
            }
            return finalStatus;
        }

        private HansoftEnumValue GetAggregatedStatus(Task task)
        {
            HansoftEnumValue aggregatedStatus = null;
            if (task.DeepLeaves.Count > 0)
            {
                foreach (Task child in task.DeepLeaves)
                {
                    aggregatedStatus = CalculateNewStatus(task, aggregatedStatus, child.Status);
                }
            }
            else
            {
                aggregatedStatus = task.Status;
            }

            return aggregatedStatus;
        }

        public string FormatString(string formatString)
        {
            return string.Format(formatString, new object[] { team.Substring(0, (team.Length > 20) ? 20 : team.Length), status.Text, completedPoints + "/" + totalPoints, completedStories + "/" + totalStories });
        }
    }

    public class SAFeExtension
    {
        /// <summary>
        /// Creates the ProgramFeatureSummary and updates the points value based on the linked values.
        /// </summary>
        /// <param name="current_task"></param>
        /// <param name="updatePoints">If set to true the points for the master item will be update with the aggregated points from the linked items.</param>
        /// <returns>A asci art table with containing a summary of what needs to be done.</returns>
        public static string ProgramFeatureSummary(Task current_task, bool updatePoints)
        {
            Dictionary<string, TeamCollection> teamCollection = new Dictionary<string, TeamCollection>();
            StringBuilder sb = new StringBuilder();
            int totalLinkedPoints = 0;
            foreach (Task task in current_task.LinkedTasks)
            {
                string team = task.Project.Name;
                if (!team.StartsWith("Program") && !team.StartsWith("Port"))
                {
                    Console.WriteLine("PFS - Processing " + current_task.Name + " pointing to " + task.Name);
                    if (!teamCollection.ContainsKey(team))
                    {
                        TeamCollection collection = new TeamCollection();
                        collection.team = team;
                        teamCollection.Add(team, collection);

                    }
                    teamCollection[team].addTask(task);
                    totalLinkedPoints = teamCollection[team].totalPoints;
                }
            }
            if (teamCollection.Count > 0)
            {
                string format = "<CODE>{0,-20} │ {1,-14} │ {2, -13} │ {3, -10}</CODE>";
                sb.Append(string.Format(format, new object[] { "Name", "Status", "Points", "Stories" }));
                sb.Append("\n<CODE>─────────────────────┼────────────────┼───────────────┼───────────</CODE>\n");
                bool breakOn = false;
                foreach (KeyValuePair<string, TeamCollection> pair in teamCollection)
                {
                    sb.Append(pair.Value.FormatString(format));
                    breakOn = pair.Value.totalPoints > 1;
                }

                Console.WriteLine(sb.ToString());
                if (breakOn)
                {
                    Console.WriteLine("WOW");
                }
            }
            if (updatePoints)
            {
                current_task.Points = totalLinkedPoints;
            }
            return sb.ToString();
        }

        /// <summary>
        /// Creates a feature summary suitable to display for each epic in a SAFe portfolio project where
        /// the associated features are linked to the epic they belong to.
        /// </summary>
        /// <param name="current_task">The item representing an epic in the portfolio project.</param>
        /// <returns></returns>
        public static string FeatureSummary(Task current_task, bool usePoints, string completedColumn)
        {

            StringBuilder sb = new StringBuilder();
            List<Task> featuresInDevelopment = new List<Task>();
            List<Task> featuresInReleasePlanning = new List<Task>();
            List<Task> featuresInBacklog = new List<Task>();
            foreach (Task task in current_task.LinkedTasks)
            {
                if (task.Project != current_task.Project)
                {
                    if (task.Parent.Name == "Development")
                    {
                        featuresInDevelopment.Add(task);
                    }
                    else if (task.Parent.Name == "Release planning")
                    {
                        featuresInReleasePlanning.Add(task);
                    }
                    else if (task.Parent.Name == "Feature backlog")
                    {
                        featuresInBacklog.Add(task);
                    }
                }
            }

            sb.Append(string.Format("<BOLD>Development ({0})</BOLD>", featuresInDevelopment.Count));
            sb.Append('\n');
            if (featuresInDevelopment.Count > 0)
            {
                string format = "<CODE>{0,-20} │ {1,-13} │ {2, 8} │ {3, 4} │ {4, 10} │ {5, -20}</CODE>";
                sb.Append(string.Format(format, new object[] { "Name", "Status", "Done", "↓14", "Est. Done", "Product Owner" }));
                sb.Append('\n');
                sb.Append("<CODE>─────────────────────┼───────────────┼──────────┼──────┼────────────┼─────────────────────</CODE>");
                sb.Append('\n');
                foreach (Task task in featuresInDevelopment)
                {
                    double estimate = 0;
                    if (usePoints)

                        estimate = task.AggregatedPoints;
                    else
                        estimate = task.AggregatedEstimatedDays;
                    string daysDone = task.GetCustomColumnValue(completedColumn) + "/" + estimate;
                    sb.Append(string.Format(format, new object[] { task.Name, task.AggregatedStatus, daysDone, task.GetCustomColumnValue("Velocity (14 days)"), task.GetCustomColumnValue("Estimated done"), task.GetCustomColumnValue("Product Owner") }));
                    sb.Append('\n');
                }
            }
            sb.Append('\n');

            sb.Append(string.Format("<BOLD>Release planning ({0})</BOLD>", featuresInReleasePlanning.Count));
            sb.Append('\n');
            if (featuresInReleasePlanning.Count > 0)
            {
                string format = "<CODE>{0,-20} │ {1, 5} │ {2, -20} │ {3, -20}</CODE>";
                sb.Append(string.Format(format, new object[] { "Name", "Est.", "Team", "Product Owner" }));
                sb.Append('\n');
                sb.Append("<CODE>─────────────────────┼───────┼──────────────────────┼─────────────────────</CODE>");
                sb.Append('\n');
                foreach (Task task in featuresInReleasePlanning)
                {
                    double estimate = 0;
                    if (usePoints)
                        estimate = task.AggregatedPoints;
                    else
                        estimate = task.AggregatedEstimatedDays;
                    sb.Append(string.Format(format, new object[] { task.Name, estimate, task.GetCustomColumnValue("Team"), task.GetCustomColumnValue("Product Owner") }));
                    sb.Append('\n');
                }
            }
            sb.Append('\n');

            sb.Append(string.Format("<BOLD>Feature backlog ({0})</BOLD>", featuresInBacklog.Count));
            sb.Append('\n');
            if (featuresInBacklog.Count > 0)
            {
                string format = "<CODE>{0,-20} │ {1, 5} │ {2, 8}";
                sb.Append(string.Format(format, new object[] { "Name", "Est.", "PSI" }));
                sb.Append('\n');
                sb.Append("─────────────────────┼───────┼─────────</CODE>");
                sb.Append('\n');
                foreach (Task task in featuresInBacklog)
                {
                    double estimate = 0;
                    if (usePoints)
                        estimate = task.AggregatedPoints;
                    else
                        estimate = task.AggregatedEstimatedDays;
                    sb.Append(string.Format(format, new object[] { task.Name, estimate, ListUtils.ToString(new List<HansoftItem>(task.TaggedToReleases)) }));
                    sb.Append('\n');
                }
            }
            return sb.ToString();
        }
    }


}
