using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using HPMSdk;
using Hansoft.ObjectWrapper;
using Hansoft.ObjectWrapper.CustomColumnValues;

namespace Hansoft.Jean.Behavior.DeriveBehavior.Expressions
{

    public class TaskCollection
    {
        public int totalPoints = 0;
        public int completedPoints = 0;
        public string status = "";
        public List<string> taskHeaders = new List<string>();
        private string formatString = "<CODE>{0,-36} │ {1, -7}</CODE>";
        public StringBuilder detailedDescription = new StringBuilder();
        public TaskCollection(string status, int totalPoints, int completedPoints)
        {
            this.totalPoints = totalPoints;
            this.completedPoints = completedPoints;
            this.status = status;
            detailedDescription.Append(string.Format(formatString, new object[] { "Name", "Points" }));
            detailedDescription.Append("\n<CODE>─────────────────────────────────────┼─────────</CODE>\n");
        }

        public void addTaskInformation(Task task)
        {
            taskHeaders.Add(task.Name);
            this.totalPoints += task.Points;
            this.completedPoints = ((task.Status.Equals("Completed")) ? task.Points : 0);
            string shortName = task.Name.Substring(0, (task.Name.Length > 36) ? 35 : task.Name.Length) + ((task.Name.Length > 36) ? "…" : "");
            detailedDescription.Append(string.Format(formatString, new object[] { shortName, task.Points }) + "\n");
        }
    }

    public class TeamCollection
    {
        public string team = "";
        public int completedPoints = 0;
        public int completedStories = 0;
        public int totalPoints = 0;
        public int totalStories = 0;
        public HansoftEnumValue status = null;

        public Dictionary<string, TaskCollection> taskGroup = null;
        public TeamCollection(Dictionary<string, TaskCollection> taskGroup)
        {
            this.taskGroup = taskGroup;
        }
        public void addTask(Task task)
        {
            if (task.DeepLeaves.Count > 0)
            {
                foreach (Task child in task.DeepLeaves)
                {
                    totalStories += 1;
                    completedStories += ((child.Status.Equals(EHPMTaskStatus.Completed)) ? 1 : 0);

                    totalPoints += child.Points;
                    completedPoints += ((child.Status.Equals(EHPMTaskStatus.Completed)) ? child.Points : 0);

                    status = CalculateNewStatus(child, status, child.Status);
                    if (!taskGroup.ContainsKey(child.Status.Text))
                    {
                        taskGroup.Add(child.Status.Text, new TaskCollection(child.Status.Text, 0, 0));
                    }
                    taskGroup[child.Status.Text].addTaskInformation(child);
                }
            }
            else
            {
                totalStories += 1;
                completedStories += ((task.Status.Equals(EHPMTaskStatus.Completed)) ? 1 : 0);

                totalPoints += task.Points;
                completedPoints += ((task.Status.Equals(EHPMTaskStatus.Completed)) ? task.Points : 0);
                status = task.Status;
                if (!taskGroup.ContainsKey(status.Text))
                {
                    taskGroup.Add(status.Text, new TaskCollection(status.Text, 0, 0));
                }
                taskGroup[status.Text].addTaskInformation(task);
            }
        }

        public static HansoftEnumValue CalculateNewStatus(Task task, HansoftEnumValue prevStatus, HansoftEnumValue newStatus)
        {
            HansoftEnumValue finalStatus = null;
            // If either prevStatus or newStatus is null, the result is "the other value" (which can be null)
            if (prevStatus == null)
            {
                finalStatus = newStatus;
            }
            else if (newStatus == null)
            {
                finalStatus = prevStatus;
            }
            else if (prevStatus.Equals(newStatus))
            {
                finalStatus = prevStatus;
            }
            // If any story is Blocked the feature is Blocked
            else if (prevStatus.Text.Equals("Blocked") || newStatus.Text.Equals("Blocked"))
            {
                finalStatus = HansoftEnumValue.FromString(task.ProjectID, EHPMProjectDefaultColumn.ItemStatus, "Blocked");
            }
            // For all other combinations the result is InProgress
            else
            {
                finalStatus = HansoftEnumValue.FromString(task.ProjectID, EHPMProjectDefaultColumn.ItemStatus, "In progress");
            }
            return finalStatus;
        }

        public string FormatString(string formatString)
        {
            string teamShort = team.Substring(0, (team.Length > 20) ? 19 : team.Length) + ((team.Length > 20) ? "…" : "");
            return string.Format(formatString, new object[] { teamShort, status.Text, completedPoints + "/" + totalPoints, completedStories + "/" + totalStories });
        }
    }

    public class SAFeExtension
    {
        public static bool debug = false;

        public static Task createNewTask(Task parent, TaskCollection taskCollection)
        {
            Task newTask = null;
            foreach (Task subtask in parent.Leaves)
            {
                if (subtask.Name.Equals(taskCollection.status))
                {
                    newTask = subtask;
                }
            }
            if (newTask == null && taskCollection.taskHeaders.Count > 0)
            {
                HPMTaskCreateUnifiedReference parentRefId = new HPMTaskCreateUnifiedReference();
                parentRefId.m_RefID = parent.UniqueID;
                parentRefId.m_bLocalID = false;

                HPMTaskCreateUnified createData = new HPMTaskCreateUnified();
                createData.m_Tasks = new HPMTaskCreateUnifiedEntry[1];
                createData.m_Tasks[0] = new HPMTaskCreateUnifiedEntry();
                createData.m_Tasks[0].m_bIsProxy = false;
                createData.m_Tasks[0].m_LocalID = 1;
                createData.m_Tasks[0].m_NonProxy_WorkflowID = 0xffffffff;
                createData.m_Tasks[0].m_ParentRefIDs = new HPMTaskCreateUnifiedReference[1];
                createData.m_Tasks[0].m_ParentRefIDs[0] = parentRefId;
                createData.m_Tasks[0].m_PreviousRefID = new HPMTaskCreateUnifiedReference();
                createData.m_Tasks[0].m_PreviousRefID.m_RefID = -1;
                createData.m_Tasks[0].m_PreviousWorkPrioRefID = new HPMTaskCreateUnifiedReference();
                createData.m_Tasks[0].m_PreviousWorkPrioRefID.m_RefID = -2;
                createData.m_Tasks[0].m_TaskLockedType = EHPMTaskLockedType.BacklogItem;
                createData.m_Tasks[0].m_TaskType = EHPMTaskType.Planned;
                HPMChangeCallbackData_TaskCreateUnified result = SessionManager.Session.TaskCreateUnifiedBlock(parent.ProjectID.m_ID, createData);

                SessionManager.Session.TaskSetFullyCreated(SessionManager.Session.TaskRefGetTask(result.m_Tasks[0].m_TaskRefID));
                newTask = Task.GetTask(result.m_Tasks[0].m_TaskRefID);
                newTask.Name = taskCollection.status;
            }
            if (taskCollection.taskHeaders.Count > 0)
            {
                //if (!newTask.GetCustomColumnValue("Task summary").Equals(taskCollection.detailedDescription.ToString()))
                //{
                //    newTask.SetCustomColumnValue("Task summary", taskCollection.detailedDescription.ToString());
                //}

                if (parent.Points > 0)
                {
                    parent.Points = 0;
                }
                if (newTask.Points != taskCollection.totalPoints)
                {
                    newTask.Points = taskCollection.totalPoints;
                    newTask.WorkRemaining = taskCollection.totalPoints - taskCollection.completedPoints;
                }
                if (!newTask.Status.Text.Equals(taskCollection.status))
                {
                    newTask.Status = HansoftEnumValue.FromString(parent.ProjectID, EHPMProjectDefaultColumn.ItemStatus, taskCollection.status);
                    try
                    {
                        SessionManager.Session.TaskSetCompleted(newTask.UniqueTaskID, taskCollection.status.Equals("Completed"), true);
                    }
                    catch (Exception epeen)
                    {
                    }
                }
            }
            else if (newTask != null)
            {
                try
                {
                    SessionManager.Session.TaskDelete(newTask.UniqueTaskID);
                }
                catch (Exception epeen)
                {
                }
            }

            return newTask;
        }

        /// <summary>
        /// Creates the ProgramFeatureSummary and updates the points value based on the linked values.
        /// </summary>
        /// <param name="current_task"></param>
        /// <param name="updateTaskStatus">If set to true the points for the master item will be update with the aggregated points from the linked items.</param>
        /// <returns>A asci art table with containing a summary of what needs to be done.</returns>
        public static string ProgramFeatureSummary(Task current_task, bool updateTaskStatus)
        {
            Dictionary<string, TeamCollection> teamCollection = new Dictionary<string, TeamCollection>();
            Dictionary<string, TaskCollection> taskGroup = new Dictionary<string, TaskCollection>();
            taskGroup.Add("In progress", new TaskCollection("In progress", 0, 0));
            taskGroup.Add("Completed", new TaskCollection("Completed", 0, 0));
            taskGroup.Add("Blocked", new TaskCollection("Blocked", 0, 0));
            taskGroup.Add("Not done", new TaskCollection("Not done", 0, 0));
            taskGroup.Add("To be deleted", new TaskCollection("To be deleted", 0, 0));
            StringBuilder sb = new StringBuilder();
            foreach (Task task in current_task.LinkedTasks)
            {
                string team = task.Project.Name;
                if (!team.StartsWith("Program") && !team.StartsWith("Port"))
                {
                    if (!teamCollection.ContainsKey(team))
                    {
                        TeamCollection collection = new TeamCollection(taskGroup);
                        collection.team = team;
                        teamCollection.Add(team, collection);

                    }
                    teamCollection[team].addTask(task);
                }
            }
            if (teamCollection.Count > 0)
            {
                string format = "<CODE>{0,-20} │ {1,-14} │ {2, -13} │ {3, -10}</CODE>";
                sb.Append(string.Format(format, new object[] { "Name", "Status", "Points", "Stories" }));
                sb.Append("\n<CODE>─────────────────────┼────────────────┼───────────────┼───────────</CODE>\n");
                foreach (KeyValuePair<string, TeamCollection> pair in teamCollection)
                {
                    sb.Append(pair.Value.FormatString(format) + "\n");
                }
                foreach (KeyValuePair<string, TaskCollection> taskPair in taskGroup)
                {
                    createNewTask(current_task, taskPair.Value);
                }
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
                    string taskShort = task.Name.Substring(0, (task.Name.Length > 20) ? 19 : task.Name.Length) + ((task.Name.Length > 20) ? "…" : "");
                    sb.Append(string.Format(format, new object[] { taskShort, task.AggregatedStatus, daysDone, task.GetCustomColumnValue("Velocity (14 days)"), task.GetCustomColumnValue("Estimated done"), task.GetCustomColumnValue("Product Owner") }));
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
                    string taskShort = task.Name.Substring(0, (task.Name.Length > 20) ? 19 : task.Name.Length) + ((task.Name.Length > 20) ? "…" : "");
                    sb.Append(string.Format(format, new object[] { taskShort, estimate, task.GetCustomColumnValue("Team"), task.GetCustomColumnValue("Product Owner") }));
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
                    string taskShort = task.Name.Substring(0, (task.Name.Length > 20) ? 19 : task.Name.Length) + ((task.Name.Length > 20) ? "…" : "");
                    sb.Append(string.Format(format, new object[] { taskShort, estimate, ListUtils.ToString(new List<HansoftItem>(task.TaggedToReleases)) }));
                    sb.Append('\n');
                }
            }
            return sb.ToString();
        }
    }


}

