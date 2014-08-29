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
using System.Collections;

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
        public SortedSet<string> plannedSprints = new SortedSet<string>();
        public Dictionary<string, TaskCollection> taskGroup = null;
        public TeamCollection(Dictionary<string, TaskCollection> taskGroup)
        {
            this.taskGroup = taskGroup;
        }
        public void addTask(Task task)
        {
            CustomColumnValue val = task.GetCustomColumnValue("Planned sprint");
            if (val != null && !String.IsNullOrEmpty(val.ToString()))
            {
                plannedSprints.Add(val.ToString());
            }

            if (task.DeepLeaves.Count > 0)
            {
                foreach (Task child in task.DeepLeaves)
                {
                    updateFromTask(child);
                }
            }
            else
            {
                updateFromTask(task);
            }
        }

        private void updateFromTask(Task task)
        {
            totalStories += 1;
            completedStories += ((task.Status.Equals(EHPMTaskStatus.Completed)) ? 1 : 0);

            totalPoints += task.Points;
            completedPoints += ((task.Status.Equals(EHPMTaskStatus.Completed)) ? task.Points : 0);
            HansoftEnumValue taskStatus = getAggregatedStatus(task);
            status = CalculateNewStatus(task, status, taskStatus);

            if (!taskGroup.ContainsKey(taskStatus.Text))
            {
                taskGroup.Add(taskStatus.Text, new TaskCollection(taskStatus.Text, 0, 0));
            }
            taskGroup[taskStatus.Text].addTaskInformation(task);
        }
        private static HansoftEnumValue getAggregatedStatus(Task task)
        {
            HPMTaskSummary summaryStatus = SessionManager.Session.TaskRefGetSummary(task.Id);
            if (task.Status.Text == task.AggregatedStatus.Text)
            {
                return task.Status;
            }
            
            string statusString = "Not Done";
            switch (summaryStatus.m_TaskStatus)
            {
                case EHPMTaskStatus.Blocked: statusString = "Blocked"; break;
                case EHPMTaskStatus.Completed: statusString = "Completed"; break;
                case EHPMTaskStatus.Deleted: statusString = "Deleted"; break;
                case EHPMTaskStatus.InProgress: statusString = "In Progress"; break;
                case EHPMTaskStatus.NewVersionOfSDKRequired: statusString = "New Version Of SDK Required"; break;
                case EHPMTaskStatus.NoStatus: statusString = "No Status"; break;
                case EHPMTaskStatus.NotDone: statusString = "Not Done"; break;
            }
            
            return HansoftEnumValue.FromString(task.ProjectID, EHPMProjectDefaultColumn.ItemStatus, statusString);
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
            else if (prevStatus.Text.Equals(newStatus.Text))
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
            return string.Format(formatString, new object[] { teamShort, status.Text, completedPoints + "/" + totalPoints, completedStories + "/" + totalStories, String.Join(", ", plannedSprints) });
        }
    }

    public class SAFeExtension
    {
        public static bool debug = false;

        public static void CreateProxyItem(SprintSearchCollection SprintRefID, HPMUniqueID ProjectID, HPMUniqueID masterRefID)
        {
            HPMTaskCreateUnifiedReference parent = new HPMTaskCreateUnifiedReference();
            parent.m_bLocalID = false;
            parent.m_RefID = SprintRefID.childID;

            HPMTaskCreateUnifiedReference prevTaskID = new HPMTaskCreateUnifiedReference();
            prevTaskID.m_bLocalID = false;
            prevTaskID.m_RefID = SprintRefID.childID;

            HPMTaskCreateUnified ProxyTaskCreate = new HPMTaskCreateUnified();
            ProxyTaskCreate.m_Tasks = new HPMTaskCreateUnifiedEntry[1];
            ProxyTaskCreate.m_Tasks[0] = new HPMTaskCreateUnifiedEntry();
            ProxyTaskCreate.m_Tasks[0].m_bIsProxy = true;
            ProxyTaskCreate.m_Tasks[0].m_LocalID = 0;
            ProxyTaskCreate.m_Tasks[0].m_TaskType = EHPMTaskType.Planned;
            ProxyTaskCreate.m_Tasks[0].m_TaskLockedType = EHPMTaskLockedType.BacklogItem;
            ProxyTaskCreate.m_Tasks[0].m_ParentRefIDs = new HPMTaskCreateUnifiedReference[1];
            ProxyTaskCreate.m_Tasks[0].m_ParentRefIDs[0] = parent;
            ProxyTaskCreate.m_Tasks[0].m_PreviousWorkPrioRefID = new HPMTaskCreateUnifiedReference();
            ProxyTaskCreate.m_Tasks[0].m_PreviousWorkPrioRefID.m_RefID = -2;
            ProxyTaskCreate.m_Tasks[0].m_Proxy_ReferToRefTaskID = masterRefID;

            prevTaskID.m_bLocalID = true;
            prevTaskID.m_RefID = 0;

            HPMChangeCallbackData_TaskCreateUnified proxyResult = SessionManager.Session.TaskCreateUnifiedBlock(ProjectID, ProxyTaskCreate);
        }

        //public static Task CreateTask(Task parent, HPMUniqueID ProjectID, string status, HPMUniqueID sprintID)
        public static Task CreateTask(Task parentTask, HPMUniqueID ProjectID, string newTaskName, SprintSearchCollection sprintSearchResult)
        {
            Task newTask = null;
            try
            {
                HPMUniqueID backlogProjectID = SessionManager.Session.ProjectOpenBacklogProject(ProjectID);

                HPMTaskCreateUnifiedReference parentRefId = new HPMTaskCreateUnifiedReference();
                parentRefId.m_RefID = parentTask.UniqueID;
                parentRefId.m_bLocalID = false;

                HPMTaskCreateUnified createData = new HPMTaskCreateUnified();
                createData.m_Tasks = new HPMTaskCreateUnifiedEntry[1];
                createData.m_Tasks[0] = new HPMTaskCreateUnifiedEntry();
                createData.m_Tasks[0].m_bIsProxy = false;
                createData.m_Tasks[0].m_LocalID = 0;
                createData.m_Tasks[0].m_ParentRefIDs = new HPMTaskCreateUnifiedReference[1];
                createData.m_Tasks[0].m_ParentRefIDs[0] = parentRefId;
                createData.m_Tasks[0].m_PreviousRefID = new HPMTaskCreateUnifiedReference();
                createData.m_Tasks[0].m_PreviousRefID.m_RefID = -1;
                createData.m_Tasks[0].m_PreviousWorkPrioRefID = new HPMTaskCreateUnifiedReference();
                createData.m_Tasks[0].m_PreviousWorkPrioRefID.m_RefID = -2;
                createData.m_Tasks[0].m_TaskLockedType = EHPMTaskLockedType.BacklogItem;
                createData.m_Tasks[0].m_TaskType = EHPMTaskType.Planned;
                HPMChangeCallbackData_TaskCreateUnified result = SessionManager.Session.TaskCreateUnifiedBlock(backlogProjectID, createData);
                SessionManager.Session.TaskSetFullyCreated(SessionManager.Session.TaskRefGetTask(result.m_Tasks[0].m_TaskRefID));
                newTask = Task.GetTask(result.m_Tasks[0].m_TaskRefID);
                HPMUniqueID masterRefID = newTask.UniqueID;
                newTask.Name = newTaskName;
                if (sprintSearchResult.childID != null)
                {
                    CreateProxyItem(sprintSearchResult, ProjectID, masterRefID);
                }
            }
            catch (Exception)
            {
                if (newTask != null)
                {
                    SessionManager.Session.TaskDelete(newTask.UniqueTaskID);
                }
            }

            return newTask;
        }

        public struct SprintSearchCollection
        {
            public HPMUniqueID childID;
            public HPMUniqueID sprintID;
            public SprintSearchCollection(HPMUniqueID childID, HPMUniqueID sprintID)
            {
                this.childID = childID;
                this.sprintID = sprintID;
            }
        }

        public static SprintSearchCollection findSprintTaskID(Task parentTask)
        {
            HPMFindContext FindContext = new HPMFindContext();

            HPMFindContextData FindContextData = SessionManager.Session.UtilPrepareFindContext("Itemname: \"Program - PI\" AND ! Itemname: \":\"", parentTask.Project.UniqueID, EHPMReportViewType.AgileMainProject, FindContext);
            HPMTaskEnum SprintIDEnum = SessionManager.Session.TaskFind(FindContextData, EHPMTaskFindFlag.None);

            foreach (HPMUniqueID searchID in SprintIDEnum.m_Tasks)
            {
                HPMUniqueID SprintRefID = SessionManager.Session.TaskGetMainReference(searchID);
                Task sprint = Task.GetTask(SprintRefID);
                foreach (Task child in sprint.DeepChildren)
                {
                    if (child.UniqueTaskID == parentTask.UniqueTaskID)
                    {
                        return new SprintSearchCollection(child.UniqueID, SprintRefID);
                    }
                }
            }

            return new SprintSearchCollection(null, null);
        }

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
                SprintSearchCollection searchResult = findSprintTaskID(parent);
                newTask = CreateTask(parent, parent.Project.UniqueID, taskCollection.status, searchResult);
                if (newTask == null)
                {
                    return null;
                }
            }

            if (taskCollection.taskHeaders.Count > 0)
            {
                if (!newTask.GetCustomColumnValue("Item count").Equals(taskCollection.taskHeaders.Count.ToString()))
                {
                    newTask.SetCustomColumnValue("Item count", taskCollection.taskHeaders.Count.ToString());
                }

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
                    catch (Exception)
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
                catch (Exception)
                {
                }
            }

            return newTask;
        }

        private static bool hasParentHeader(HansoftItem current_task, String parentFolder)
        {
            HansoftItem parent = current_task.Parent;
            if (current_task == null || parent == null)
            {
                return false;
            }
            if (parent.Name.ToLower() == "product backlog" && current_task.Name.ToLower() == parentFolder.ToLower())
            {
                return true;
            }
            else
            {
                return hasParentHeader(parent, parentFolder);
            }
        }

        /// <summary>
        /// Creates the ProgramFeatureSummary and updates the points value based on the linked values.
        /// </summary>
        /// <param name="current_task"></param>
        /// <param name="updateTaskStatus">If set to true the points for the master item will be update with the aggregated points from the linked items.</param>
        /// <returns>A ASCI art table with containing a summary of what needs to be done.</returns>
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
                if (team.ToLower().StartsWith("team - "))
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
                string format = "<CODE>{0,-20} │ {1,-11} │ {2, -6} │ {3, -8}│ {4, -15}</CODE>";
                sb.Append(string.Format(format, new object[] { "Name", "Status", "Points", "Stories", "Planned sprint" }));
                sb.Append("\n<CODE>─────────────────────┼─────────────┼────────┼─────────┼───────────────────</CODE>\n");
                foreach (KeyValuePair<string, TeamCollection> pair in teamCollection)
                {
                    sb.Append(pair.Value.FormatString(format) + "\n");
                }

                if (hasParentHeader(current_task.Parent, "Development"))
                {
                    foreach (KeyValuePair<string, TaskCollection> taskPair in taskGroup)
                    {
                        createNewTask(current_task, taskPair.Value);
                    }
                }

                try
                {
                    CustomColumnValue v = current_task.GetCustomColumnValue("Team");
                    // Intead of creating the list I jut simply get the existing list and clear it
                    IList selectedTeams = v.ToStringList();
                    selectedTeams.Clear();
                    foreach (KeyValuePair<string, TeamCollection> pair in teamCollection)
                    {
                        string name = pair.Key.Substring(7);
                        if (!selectedTeams.Contains(name))
                        {
                            selectedTeams.Add(name);
                        }
                    }
                    CustomColumnValue.FromStringList(current_task, current_task.ProjectView.GetCustomColumn("Team"), selectedTeams);
                    CustomColumnValue newValue = CustomColumnValue.FromStringList(current_task, current_task.ProjectView.GetCustomColumn("Team"), selectedTeams);
                    current_task.SetCustomColumnValue("Team", newValue);
                }
                catch (Exception)
                {

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
                    if (hasParentHeader(task, "Development"))
                    {
                        featuresInDevelopment.Add(task);
                    }
                    else if (hasParentHeader(task, "Release planning"))
                    {
                        featuresInReleasePlanning.Add(task);
                    }
                    else if (hasParentHeader(task, "Feature backlog"))
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

