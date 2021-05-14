using Microsoft.Win32.TaskScheduler;
using SBO.Hub.Models;
using System;

namespace SBO.Hub.Helpers
{
    public class WindowsTaskScheduler
    {
        public string CreateTask(TaskSchedulerModel taskSchedulerModel)
        {
            try
            {
                TaskService ts = new TaskService();
                Task task = ts.FindTask(taskSchedulerModel.TaskName);
                if (task != null)
                {
                    ts.RootFolder.DeleteTask(taskSchedulerModel.TaskName);
                }

                TaskDefinition td = ts.NewTask();
                td.RegistrationInfo.Description = taskSchedulerModel.Description;
                td.Settings.Hidden = taskSchedulerModel.Hidden;

                Trigger trigger = new DailyTrigger();
                switch (taskSchedulerModel.TaskSchedulerFrequencyEnum)
                {
                    case Enums.TaskSchedulerFrequencyEnum.Secondly:
                    case Enums.TaskSchedulerFrequencyEnum.Minutely:
                    case Enums.TaskSchedulerFrequencyEnum.Hourly:
                    case Enums.TaskSchedulerFrequencyEnum.Dayly:
                        trigger = new DailyTrigger();
                        switch (taskSchedulerModel.TaskSchedulerFrequencyEnum)
                        {
                            case Enums.TaskSchedulerFrequencyEnum.Secondly:
                                trigger.Repetition.Interval = TimeSpan.FromSeconds(taskSchedulerModel.Interval);
                                break;
                            case Enums.TaskSchedulerFrequencyEnum.Minutely:
                                trigger.Repetition.Interval = TimeSpan.FromMinutes(taskSchedulerModel.Interval);
                                break;
                            case Enums.TaskSchedulerFrequencyEnum.Hourly:
                                trigger.Repetition.Interval = TimeSpan.FromHours(taskSchedulerModel.Interval);
                                break;
                        }

                        break;
                    case Enums.TaskSchedulerFrequencyEnum.Weekly:
                        trigger = new WeeklyTrigger(weeksInterval: (short)taskSchedulerModel.Interval);
                        break;
                    case Enums.TaskSchedulerFrequencyEnum.Monthly:
                        trigger = new MonthlyTrigger();
                        break;
                }

                if (taskSchedulerModel.StartDate.HasValue)
                {
                    trigger.StartBoundary = taskSchedulerModel.StartDate.Value;
                }
                td.Triggers.Add(trigger);
                td.Actions.Add(new ExecAction(taskSchedulerModel.FilePath, taskSchedulerModel.Arguments, null));
                ts.RootFolder.RegisterTaskDefinition(taskSchedulerModel.TaskName, td);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            { }
            return String.Empty;
        }
    }
}
