using SBO.Hub.Enums;
using System;

namespace SBO.Hub.Models
{
    public class TaskSchedulerModel
    {
        public string Server { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public string Domain { get; set; }
        public string TaskName { get; set; }
        public string Description { get; set; }
        public bool Hidden { get; set; }
        public string FilePath { get; set; }
        public string Arguments { get; set; }
        public string Frequency { get; set; }
        public int Interval { get; set; }
        public DateTime? StartDate { get; set; }

        public TaskSchedulerFrequencyEnum TaskSchedulerFrequencyEnum
        {
            get
            {
                switch (Frequency)
                {
                    case "S":
                        return TaskSchedulerFrequencyEnum.Secondly;
                    case "M":
                        return TaskSchedulerFrequencyEnum.Minutely;
                    case "H":
                        return TaskSchedulerFrequencyEnum.Hourly;
                    case "D":
                        return TaskSchedulerFrequencyEnum.Dayly;
                    case "W":
                        return TaskSchedulerFrequencyEnum.Weekly;
                    case "O":
                        return TaskSchedulerFrequencyEnum.Monthly;
                }
                return TaskSchedulerFrequencyEnum.None;
            }
        }
    }
}
