using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using Kendo.Mvc.UI;

namespace OWL_Site.Models
{
    public class MeetingViewModel : ISchedulerEvent
    {
        public int MeetingID { get; set; }

        [Required]
        public string Title { get; set; }

        public string Description { get; set; }

        private DateTime start;
        [Required]
        public DateTime Start
        {
            get
            {
                return start;
            }
            set
            {
                start = value.ToUniversalTime();
            }
        }

        public string StartTimezone { get; set; }

        private DateTime end;

        [Required]
        public DateTime End
        {
            get
            {
                return end;
            }
            set
            {
                end = value.ToUniversalTime();
            }
        }

        public string EndTimezone { get; set; }

        public string RecurrenceRule { get; set; }
        public int? RecurrenceID { get; set; }
        public string RecurrenceException { get; set; }
        public bool IsAllDay { get; set; }
        //public string Timezone { get; set; }
        public int RoomID { get; set; }
        //public int? StylistID { get; set; }
        public string OpLink { get; set; }
        public IEnumerable<string> Attendees { get; set; }
        public string AddAttend { get; set; }
        public string FileLink { get; set; }
        public bool Record { get; set; }
        public string Recfile { get; set; }
        public string InitName { get; set; }
        public string FName { get; set; }

        public Meeting ToEntity()
        {
            var meeting = new Meeting
            {
                MeetingID = MeetingID,
                Title = Title,
                Start = Start,
                StartTimezone = StartTimezone,
                End = End,
                EndTimezone = EndTimezone,
                Description = Description,
                IsAllDay = IsAllDay,
                RecurrenceRule = RecurrenceRule,
                RecurrenceException = RecurrenceException,
                RecurrenceID = RecurrenceID,
                RoomID = RoomID,
                Oplink = OpLink,
                AddAttend = AddAttend,
                FileLink = FileLink,
                Record = Record,
                Recfile = Recfile,
                InitName = InitName,
                FName = FName
            };

            return meeting;
        }


    }
}