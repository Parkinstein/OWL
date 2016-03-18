using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Diagnostics;
using System.Linq;
using System.Web.Mvc;
using Kendo.Mvc.Extensions;

namespace OWL_Site.Models
{
    public class SchedulerMeetingService
    {
        private aspnetdbEntities db;

        public SchedulerMeetingService(aspnetdbEntities context)
        {
            db = context;
        }

        public SchedulerMeetingService()
            : this(new aspnetdbEntities())
        {
        }

        public IQueryable<MeetingViewModel> GetAll()
        {
            return db.Meetings.ToList().Select(meeting => new MeetingViewModel
            {
                MeetingID = meeting.MeetingID,  //.MeetingID,
                Title = meeting.Title,
                Start = DateTime.SpecifyKind(meeting.Start, DateTimeKind.Utc),
                End = DateTime.SpecifyKind(meeting.End, DateTimeKind.Utc),
                StartTimezone = meeting.StartTimezone,
                EndTimezone = meeting.EndTimezone,
                Description = meeting.Description,
                IsAllDay = meeting.IsAllDay,
                RoomID = (int)meeting.RoomID,
                RecurrenceRule = meeting.RecurrenceRule,
                RecurrenceException = meeting.RecurrenceException,
                RecurrenceID = meeting.RecurrenceID,
                Attendees = meeting.MeetingAttendees.Select(m => m.AttendeeID).ToArray(),
                OpLink = meeting.Oplink,
                AddAttend = meeting.AddAttend,
                FileLink = meeting.FileLink,
                Record = meeting.Record,
                Recfile = meeting.Recfile,
                InitName = meeting.InitName,
                FName = meeting.FName

            }).AsQueryable();
        }
        



        public void Insert(MeetingViewModel meeting, ModelStateDictionary modelState)
        {
            if (ValidateModel(meeting, modelState))
            {
                if (meeting.Attendees == null)
                {
                    meeting.Attendees = new string[0];
                }

                if (string.IsNullOrEmpty(meeting.Title))
                {
                    meeting.Title = "";
                }

                var entity = meeting.ToEntity();
                entity.RoomID = meeting.RoomID;
                foreach (var attendeeId in meeting.Attendees)
                {
                    MeetingAttendee attendee = new MeetingAttendee();
                    attendee.MeetingID = meeting.MeetingID;
                    attendee.AttendeeID = attendeeId;
                    db.MeetingAttendees.Add(attendee);
                }
                db.Meetings.Add(entity);  
                db.SaveChanges();
                meeting.MeetingID = entity.MeetingID;
            }
        }

        public void Update(MeetingViewModel meeting, ModelStateDictionary modelState)
        {
            if (ValidateModel(meeting, modelState))
            {
                if (string.IsNullOrEmpty(meeting.Title))
                {
                    meeting.Title = "";
                }
                var entity = db.Meetings.FirstOrDefault(m => m.MeetingID == meeting.MeetingID);
                entity.Title = meeting.Title;
                entity.Start = meeting.Start;
                entity.End = meeting.End;
                entity.Description = meeting.Description;
                entity.IsAllDay = meeting.IsAllDay;
                entity.RoomID = meeting.RoomID;
                entity.RecurrenceID = meeting.RecurrenceID;
                entity.RecurrenceRule = meeting.RecurrenceRule;
                entity.RecurrenceException = meeting.RecurrenceException;
                entity.StartTimezone = meeting.StartTimezone;
                entity.EndTimezone = meeting.EndTimezone;
                entity.Oplink = meeting.OpLink;
                entity.AddAttend = meeting.AddAttend;
                entity.FileLink = meeting.FileLink;
                entity.Record = meeting.Record;
                entity.Recfile = meeting.Recfile;
                entity.InitName = meeting.InitName;
                entity.FName = meeting.FName;
                


                foreach (var meetingAttendee in entity.MeetingAttendees.ToList())
                {
                    db.MeetingAttendees.Remove(meetingAttendee);
                }

                if (meeting.Attendees != null)
                {
                    foreach (var attendeeId in meeting.Attendees)
                    {
                        var meetingAttendee = new MeetingAttendee
                        {
                            MeetingID = meeting.MeetingID,
                            AttendeeID = attendeeId
                        };

                        db.MeetingAttendees.Add(meetingAttendee);
                    }
                }

                db.SaveChanges();
            }
        }

        public void Delete(MeetingViewModel meeting, ModelStateDictionary modelState)
        {
            if (meeting.Attendees == null)
            {
                meeting.Attendees = new string[0];
            }

            var entity = meeting.ToEntity();

            db.Meetings.Attach(entity);

            var attendees = meeting.Attendees.Select(attendee => new MeetingAttendee
            {
                AttendeeID = attendee,
                MeetingID = entity.MeetingID
            });

            foreach (var attendee in attendees)
            {
                db.MeetingAttendees.Attach(attendee);
            }

            entity.MeetingAttendees.Clear();

            var recurrenceExceptions = db.Meetings.Where(m => m.RecurrenceID == entity.MeetingID);

            foreach (var recurrenceException in recurrenceExceptions)
            {
                db.Meetings.Remove(recurrenceException);
            }

            db.Meetings.Remove(entity);
            db.SaveChanges();
        }

        private bool ValidateModel(MeetingViewModel appointment, ModelStateDictionary modelState)
        {
            if (appointment.Start > appointment.End)
            {
                modelState.AddModelError("errors", "Конечная дата должна быть позже начальной");
                return false;
            }

            return true;
        }

        public void Dispose()
        {
            db.Dispose();
        }
    }
}