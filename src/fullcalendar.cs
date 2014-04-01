#region License
// Copyright (c) 2013, Joel Serbinski
// All rights reserved.

// Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
// Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
// Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

// THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, 
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS
// BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE 
// GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, 
// STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#endregion

using System;
using System.Data;
using Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json;


namespace jQueryPlugin_wrappers_csharp
{

/* --[ FullCalendar ]--
 * DESC: wrapper class for operating the FullCalendar plugin
 * 
 * NOTES: classes prefixed with "fc_" mirror FullCalendar (our implementation)
 * 
 *  
 */

public class FullCalendar
{
    public fc_event[] events;


	public FullCalendar()
	{
        events = new fc_event[0];
	}
    // Mirrors the "event object"
    public class fc_event
    {
        public string id="";
        public string title="";
        public bool allDay = false;
        public DateTime start = DateTime.Now;
        public DateTime end = DateTime.Now;
        public string url;
        public string className;
        public bool editable;
        public string tooltip;

        public string color;
        public string textColor;
        public string borderColor;
    }
    // the desired settings for outputting the dynamic javascript needed
    public class js
    {
        public string id; // REQUIRED: the id of the object
        public string events = ""; // REQUIRED: the ajax call for json cal data
        public string header = "{left:'title', center:'agendaWeek,month', right:'today prev,next'}";
        public bool editable = true;
        public bool allDaySlot = true;
        public string defaultView = "'month'";  // agendaWeek
        public int defaultEventMinutes = 60; // this is for events with "no end date", which are ALL DAY entries
        public string aspectRatio = "1.5";
        public string height = "''"; // default is to NOT set the height
        // This is event that runs when it draws the element, so attach our tooltip (qtip) plugin here
        public string eventRender = "function (event, element) {" +
                    "element.qtip({" +
                        "content: event.tooltip," + // This is our tooltip content...
                        "position: {" +
                            "my: 'top left'," +
                            "target: 'mouse'," +
                            "viewport: $(window), " + // Keep it on-screen at all times if possible
                            "adjust: {x: 10, y: 10}" +
                        "}," +
                        "hide: {fixed: true }," + // Helps to prevent the tooltip from hiding ocassionally when tracking!
                        "style: 'ui-tooltip-shadow'" +
                    "});" +                    
        "}";
        // The following need to be set to handle event changing...defaults here for examples
        public string eventDrop = "function(){return false;}";
            /* example:
            "function(event,dayDelta,minuteDelta,allDay,revertFunc){" +
            "alert('The end date of ' + event.title + 'has been moved ' + dayDelta + ' days and ' + minuteDelta + ' minutes.');" +
            "if (!confirm('is this okay?')){revertFunc();}" +
            "}"; */
        public string eventResize = "function(){return false;}";
            /* example:
            "function(event,dayDelta,minuteDelta,revertFunc) {"+
            "alert('The end date of ' + event.title + 'has been moved ' + dayDelta + ' days and ' + minuteDelta + ' minutes.');"+
            "if (!confirm('is this okay?')){revertFunc();}"+
            "}";
             */

    }

    // ------------------------------- FUNCTIONS
    public string RenderJSON()
    {        
        // Convert with ISO8601 Date/Time
        return JsonConvert.SerializeObject(this.events, new Newtonsoft.Json.Converters.IsoDateTimeConverter());
    }
    // fc_js :: bulids the javascript output for creating the FullCalendar object
    public string fc_js(js vars)
    {
        string output = "$('#" + vars.id + "').fullCalendar({" +
            "events:" + vars.events + ","+
            "header:" + vars.header + "," +
            "editable:" + vars.editable.ToString().ToLower() + "," +
            "allDaySlot:" + vars.allDaySlot.ToString().ToLower() + "," +
            "defaultView:" + vars.defaultView + "," +
            "defaultEventMinutes:" + vars.defaultEventMinutes.ToString() + "," +
            "aspectRatio:" + vars.aspectRatio + "," +
            "height:" + vars.height + "," +

            "eventRender:" + vars.eventRender+","+
            "eventDrop:" + vars.eventDrop + "," +
            "eventResize:"+vars.eventResize

        + "});"; 
        
        return output;
    }

    public int findEvent(string id)
    {
        for (int t = 0; t < this.events.Length; t++)
        {
            if (this.events[t].id == id) { return t; }
        }
        return -1;
    }

    // --[ addEvent ]--
    // DESC: Adds a generic event to the array
    public int addEvent(string id, string title, DateTime Start,DateTime End,bool allDay=false
        , bool editable=false, string tooltip="", int color=255)
    {
        Array.Resize(ref events, events.Length + 1);
        this.events[this.events.Length - 1] = new fc_event();
        this.events[this.events.Length - 1].id = id;
        this.events[this.events.Length - 1].title = title;
        this.events[this.events.Length - 1].start = Start;
        this.events[this.events.Length - 1].end = End;
        this.events[this.events.Length - 1].allDay = allDay;
        this.events[this.events.Length - 1].editable = editable;
        if (tooltip != "") {
            this.events[this.events.Length - 1].tooltip = tooltip;
        }
        else {
            this.events[this.events.Length - 1].tooltip = "Title: "+title;
        }
        this.events[this.events.Length - 1].color = color_HTMLHEX(color);
        return this.events.Length - 1;
    }
    // --[ delEvent ]--
    // DESC: Removes an event from the array
    public void delEvent(string id)
    {
        fc_event[] newEvents = new fc_event[this.events.Length-1];
        if (this.events != null)
        {
            //. if only one
            if (this.events.Length == 1) { this.events = new fc_event[0]; return; }
            //. rebuild array without undesired event
            int t,a;
            // we can leave the array the same until we reach the control to be removed
            for (t = 0; t < this.events.Length; t++) {
                if (this.events[t].id == id){break;}
            }
            // check if the removed control is at the end of array..otherwise, shift remaining array
            if (t == this.events.Length - 1) { 
                Array.Resize(ref this.events, t);
            }
            else {
                for (a=t;a<this.events.Length-1;a++){
                this.events[a] = this.events[a+1];
                }
            }
        }
    }
    // --[ moveEvent ]--
    // DESC: Moves both the start and end dates by Deltas
    public void moveEvent(string id, double dayDelta, double minuteDelta)
    {
        int iEvent = this.findEvent(id);
        this.events[iEvent].start = this.events[iEvent].start.AddDays(dayDelta);
        this.events[iEvent].start = this.events[iEvent].start.AddMinutes(minuteDelta);
        this.events[iEvent].end = this.events[iEvent].end.AddDays(dayDelta);
        this.events[iEvent].end = this.events[iEvent].end.AddMinutes(minuteDelta);
    }
    // --[ resizeEvent ]--
    // DESC: Resize end date by Delta
    public void resizeEvent(string id, double minuteDelta)
    {
        int iEvent = this.findEvent(id);
        //this.events[iEvent].end = this.events[iEvent].end.AddDays(dayDelta);
        // Check to make sure not trying to go past starting time...
        if (this.events[iEvent].end.AddMinutes(minuteDelta) >= this.events[iEvent].start)
        {
            this.events[iEvent].end = this.events[iEvent].end.AddMinutes(minuteDelta);
        }
        else
        { // just make same as start
            this.events[iEvent].end = this.events[iEvent].start;
        }
    }


    // --------------------------------------------------------------------------------------------
    // REF: http://technet.microsoft.com/en-us/library/ee692908.aspx
    public void add_OutlookCalendar(DateTime Start, DateTime End)
    {
        Microsoft.Office.Interop.Outlook.Application OutlookApp = null;
        Microsoft.Office.Interop.Outlook.NameSpace mapiNamespace = null;
        Microsoft.Office.Interop.Outlook.MAPIFolder CalendarFolder = null;
        Microsoft.Office.Interop.Outlook.Items outlookCalendarItems = null;

        OutlookApp = new Microsoft.Office.Interop.Outlook.Application();
        mapiNamespace = OutlookApp.GetNamespace("MAPI"); ;
        CalendarFolder = mapiNamespace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderCalendar);
        string filterDateRange = "[Start] >= '" + Start.ToString("g")+ "' AND [End] <= '" + End.ToString("g") + "'";
        outlookCalendarItems = CalendarFolder.Items.Restrict(filterDateRange); // Only get the appts. by dates needed
        outlookCalendarItems.IncludeRecurrences = true;
        foreach (Microsoft.Office.Interop.Outlook.AppointmentItem item in outlookCalendarItems)
        {

            if (item.IsRecurring)
            {
                Microsoft.Office.Interop.Outlook.RecurrencePattern rp = item.GetRecurrencePattern();
                /*
                 *    RecurrenceType    Properties      Example
                 *    ______________________________________________________________________________
                      olRecursDaily     Interval        Every N days
                                        DayOfWeekMask   Every Tuesday, Wednesday, and Thursday

                      olRecursMonthly   Interval        Every N months
                                        DayOfMonth      The Nth day of the month

                      olRecursMonthNth  Interval        Every N months
                                        Instance        The Nth Tuesday
                                        DayOfWeekMask   Every Tuesday and Wednesday

                      olRecursWeekly    Interval        Every N weeks
                                        DayOfWeekMask   Every Tuesday, Wednesday, and Thursday

                      olRecursYearly    DayOfMonth      The Nth day of the month
                                        MonthOfYear     February

                      olRecursYearNth   Instance        The Nth Tuesday
                                        DayOfWeekMask   Tuesday, Wednesday, Thursday
                                        MonthOfYear     February
                 * 
                 * VALID properties for each type from MSDN
                    olRecursDaily 	
                    Duration , EndTime, Interval, NoEndDate, Occurrences, PatternStartDate, PatternEndDate, StartTime

                    olRecursWeekly 	
                    DayOfWeekMask , Duration, EndTime, Interval, NoEndDate, Occurrences, PatternStartDate, PatternEndDate, StartTime

                    olRecursMonthly 	
                    DayOfMonth , Duration, EndTime, Interval, NoEndDate, Occurrences, PatternStartDate, PatternEndDate, StartTime

                    olRecursMonthNth 	
                    DayOfWeekMask , Duration, EndTime, Interval, Instance, NoEndDate, Occurrences, PatternStartDate, PatternEndDate, StartTime

                    olRecursYearly 	
                    DayOfMonth , Duration, EndTime, Interval, MonthOfYear, NoEndDate, Occurrences, PatternStartDate, PatternEndDate, StartTime

                    olRecursYearNth 	
                    DayOfWeekMask , Duration, EndTime, Interval, Instance, NoEndDate, Occurrences, PatternStartDate, PatternEndDate, StartTime
                */
                AppointmentItem reoccuringAppt;
                for (DateTime dateChk = Start; dateChk <= End; dateChk=dateChk.AddDays(1))
                {
                    try
                    { // just see if one exists for this date, if not continue!
                        reoccuringAppt = rp.GetOccurrence(DateTime.Parse(dateChk.ToShortDateString() + " " + item.Start.ToShortTimeString()));
                        add_OutlookCalendar_item(reoccuringAppt);
                    }
                    catch {
                        continue;
                    }
                    //// For each date requested, we want to know if a 
                    //switch (rp.RecurrenceType)
                    //{
                    //    case Microsoft.Office.Interop.Outlook.OlRecurrenceType.olRecursDaily:
                    //    case Microsoft.Office.Interop.Outlook.OlRecurrenceType.olRecursWeekly:
                    //        // Basically, we just decided that we can check to see if an appt. exists for any of the days of the week
                    //        // and accomplish adding a week to the schedule
                    //        if ((rp.DayOfWeekMask | OlDaysOfWeek.olMonday) == OlDaysOfWeek.olMonday)
                    //        {
                    //        }
                    //        switch (rp.DayOfWeekMask)
                    //        {
                    //            case Microsoft.Office.Interop.Outlook.OlDaysOfWeek.olMonday:

                    //                break;
                    //            case Microsoft.Office.Interop.Outlook.OlDaysOfWeek.olTuesday:
                    //                break;
                    //        }
                    //        //dateChk.AddDays(
                    //        //nextoccurencestartdate = dateadd("d", occurencestardate,1)
                    //        break;
                    //    case Microsoft.Office.Interop.Outlook.OlRecurrenceType.olRecursMonthly:
                    //    case Microsoft.Office.Interop.Outlook.OlRecurrenceType.olRecursMonthNth:
                    //    case Microsoft.Office.Interop.Outlook.OlRecurrenceType.olRecursYearly:
                    //    case Microsoft.Office.Interop.Outlook.OlRecurrenceType.olRecursYearNth:
                    //        // for all these, either an appt. exists for the 
                    //        break;
                    //}
                }
                //DateTime first = DateTime.Parse(Start.ToShortDateString+ new DateTime(2008, 8, 31, item.Start.Hour, item.Start.Minute, 0);
                //DateTime last = new DateTime(2008, 10, 1);
                //Microsoft.Office.Interop.Outlook.AppointmentItem recur = null;
                //for (DateTime cur = first; cur <= last; cur = cur.AddDays(1))
                //{
                //    try
                //    {
                //        recur = rp.GetOccurrence(cur);
                //        MessageBox.Show(recur.Subject + " -> " + cur.ToLongDateString());
                //    }
                //    catch
                //    { }
                //}
            }
            else
            { // One-time Appts.
                add_OutlookCalendar_item(item);
            }
        }

    }

    private void add_OutlookCalendar_item(AppointmentItem item) {
                Array.Resize(ref this.events, this.events.Length + 1);
                this.events[this.events.Length - 1] = new fc_event();
                this.events[this.events.Length - 1].id = item.EntryID; // item.GlobalAppointmentID;
                this.events[this.events.Length - 1].title = item.Subject;
                this.events[this.events.Length - 1].start = item.Start;
                this.events[this.events.Length - 1].end = item.End;
                this.events[this.events.Length - 1].editable = true;
    }

    private string color_HTMLHEX(int color) {        
        return string.Format("#{0:X2}{1:X2}{2:X2}",
                (byte)((color & 0xff0000) >> 0x10),
                (byte)((color & 0xff00) >> 8),
                (byte)(color & 0xff)
                );

    }

} // end FullCalendar
}