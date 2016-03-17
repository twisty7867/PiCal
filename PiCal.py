from icalendar import Calendar
from datetime import datetime, date, time
import requests

def separate_datetime(dt, end=False):
    if type(dt) is date:
        return dt, time.max if end else time.min
    else:
        return dt.date(), dt.time()

class CalendarEvent(object):

    def __init__(self, description='', status='', start_time=time.min, end_time=time.max, location='', all_day=False):
        self.description = description
        self.status = status
        self.start_time = start_time
        self.end_time = end_time
        self.location = location
        self.all_day = all_day

    @staticmethod
    def sort_key(evt):
        key = (evt.start_time.hour * 60) + evt.start_time.minute
        if evt.status == "BUSY":
            key += 0.1
        elif evt.status == "TENTATIVE":
            key += 0.2
        elif evt.status == "FREE":
            key += 0.3
        return key

    def location_str(self):

        return "" if self.location is None else " ({0})".format(self.location)

    def status_str(self):

        return "" if self.status == "BUSY" else "({0}) ".format(self.status)

    def description_str(self):

        return self.description[4:] if self.description.startswith("FW: ") else self.description

    def __str__(self):

        if self.all_day:
            return "[all-day]  {0}{1}{2}".format(
                self.status_str(),
                self.description_str(),
                self.location_str())
        else:
            return "[{0:%H:%M}-{1:%H:%M}] {2}{3}{4}".format(
                self.start_time,
                self.end_time,
                self.status_str(),
                self.description_str(),
                self.location_str())


    @staticmethod
    def from_component(component):

        start_date, start_time = separate_datetime(component.get('dtstart').dt)
        end_date, end_time = separate_datetime(component.get('dtend').dt, True)

        return CalendarEvent(
            component.get('summary'),
            component.get('X-MICROSOFT-CDO-BUSYSTATUS'),
            start_time,
            end_time,
            component.get('location'),
            component.get('X-MICROSOFT-CDO-ALLDAYEVENT') == "TRUE")


class CalendarWrapper(object):

    def __init__(self, cal=None):
        self.cal = cal

    @staticmethod
    def from_url(url):

        ical_content = requests.get(url).text

        return CalendarWrapper.from_string(ical_content)

    @staticmethod
    def from_string(ical_content):

        return CalendarWrapper(Calendar.from_ical(ical_content))

    def get_todays_events(self):

        now = datetime.now()

        for component in self.cal.walk():
            if component.name == "VEVENT":

                start_date, start_time = separate_datetime(component.get('dtstart').dt)

                # TODO: fix this to handle multi-day events
                occurs_today = start_date == now.date()

                if occurs_today:
                    yield CalendarEvent.from_component(component)

def main():
    cw = CalendarWrapper.from_url('https://outlook.office365.com/owa/calendar/6f91291dec8b49b7b6fb589e1c53c2c4@adobe.com/0a5ffcf1b2db4ab683b8c22384b5e6317191843820375159142/calendar.ics')

    todays_events = sorted(cw.get_todays_events(),key=CalendarEvent.sort_key)

    if todays_events[0].status == 'OOF' and todays_events[0].all_day:
        print "I'm out of the office today."
        print todays_events[0].description
        print todays_events[0].location
        return

    events = [x for x in todays_events if x.status == 'BUSY']

    now = datetime.now().time()

    print len(events)

    current = "free"
    next = "free"

    for i in xrange(len(events)):

        if events[i].start_time < now < events[i].end_time:

            current = events[i]

            if i < len(events) - 1:
                next = events[i+1]

            continue

        elif events[i].start_time > now:

            next = events[i]
            continue

    print current
    print next

if __name__ == '__main__':
    main()
