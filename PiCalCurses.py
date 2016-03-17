import curses
from time import sleep
from PiCal import *
import locale
locale.setlocale(locale.LC_ALL, '')

FREE = 'Free!'

def display_event(win, evt):

    win.clear()

    if evt == FREE:
        win.addstr(1, 0, FREE, curses.A_BOLD)

    else:
        win.addstr(0, 0, evt.description.encode('utf_8'), curses.A_BOLD)
        win.addstr(1, 0, evt.location.encode('utf_8'))
        win.addstr(2, 0, "Start: {0:%H:%M}".format(evt.start_time), curses.A_DIM)
        win.addstr(3, 0, "End: {0:%H:%M}".format(evt.end_time), curses.A_DIM)

def display_events(stdscr):

    height, width = stdscr.getmaxyx()
    cw = CalendarWrapper.from_url('https://outlook.office365.com/owa/calendar/6f91291dec8b49b7b6fb589e1c53c2c4@adobe.com/0a5ffcf1b2db4ab683b8c22384b5e6317191843820375159142/calendar.ics')
    todays_events = sorted(cw.get_todays_events(),key=CalendarEvent.sort_key)

    current_pad = stdscr.subpad(4, width - 10, 3, 4)
    next_pad = stdscr.subpad(4, width - 10, 10, 4)

    if todays_events[0].status == 'OOF' and todays_events[0].all_day:
        current_pad.addstr(0, 0, "I'm out of the office today.")
        current_pad.addstr(2, 0, todays_events[0].description.encode('utf_8'), curses.A_BOLD)
        current_pad.addstr(3, 0, todays_events[0].location.encode('utf_8'))
        next_pad.clear()
        stdscr.refresh()
        return

    events = [x for x in todays_events if x.status == 'BUSY']

    now = datetime.now().time()

    current = FREE
    next = FREE

    for i in xrange(len(events)):

        if events[i].start_time < now < events[i].end_time:

            current = events[i]

            if i < len(events) - 1:
                next = events[i+1]

            continue

        elif events[i].start_time > now:

            next = events[i]
            continue

    display_event(current_pad, current)
    display_event(next_pad, next)

    stdscr.refresh()

def main():

    stdscr = curses.initscr()

    try:
        height, width = stdscr.getmaxyx()
        stdscr.border()
        stdscr.addstr(1, 1, " CURRENTLY".ljust(width-2), curses.A_REVERSE)
        stdscr.addstr(8, 1, " NEXT".ljust(width-2), curses.A_REVERSE)
        stdscr.refresh()

        while True:
            display_events(stdscr)
            sleep(15)

    except KeyboardInterrupt:
        pass

    finally:
        curses.endwin()
if __name__ == '__main__':
    main()
