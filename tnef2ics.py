#!/usr/bin/python

from icalendar import Calendar, Event
from tnefparse import TNEF
from datetime import datetime, timezone, timedelta as td
import re
import pytz
import sys

if len(sys.argv) < 2:
    sys.exit("First argument must be the path to a winmail.dat file to convert!")
else:
    infile = str(sys.argv[1])

if len(sys.argv) == 3:
    outfile = str(sys.argv[2])
else:
    outfile = 'invite.ics'

newcal = Calendar()
newevent = Event()

def parse_tz(tzstr):
    p = re.compile('UTC([+-])(\d\d):(\d\d)')
    m = p.search(tzstr)
    if m:
        sign = m.group(1)
        try:
            hs = m.group(2).lstrip('0')
            ms = m.group(3).lstrip('0')
        except:
            return None

        tz_offset = td(hours=int(hs) if hs else 0,
                       minutes=int(ms) if ms else 0)

        return timezone(-tz_offset if sign == '-' else tz_offset)


with open(infile, "rb") as tneffile:
    tnefobj = TNEF(tneffile.read())
    
    # The MAPI properties are returned in the form of a list of MAPIAtrributes objects
    # Not sure how we're supposed to look up values in this list.
    # Going through all of them and looking for the ones I'm looking for instead
    
    newcal.add('prodid', '-//TNEF to ICS converter//labs.timtom.ch//')
    newcal.add('version', '2.0')
    
    for value in tnefobj.mapiprops:
        if value.name_str == 'MAPI_SUBJECT':
            newevent.add('summary', value.data)
        if value.name_str == 'MAPI_OUTLOOK_LOCATION':
            newevent.add('location', value.data)
        if value.name_str == 'MAPI_TIME_ZONE_DESCRIPTION':
            tzstring = value.data
        if value.name_str == 'MAPI_START_DATE':
            starttime = value.data
        if value.name_str == 'MAPI_END_DATE':
            endtime = value.data
        if value.name_str == 'MAPI_CREATOR_NAME':
            orgname = value.data
        if value.name_str == 'MAPI_SENDER_EMAIL_ADDRESS"':
            orgemail = value.data
    
    #starttime_object = datetime.strptime(starttime, '%Y-%m-%d %H:%M:%s')
    
    
    if timezone != '':
        tz = parse_tz(tzstring)
        starttime = starttime.replace(tzinfo=tz)
        endtime = endtime.replace(tzinfo=tz)
        #print(starttime.strftime('%Y%m%dT%H%M%SZ'))
    newevent.add('dtstart',starttime)
    newevent.add('dtend',endtime)
    
    newevent.add('description',tnefobj.body)
    
newcal.add_component(newevent)



with open(outfile, "wb") as outfile:
    outfile.write(newcal.to_ical())