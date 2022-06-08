#!/usr/bin/python

from icalendar import Calendar, Event, vCalAddress, vText
from tnefparse import TNEF
from datetime import datetime, timezone, timedelta as td
import re
import pytz
import sys
from abc import ABC
from html.parser import HTMLParser

# Simple HTML cleanup helper class for the event description
class HTMLFilter(HTMLParser, ABC):
    """
    A simple no dependency HTML -> TEXT converter.
    Based on https://stackoverflow.com/a/62180428
    Usage:
          str_output = HTMLFilter.convert_html_to_text(html_input)
    """
    def __init__(self, *args, **kwargs):
        self.text = ''
        self.in_body = False
        super().__init__(*args, **kwargs)

    def handle_starttag(self, tag: str, attrs):
        if tag.lower() == "body":
            self.in_body = True
        elif tag.lower() in ("br", "p"):
            self.text += "\n" 

    def handle_endtag(self, tag):
        if tag.lower() == "body":
            self.in_body = False

    def handle_startendtag(self, tag):
        if tag.lower() == "br":
            self.text += "\n"

    def handle_data(self, data):
        if self.in_body:
            self.text += data

    @classmethod
    def convert_html_to_text(cls, html: str) -> str:
        f = cls()
        f.feed(html)
        return f.text.strip()  

# Helper function to extract timezone information
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


# Process script arguments

if len(sys.argv) < 2:
    sys.exit("First argument must be the path to a winmail.dat file to convert!")
else:
    infile = str(sys.argv[1])

if len(sys.argv) == 3:
    outfile = str(sys.argv[2])
else:
    outfile = 'invite.ics'

# Init icalendar objects
newcal = Calendar()
newevent = Event()

# Process TNEF input file by looking for some properties
with open(infile, "rb") as tneffile:
    tnefobj = TNEF(tneffile.read())
    
    # The MAPI properties are returned in the form of a list of MAPIAtrributes objects
    # Not sure how we're supposed to look up values in this list.
    # Going through all of them and looking for the ones I'm looking for instead
    
    newcal.add('prodid', '-//TNEF to ICS converter script//https://github.com/timtomch/tnef2ics//')
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
        if value.name_str == 'MAPI_SENDER_EMAIL_ADDRESS':
            orgemail = value.data

    try:
        organizer = vCalAddress('MAILTO:' + orgemail)
        organizer.params['cn'] = orgname
        newevent.add('organizer',organizer)
    except:
        # One of the organizer parameters is missing, no biggie
        pass

    try:
        # Looks like regardless of the time zone description specified, times in the MAPI tags are UTC...
        #tz = parse_tz(tzstring)
        tz = pytz.timezone("UTC")
        newevent.add('dtstart',tz.localize(starttime))
        newevent.add('dtend',tz.localize(endtime))
    except:
        # No date information found, are you sure this is a meeting invite?
        print("No date information found, event not created.",file=sys.stderr)
        sys.exit(1)

# Add the message body as description, after converting the HTML to plain text with newlines.        
newevent.add('description',HTMLFilter.convert_html_to_text(tnefobj.htmlbody))
    
# Add the event we just created to the calendar object
newcal.add_component(newevent)

# Write out the resulting calendar object to the specified output file path
with open(outfile, "wb") as outfile:
    outfile.write(newcal.to_ical())