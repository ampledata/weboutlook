"""
Microsoft Outlook Web Access scraper

Retrieves full, raw e-mails from Microsoft Outlook Web Access by
screen scraping. Can do the following:

    * Log into a Microsoft Outlook Web Access account with a given username
      and password.
    * Retrieve all e-mail IDs from the first page of your Inbox.
    * Retrieve the full, raw source of the e-mail with a given ID.
    * Delete an e-mail with a given ID (technically, move it to the "Deleted
      Items" folder).

The main class you use is OutlookWebScraper. See the docstrings in the code
and the "sample usage" section below.

This module does no caching. Each time you retrieve something, it does a fresh
HTTP request. It does cache your session, though, so that you only have to log
in once.

Updated by Greg Albrecht <gba@gregalbrecht.com>
Based on http://code.google.com/p/weboutlook/ by Adrian Holovaty <holovaty@gmail.com>.
"""

# Documentation / sample usage:
#
# # Throws InvalidLogin exception for invalid username/password.
# >>> s = OutlookWebScraper('https://webmaildomain.com', 'username', 'invalid password')
# >>> s.login()
# Traceback (most recent call last):
#     ...
# scraper.InvalidLogin
#
# >>> s = OutlookWebScraper('https://webmaildomain.com', 'username', 'correct password')
# >>> s.login()
#
# # Display IDs of messages in the inbox.
# >>> s.inbox()
# ['/Inbox/Hey%20there.EML', '/Inbox/test-3.EML']
#
# # Display IDs of messages in the "sent items" folder.
# >>> s.get_folder('sent items')
# ['/Sent%20Items/test-2.EML']
#
# # Display the raw source of a particular message.
# >>> print s.get_message('/Inbox/Hey%20there.EML')
# [...]
#
# # Delete a message.
# >>> s.delete_message('/Inbox/Hey%20there.EML')

# Copyright (C) 2006 Adrian Holovaty <holovaty@gmail.com>
#
# This program is free software; you can redistribute it and/or modify it under
# the terms of the GNU General Public License as published by the Free Software
# Foundation; either version 2 of the License, or (at your option) any later
# version.
#
# This program is distributed in the hope that it will be useful, but WITHOUT
# ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS
# FOR A PARTICULAR PURPOSE. See the GNU General Public License for more
# details.
#
# You should have received a copy of the GNU General Public License along with
# this program; if not, write to the Free Software Foundation, Inc., 59 Temple
# Place, Suite 330, Boston, MA 02111-1307 USA

import re, socket, urllib, urlparse, urllib2
from Cookie import SimpleCookie
import logging
from logging.handlers import *

__version__ = '0.1.1'
__author__ = 'Greg Albrecht <gba@gregalbrecht.com>'

logger = logging.getLogger('weboutlook')
logger.setLevel(logging.INFO)
consolelogger = logging.StreamHandler()
consolelogger.setLevel(logging.INFO)
logger.addHandler(consolelogger)

socket.setdefaulttimeout(15)

class InvalidLogin(Exception):
    pass

class RetrievalError(Exception):
    pass

def create_opener(base_url, username, password):
    logger.debug(locals())
    password_mgr = urllib2.HTTPPasswordMgrWithDefaultRealm()
    password_mgr.add_password(None, base_url, username, password)
    handler = urllib2.HTTPBasicAuthHandler(password_mgr)
    return urllib2.build_opener(handler)

def top_url(url):
    logger.debug(locals())
    p = urlparse.urlparse(url)
    return '%s://%s/' % (p.scheme, p.netloc)

class CookieScraper(object):
    "Scraper that keeps track of getting and setting cookies."
    def __init__(self):
        logger.debug(locals())
        self._cookies = SimpleCookie()
    
    
    def get_page(self, url, post_data=None, headers=()):
        """
        Helper method that gets the given URL, handling the sending and storing
        of cookies. Returns the requested page as a string.
        """
        logger.debug(locals())
        opener = create_opener(top_url(url), self.username, self.password)
        request = urllib2.Request(url)
        request.add_header('Cookie', self._cookies.output(attrs=[], header='').strip())
        for k, v in headers:
            request.add_header(k, v)
        try:
            f = opener.open(request)
        except IOError, e:
            if isinstance(e, urllib2.HTTPError):
                code = e.code
            else:
                code = e[1]
            if code == 302:
                # Got a 302 redirect, but check for cookies before redirecting.
                # e[3] is a httplib.HTTPMessage instance.
                if e[3].dict.has_key('set-cookie'):
                    self._cookies.load(e[3].dict['set-cookie'])
                return self.get_page(e[3].getheader('location'))
            elif code == 401:
                raise InvalidLogin
            else:
                raise
        if f.headers.dict.has_key('set-cookie'):
            self._cookies.load(f.headers.dict['set-cookie'])
        return f.read()
    

class OutlookWebScraper(CookieScraper):
    def __init__(self, domain, username, password):
        logger.debug(locals())
        self.domain = domain
        self.username, self.password = username, password
        self.is_logged_in = False
        self.base_href = None
        super(OutlookWebScraper, self).__init__()
    
    
    def login(self):
        logger.debug(locals())
        destination = urlparse.urljoin(self.domain, 'exchange/')
        url         = destination
        html        = self.get_page(url=url)
        if 'You could not be logged on to Outlook Web Access' in html:
            raise InvalidLogin
        m = re.search(r'(?i)<BASE href="([^"]*)">', html)
        if not m:
            raise RetrievalError, "Couldn't find <base href> on page after logging in."
        self.base_href      = m.group(1)
        self.is_logged_in   = True
    
    
    def inbox(self):
        """
        Returns the message IDs for all messages on the first page of the
        Inbox, regardless of whether they've already been read.
        """
        logger.debug(locals())
        return self.get_folder('Inbox')
    
    
    def get_folder(self, folder_name):
        """
        Returns the message IDs for all messages on the first page of the
        folder with the given name, regardless of whether the messages have
        already been read. The folder name is case insensitive.
        """
        logger.debug(locals())
        if not self.is_logged_in: self.login()
        url = self.base_href + urllib.quote(folder_name) + '/?Cmd=contents'
        html = self.get_page(url)
        message_urls = re.findall(r'(?i)NAME=MsgID value="([^"]*)"', html)
        return message_urls
    
    
    def get_message(self, msgid):
        "Returns the raw e-mail for the given message ID."
        logger.debug(locals())
        if not self.is_logged_in: self.login()
        # Sending the "Translate=f" HTTP header tells Outlook to include
        # full e-mail headers. Figuring that out took way too long.
        return self.get_page(self.base_href + msgid + '?Cmd=body', headers=[('Translate', 'f')])
    
    
    def delete_message(self, msgid):
        "Deletes the e-mail with the given message ID."
        logger.debug(locals())
        if not self.is_logged_in: self.login()
        return self.get_page(self.base_href + msgid, urllib.urlencode({
            'MsgId': msgid,
            'Cmd': 'delete',
            'ReadForm': '1',
        }))
    
