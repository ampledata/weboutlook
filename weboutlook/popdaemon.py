"""
POP e-mail server for Microsoft Outlook Web Access scraper

This wraps the Outlook Web Access scraper by providing a POP interface to it.

Run this file from the command line to start the server on localhost port 8110.
The server will run until you end the process. If you specify the "--once"
command-line option, it'll only run one full POP transaction and quit. The
latter is useful for the "precommand" option in KMail, which you can configure
to run this script every time you check e-mail so it starts the server in the
background.

Note that you'll have to specify WEBMAIL_SERVER in this file.
"""

# Based on gmailpopd.py by follower@myrealbox.com,
# which was in turn based on smtpd.py by Barry Warsaw.
#
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

import logging
from logging.handlers import *

import asyncore, asynchat, socket, sys

# relative import
from scraper import InvalidLogin, OutlookWebScraper
from weboutlook_conf import *

__version__ = 'Python Outlook Web Access POP3 proxy version 0.0.1.2'

TERMINATOR = '\r\n'

logger = logging.getLogger('weboutlook')
logger.setLevel(logging.INFO)
consolelogger = logging.StreamHandler()
consolelogger.setLevel(logging.INFO)
logger.addHandler(consolelogger)

def quote_dots(lines):
    for line in lines:
        if line.startswith("."):
            line = "." + line
        yield line

class POPChannel(asynchat.async_chat):
    def __init__(self, conn, quit_after_one):
        logger.debug(locals())
        self.quit_after_one = quit_after_one
        asynchat.async_chat.__init__(self, conn)
        self.__line = []
        self.push('+OK %s %s' % (socket.getfqdn(), __version__))
        self.set_terminator(TERMINATOR)
        self._activeDataChannel = None

    # Overrides base class for convenience
    def push(self, msg):
        logger.debug(locals())
        asynchat.async_chat.push(self, msg + TERMINATOR)

    # Implementation of base class abstract method
    def collect_incoming_data(self, data):
        logger.debug(locals())
        self.__line.append(data)

    # Implementation of base class abstract method
    def found_terminator(self):
        logger.debug(locals())
        line = ''.join(self.__line)
        self.__line = []
        if not line:
            self.push('500 Error: bad syntax')
            return
        method = None
        i = line.find(' ')
        if i < 0:
            command = line.upper()
            arg = None
        else:
            command = line[:i].upper()
            arg = line[i+1:].strip()
        method = getattr(self, 'pop_' + command, None)
        if not method:
            self.push('-ERR Error : command "%s" not implemented' % command)
            return
        method(arg)
        return

    def pop_UIDL(self, which=None):
        logger.debug(locals())
        """Return message digest (unique id) list.

        If 'which', result contains unique id for that message
        in the form 'response mesgnum uid', otherwise result is
        the list ['response', ['mesgnum uid', ...], octets]
        """
        return self.pop_LIST(arg=which)


    def pop_USER(self, user):
        logger.debug(locals())
        # Logs in any username.
        if not user:
            self.push('-ERR: Syntax: USER username')
        else:
            self.username = ''.join((USER_PREFIX,user)) # Store for later.
            logger.info("username=%s" % self.username)
            self.push('+OK Password required')

    def pop_PASS(self, password=''):
        logger.debug(locals())
        self.scraper = OutlookWebScraper(WEBMAIL_SERVER, self.username, password)
        try:
            self.scraper.login()
        except InvalidLogin:
            self.push('-ERR Login failed. (Wrong username/password?)')
        else:
            self.push('+OK User logged in')
            self.inbox_cache = self.scraper.inbox()
            self.msg_cache = [self.scraper.get_message(msg_id) for msg_id in self.inbox_cache]

    def pop_STAT(self, arg):
        logger.debug(locals())
        dropbox_size = sum([len(msg) for msg in self.msg_cache])
        self.push('+OK %d %d' % (len(self.inbox_cache), dropbox_size))

    def pop_LIST(self, arg):
        logger.debug(locals())
        if not arg:
            num_messages = len(self.inbox_cache)
            self.push('+OK')
            for i, msg in enumerate(self.msg_cache):
                self.push('%d %d' % (i+1, len(msg)))
            self.push(".")
        else:
            # TODO: Handle per-msg LIST commands
            raise NotImplementedError

    def pop_RETR(self, arg):
        logger.debug(locals())
        if not arg:
            self.push('-ERR: Syntax: RETR msg')
        else:
            # TODO: Check request is in range.
            msg_index = int(arg) - 1
            msg = self.msg_cache[msg_index]
            msg_id = self.inbox_cache[msg_index]
            msg = msg.lstrip() + TERMINATOR

            self.push('+OK')

            for line in quote_dots(msg.split(TERMINATOR)):
                self.push(line)
            self.push('.')

            # Delete the message
            self.scraper.delete_message(msg_id)

    def pop_QUIT(self, arg):
        logger.debug(locals())
        self.push('+OK Goodbye')
        self.close_when_done()
        if self.quit_after_one:
            # This SystemExit gets propogated to handle_error(),
            # which stops the program. Slightly hackish.
            raise SystemExit

    def handle_error(self):
        logger.debug(locals())
        if self.quit_after_one:
            sys.exit(0) # Exit.
        else:
            asynchat.async_chat.handle_error(self)

class POP3Proxy(asyncore.dispatcher):
    def __init__(self, localaddr, quit_after_one):
        """
        localaddr is a tuple of (ip_address, port).

        quit_after_one is a boolean specifying whether the server should quit
        after serving one session.
        """
        logger.debug(locals())
        self.quit_after_one = quit_after_one
        asyncore.dispatcher.__init__(self)
        self.create_socket(socket.AF_INET, socket.SOCK_STREAM)
        # try to re-use a server port if possible
        self.set_reuse_addr()
        self.bind(localaddr)
        self.listen(5)

    def handle_accept(self):
        logger.debug(locals())
        conn, addr = self.accept()
        channel = POPChannel(conn, self.quit_after_one)

if __name__ == '__main__':
    from optparse import OptionParser
    parser = OptionParser("usage: %prog [options]")
    parser.add_option('--once', action='store_true', dest='once',
        help='Serve one POP transaction and then quit. (Server runs forever by default.)')
    options, args = parser.parse_args()
    proxy = POP3Proxy(('127.0.0.1', 8110), options.once is True)
    try:
        asyncore.loop()
    except KeyboardInterrupt:
        pass
