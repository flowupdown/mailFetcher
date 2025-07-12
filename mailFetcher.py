""" 
################################################################################
retrieve, delete, match mail from a IMAP server (see __init__ for docs, text)
################################################################################
"""

import imaplib, sys 
from PP4E.Internet.Email.PyMailGui import mailconfig
#import mailconfig                                  # client's mailconfig
print('user:', mailconfig.imapusername)            # script dir, pythonpath, changes

from mailParser import MailParser                  # for headers matching (4E: .)
from mailTool   import MailTool, SilentMailTool    # trace control supers (4E: .)

# index/server msgnum out of synch tests
class DeleteSynchError(Exception): pass             # msg out of synch in del
class TopNotSupported(Exception): pass              # can't run synch test
class MessageSynchError(Exception): pass            # index list out of sync

class MailFetcher(MailTool):
    """ 
    fetch mail: connect, fetch headers+mails, delete mails
    works on any machine with Python+Inet; subclass me to cache
    imlemented with the IMAP protocol;
    4E: handles decoding of full mail text on fetch for parser;
    """
    def __init__(self, imapserver=None, imapuser=None, imappswd=None, hastop=True):
        self.imapServer     = imapserver or mailconfig.imapservername
        self.imapUser       = imapuser   or mailconfig.imapusername
        self.srvrHasTop     = hastop
        self.imapPassword   = imappswd      # ask later if None
        
    def connect(self):
        self.trace('Connecting...')
        self.getPassword()                              # file, GUI, or console
        server = imaplib.IMAP4_SSL(self.imapServer)
        server.login(self.imapUser, self.imapPassword)  # login to IMAP server
        return server
    
    # use setting in client's mailconfig on import search path;
    # to tailor, this can be changed in class or per instance;
    fetchEncoding = mailconfig.fetchEncoding
    
    def decodeFullText(self, messageBytes):
        """ 
        4E, Py3.1: decode full fetched mail text bytes to str Unicode string;
        done at fetch, for later display or parsing (full mail text is always
        Unicode thereafter); decode with per-class or per-instance setting, or
        common types; could also try headers inspection, or intelligent guess
        from structure; in Python 3.2/3.3, this step may not be required: if so,
        change to return message line list intact; for more details see Chapter 13;
        
        
        an 8-bit encoding such as latin-1 will likely suffice for most emails, as
        ASCII is the original standard; this method applies to entire/full message
        text, which is really fust one part of the email encoding story: Message
        payloads and Message headers may also be encoded per email, MIME, and 
        Unicode standards; see Chapter 13 and mailParser and mailSender for more;
        """
        text = None
        kinds =  [self.fetchEncoding]           # try user setting first
        kinds += ['ascii', 'latin1', 'utf8']    # then try common types
        kinds += [sys.getdefaultencoding()]     # and platform dflt (may differ)
        for kind in kinds:                      # may cause mail saves to fail
            try:
                text = [line.decode(kind) for line in messageBytes]
                break
            except (UnicodeError, LookupError): # LookupError: bad name
                pass
            
        if text == None:
            # try returning headers + error msg, else except may kill client;
            # still try to decode headers per ascii, other, platform default;
            
            blankline = messageBytes.index(b'')
            hdrsonly  = messageBytes[:blankline]
            commons   = ['ascii', 'latin1', 'utf8']
            for common in commons:
                try:
                    text = [line.decode(common) for line in hdrsonly]
                    break
                except UnicodeError:
                    pass
            else:                                                   # none worked
                try:
                    text = [line.decode() for line in hdrsonly]     # platform dflt?
                except UnicodeError:
                    text = ['From: (sender of unknown Unicode format headers)']
            text += ['', '--Sorry: mailtools cannot decode this mail content!--']
        return text
    
    def downloadMessage(self, msgnum):
        """ 
        load full raw text of one mail msg, given its
        IMAP relative msgnum; caller must parse content
        """
        self.trace('load ' + str(msgnum))
        server = self.connect()
        try:
            server.select()
            typ, msglines = server.fetch(bytes(str(msgnum), 'utf8'), '(RFC822)')
        finally:
            server.close()
            server.logout()
        msglines = self.decodeFullText(msglines[0])    # raw bytes to Unicode str
        return '\n'.join(msglines)                  # concat lines for parsing
    
    def downloadAllHeaders(self, progress=None, loadfrom=1):
        """ 
        get sizes, raw header text only, for all or new msgs
        begins loading headers from message number loadfrom
        use loadfrom to load newly arrived mails only
        use downloadMessage to get a full msg text later
        progress is a function called with (count, total);
        returns: [headers text], [mail sizes], loadedfull?
        
        4E: add mailconfig.fetchlimit to support large email
        inboxes: if not None, only fetches that many headers,
        and returns others as dummy/empty mail; else inboxes
        like one of mine (4k emails) are not practical to use;
        4E: pass loadfrom along to downloadAllMsgs (a buglet);
        """
        if not self.srvrHasTop:                 # not all servers support TOP
            # naively load full msg text
            return self.downloadAllMessages(progress, loadfrom)
        else:
            self.trace('loading headers')
            fetchlimit = mailconfig.fetchlimit
            server = self.connect()             # mbox now locked until quit
            try:
                typ, data = server.select()
                msgCount = int(data[0].decode())
                allsizes  = []
                allhdrs   = []
                to_load = '(%s:%s)' % (loadfrom, msgCount)
                counter = [i for i in range(loadfrom, msgCount+1)]
                typ, data = server.search(None, to_load)
                for msgnum, msg_counter in zip(data[0].split(), counter):
                    if progress: progress(msg_counter, msgCount)
                    if fetchlimit and (msg_counter <= msgCount - fetchlimit):
                        # skip, add dummy hdrs
                        hdrtext = 'Subject: --mail skipped--\n\n'
                        allsizes.append('')
                        allhdrs.append(hdrtext)
                    else:
                        # fetch, retr hdrs only
                        typ, data = server.fetch(msgnum, '(RFC822.SIZE)')
                        messizes = data[0].decode().split()[2].rstrip(')')
                        allsizes.append(messizes)
                        typ, data = server.fetch(msgnum, '(BODY[HEADER])')
                        hdrlines  = self.decodeFullText(data[0])
                        allhdrs.append('\n'.join(hdrlines))
            finally:
                server.close()
                server.logout()
            assert len(allhdrs) == len(allsizes)
            self.trace('load headers exit')
            return allhdrs, allsizes, False
    
    def downloadAllMessages(self, progress=None, loadfrom=1):
        """ 
        load full message text for all msgs from loadfrom..N,
        despite any caching that may be being done in the caller;
        much slower than downloadAllHeaders, if just need hdrs;
        
        4E: support mailconfig.fetchlimit: see downloadAllHeaders;
        could use server.list() to get sizes of skipped emails here
        too, but clients probably don't care about these anyhow;
        """
        self.trace('loading full messages')
        fetchlimit = mailconfig.fetchlimit
        server = self.connect()
        try:
            typ, data = server.select()
            msgCount = int(data[0].decode())
            allsizes  = []
            allmsgs   = []
            to_load = '(%s:%s)' % (loadfrom, msgCount)
            counter = [i for i in range(loadfrom, msgCount+1)]
            typ, data = server.search(None, to_load)
            for msgnum, msg_counter in zip(data[0].split(), counter):
                if progress: progress(msg_counter, msgCount)
                if fetchlimit and (msg_counter <= msgCount - fetchlimit):
                    # skip, add dummy mail
                    mailtext = 'Subject: --mail skipped--\n\nMail skipped.\n'
                    allsizes.append('')
                    allmsgs.append(mailtext)
                else:
                    # fetch, retr full mail
                    typ, data = server.fetch(msgnum, '(RFC822.SIZE)')
                    messizes = data[0].decode().split()[2].rstrip(')')
                    allsizes.append(messizes)
                    typ, data = server.fetch(msgnum, '(RFC822)')
                    message  = self.decodeFullText(data[0])
                    allmsgs.append('\n'.join(message))
        finally:
            server.close()
            server.logout()
        print(len(allmsgs), len(allsizes), ((msgCount - loadfrom) + 1))
        print(allsizes)
        assert len(allmsgs) == (msgCount - loadfrom) + 1    # msg nums start at 1
       #assert sum(allsizes) == msBytes                     # not if loadfrom > 1
        return allmsgs, allsizes, True
    
    def deleteMessages(self, msgnums, progress=None):
        """ 
        delete multiple msgs off server; assumes email inbox
        unchanged since msgnums were last determined/loaded;
        use if msg headers not available as state information;
        fast, but poss dangerous: see deleteMessagesSafely
        """
        self.trace('deleting mails')
        server = self.connect()
        try:
            server.select()
            typ, data = server.search(None, 'ALL')
            for num in data[0].split():
                if progress: progress(num, len(msgnums))
                server.store(num, '+FLAGS', '\\Deleted')
            server.expunge()
        finally:                        # changes msgnums: reload
            server.close()
            server.logout()
            
    def deleteMessagesSafely(self, msgnums, synchHeaders, progress=None):
        """ 
        delete multipele msgs off server, but use TOP fetches to
        check for a match on each msg's header part before deleting;
        assumes the email server supports the TOP interface of POP,
        else raises TopNotSupported - client may call deleteMessages;
        
        use if the mail server might change the inbox since the email
        index was last fetched, thereby changing POP relative message
        numbers; this can happen if email is deleted in a different
        client; some ISPs may also move a mail from inbox to the
        undeliverable box in response to a failed download;
        
        synchHeaders must be a list of already loaded mail hdrs text,
        corresponding to selected msgnums (requires state); raises
        exception if any out of synch with the email server; inbox is
        locked until quit, so it shoud not change between TOP check
        and actual delete: synch check must occur here, not in caller;
        may be enought to call chckSynchError+deleteMessages, but check
        each msg here in case deletes and inserts in middle of inbox;
        """
        
        # imaplib doest't support TOP interface
        
        if not self.srvrHasTop:
            raise TopNotSupported('Safe delete cancelled')
                       
    def checkSynchError(self, synchHeaders):
        """ 
        check to see if already loaded hdrs text in synchHeaders
        list matches what is on the server, using the TOP command in
        POP to fetch headers text; use if inbox can change due to
        deletes in other client, or automatic action by email server;
        raises except if out of synch, or error while talking to server;
        
        for speed, only checks last in last: this catches inbox deletes,
        but assumes server won't insert before last (true for incoming
        mails); check inbox size first: smaller if just deletes; else
        top will differ if deletes and newly arrived messages added at
        end; result valid only when run: inbox may change after return;
        """
        
        # imaplib doesn't support TOP interface
        
        if not self.srvrHasTop:
            raise TopNotSupported('Check cancelled')
        
    def getPassword(self):
        """ 
        get IMAP password if not yet known
        not required until go to server
        from client-side file or subclass method
        """
        if not self.imapPassword:
            try:
                localfile = open(mailconfig.imappasswdfile)
                self.imapPassword = localfile.readline()[:-1]
                self.trace('local file password' + repr(self.imapPassword))
            except:
                self.imapPassword = self.askImapPassword()
                
    def askImapPassword(self):
        assert False, 'Subclass must define method'
        
    
###############################################################################
# specialized subclasses
###############################################################################

class MailFetcherConsole(MailFetcher):
    def askImapPassword(self):
        import getpass
        prompt = 'Password for %s on %s?' % (self.imapUser, self.imapServer)
        return getpass.getpass(prompt)
    
class SilentMailFetcher(SilentMailTool, MailFetcher):
    pass    #  replace trace
