import sys
from datetime import datetime
import asyncore
from smtpd import SMTPServer
import mailbox
from email.parser import BytesParser
from email.policy import default


class EmlServer(SMTPServer):
    no = 0
    def process_message(self, peer, mailfrom, rcpttos, data, **kwargs):
        filename = '%s-%d.eml' % (datetime.now().strftime('%Y%m%d%H%M%S'),
            self.no)
        filename = '/home/rudolph/Mail/rudolph'

        mbox = mailbox.mbox(filename, factory=BytesParser(policy=default).parse)
        mbox.add(data)
        #print(filename)
        #f = open(filename, 'ab')
        #f.write(data)
        #f.close
        print('%s saved.' % filename)
        self.no += 1
        mbox.flush()

def run(port = 1025):
    EmlServer(('localhost', port), None)
    try:
        asyncore.loop()
    except KeyboardInterrupt:
        pass

if __name__ == '__main__':
    run(int(sys.argv[1]))

