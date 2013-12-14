#encoding: utf-8
__author__ = 'ST'

from sqlalchemy import create_engine
import zlib


class con_db():
    def __init__(self, path):
        self.engine = create_engine('sqlite:///' + path)

    def get_mail_list(self, begin=None, end=None):
        limit = ''
        if begin and end:
            limit = 'limit %d, %d' % (begin, end)

        sql = 'select _id, fromAddress, subject, bodyCompressed, ' \
              'toAddresses, joinedAttachmentInfos  from messages ' + limit

        query = self.engine.execute(sql)

        ret = []
        for row in query:
            attachments = self.get_attachments(row[5])
            if row[3]:
                body = '<html><body>' + zlib.decompress(row[3]) + '</body></html>'
            else:
                body = ''
            ret.append({'id': row[0], 'from_address': row[1], 'subject': row[2],
                        'body': body, 'to_address': row[4], 'attachments': attachments})

        return ret

    def get_attachments(self, info):
        ret = []
        attach_info = ''
        if info:
            attach_info = info.split('|')[6::7]
        for item in attach_info:
            sql = "select distinct filename from attachments where originExtras = '" + item + "'"
            query = self.engine.execute(sql)
            for row in query:
                ret.append({'filename': row[0]})
        return ret


if __name__ == '__main__':
    path = r'C:\Users\ST\Documents\GitHub\MailManager' \
           r'\com.google.android.gm\databases\mailstore.funssuse@gmail.com.db'
    ins = con_db(path)
    query = ins.get_mail_list(10, 20)
    for i in query:
        print i.get('attachments')
