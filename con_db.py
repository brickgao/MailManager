#encoding: utf-8
__author__ = 'ST'

from sqlalchemy import create_engine
import zlib
import re
import os
import datetime


class con_db():
    def __init__(self, path):
        self.engine = create_engine('sqlite:///' + path)
        (self.db_dir, self.db_name) = os.path.split(path)
        mail_dir = os.path.dirname(self.db_dir)
        self.attach_dir = os.path.join(mail_dir, 'cache')

        self.count, = self.engine.execute('select count(*) from messages').first()

    def get_mail_list(self, begin=None, end=None):
        limit = ''
        if begin and end:
            limit = 'limit %d, %d' % (begin, end)

        sql = 'select _id, fromAddress, subject, bodyCompressed, ' \
              'toAddresses, joinedAttachmentInfos, dateSentMs, ' \
              'dateReceivedMs, snippet  from messages ' + limit

        query = self.engine.execute(sql)

        ret = []
        for row in query:
            attachments = self.get_attachments(row[5])
            if row[3]:
                body = '<html><body>' + zlib.decompress(row[3]) + '</body></html>'
            else:
                body = ''
            ret.append({'id': row[0], 'from_address': row[1], 'subject': row[2],
                        'body': body, 'to_address': row[4], 'attachments': attachments,
                        'send_data': self.format_time(row[6]),
                        'receive_date': self.format_time(row[7]), 'snippet': row[8]})

        return ret

    def get_attachments(self, info):
        ret = []
        if not info:
            return ret
        attach_id = info.split('|')[6::7]
        attach_name = info.split('|')[1::7]
        attach_size = info.split('|')[3::7]
        for index in range(0, len(attach_id)):
            sql = "select distinct filename from attachments where originExtras = '" \
                  + attach_id[index] + "'"
            query = self.engine.execute(sql)

            format_size = self.format_size(int(attach_size[index]))

            for row in query:
                file_info = {'name': attach_name[index], 'format_size': format_size}
                file_info.update(self.check_file(row[0]))
                ret.append(file_info)
        return ret

    def check_file(self, path):
        info = re.match(r'^file:///data/data/com\.google\.android\.gm/cache/(.+)/(.+)', path)
        if not info:
            return {'exist': False}
        else:
            file_path = os.path.join(self.attach_dir, info.groups()[0], info.groups()[1])
            file_name = info.groups()[1]
            exist = os.path.exists(file_path)
            if exist:
                format_size = self.format_size(os.path.getsize(file_path))
                return {'path': file_path, 'exist': exist, 'format_size': format_size}
            return {'path': file_path, 'exist': exist}

    def format_size(self, size):
        if size < 1024:
            return '%d B' % size
        if size < 1024 * 1024:
            return '%.3f KB' % (size / 1024.0)
        if size < 1024 * 1024 * 1024:
            return '%.3f MB' % (size / 1024.0 / 1024)
        return '%.3f GB' % (size / 1024.0 / 1024 / 1024)

    def format_time(self, time):
        t = datetime.datetime.fromtimestamp(time / 1000.0)
        return t

if __name__ == '__main__':
    path = r'C:\Users\ST\Documents\GitHub\MailManager' \
           r'\com.google.android.gm\databases\mailstore.funssuse@gmail.com.db'
    ins = con_db(path)

    query = ins.get_mail_list(1, 2)
    for i in query:
        print i
