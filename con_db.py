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
        self.owner = re.match(r'mailstore\.(.*)\.db', self.db_name).groups()[0]
        mail_dir = os.path.dirname(self.db_dir)
        self.attach_dir = os.path.join(mail_dir, 'cache')

        self.count, = self.engine.execute('select count(*) from messages').first()

    def get_mail_list(self, begin=None, end=None):
        limit = ''
        if begin and end:
            limit = 'limit %d, %d' % (begin, end)

        sql = 'select messageId, fromAddress, subject, bodyCompressed, ' \
              'toAddresses, joinedAttachmentInfos, dateSentMs, ' \
              'dateReceivedMs, snippet  from messages ' + limit

        query = self.engine.execute(sql)

        ret = {}
        labels = self.get_mail_label()
        for row in query:
            attachments = self.get_attachments(row[5])
            if row[3]:
                body = '<html><body>' + zlib.decompress(row[3]) + '</body></html>'
            else:
                body = ''

            d = {'id': row[0], 'from_address': row[1], 'subject': row[2],
                 'body': body, 'to_address': row[4], 'attachments': attachments,
                 'send_data': self.format_time(row[6]),
                 'receive_date': self.format_time(row[7]), 'snippet': row[8]}
            d.update({'to_add_name': self.get_mail_add(d['to_address'])[0],
                      'to_add_add': self.get_mail_add(d['to_address'])[1],
                      'from_add_name': self.get_mail_add(d['from_address'])[0],
                      'from_add_add': self.get_mail_add(d['from_address'])[1]})
            label = '^no_label'
            for lab in labels:
                if d['id'] == lab['message_id']:
                    label = lab['name']
            if label == '^no_label':
                if self.owner == d['from_add_add']:
                    label = u'发件箱'
                elif self.owner == d['to_add_add']:
                    label = u'收件箱'
            d.update({'label': label})

            if not ret.get(label):
                ret[label] = []
            ret[label].append(d)

        return ret

    def get_attachments(self, info):
        ret = []
        if not info:
            return ret
        attach_id = info.split('|')[6::8]
        attach_name = info.split('|')[1::8]
        attach_size = info.split('|')[3::8]
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

    def get_all_label(self):
        sql = "select _id, name from labels"

        ret = []
        query = self.engine.execute(sql)
        for row in query:
            if row[1] and (not row[1][0] == '^'):
                ret.append({'id':row[0], 'name':row[1]})
        return ret

    def get_mail_label(self):
        sql = "select labels_id, message_messageId from message_labels"

        query = self.engine.execute(sql)
        labels = self.get_all_label()
        label_id = [lab['id'] for lab in labels]

        ret = []
        for row in query:
            for label in labels:
                if row[0] == label['id']:
                    ret.append({'name': label['name'], 'message_id': row[1]})
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

    def get_mail_add(self, mail_add):
        query = re.match(r'"(.*)" \<(.*)\>', mail_add)
        return query.groups()


class mail_dbs():
    def __init__(self):
        self.dbs = {}

    def open_db(self, path):
        new_db = con_db(path)
        info = {new_db.owner: new_db.get_mail_list()}
        if not self.dbs.get(new_db.owner):
            self.dbs.update(info)
        else:
            for key in info[new_db.owner].keys():
                if not self.dbs.get(new_db.owner).get(key):
                    self.dbs[new_db.owner][key] = info[new_db.owner][key]
                else:
                    self.dbs[new_db.owner][key] += info[new_db.owner][key]
                    self.dbs[new_db.owner][key] = list(set(self.dbs[new_db.owner][key]))

if __name__ == '__main__':
    path = r'C:\Users\ST\Documents\GitHub\MailManager' \
           r'\com.google.android.gm\databases\mailstore.funssuse@gmail.com.db'
    path2 = r'C:\Users\ST\Documents\GitHub\MailManager' \
            r'\com.google.android.gm.2\databases\mailstore.reamandream@gmail.com.db'
    ins = mail_dbs()
    ins.open_db(path)
    ins.open_db(path2)
    i = ins.dbs
    for key in i.keys():
        print key
        for key2 in i[key].keys():
            print key2
            for i2 in i[key][key2]:
                print i2['subject']