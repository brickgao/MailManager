#encoding: utf-8
__author__ = 'ST'

from sqlalchemy import create_engine
import zlib

engine2 = create_engine('sqlite:///2.db')

q = engine2.execute('select _id, fromAddress, subject, bodyCompressed from messages')

ret = []
for row in q:
    if row[3]:
        body = '<html><body>' + zlib.decompress(row[3]) + '</body></html>'
    else:
        body = ''

    ret.append({'id': row[0], 'from_address': row[1], 'subject': row[2],
                'body': body})

for i in ret:
    print i.get('id')
    print i.get('from_address')
    print i.get('subject')
    print i.get('body')
    f = open(str(i.get('id')) + '.html', 'w')
    f.write(i.get('body'))
    f.close()
