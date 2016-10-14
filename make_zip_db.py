import os
import csv
import cStringIO as StringIO
import shelve

IS_IB = os.environ.get('INSTABASE_URI', None) is not None
RELATIVE_PATH = '/ughani/my-repo/fs/Instabase Drive/files/' if IS_IB else ''
CHUNK_SIZE = 10485760 if IS_IB else -1  # read whole file when not in ib

# open_method = ib.open if IS_IB else open
open_method = open


def chunked_read(filehandle):
    if not IS_IB: return filehandle.read()
    data = []
    while filehandle.tell() != -1:
        data.append(filehandle.read(CHUNK_SIZE))  # reads only 10M at a time.
    return ''.join(data)


read_chunked = chunked_read


class ZipCodeEntry(object):
    def __init__(self, zipcode, city, state, lat, lon):
        self.zipcode = zipcode
        self.city = city
        self.state = state
        self.lat = lat
        self.lon = lon

    def repr(self):
        return "Zip=%(zipcode)s City=%(city)s State=%(state)s Latitude=%(lat)s Longitude=%(lon)s" % dict(
            zipcode=self.zipcode, city=self.city, state=self.state, lat=self.lat, lon=self.lon)

    def __str__(self):
        return "Zip=%(zipcode)s City=%(city)s State=%(state)s Latitude=%(lat)s Longitude=%(lon)s" % dict(
            zipcode=self.zipcode, city=self.city, state=self.state, lat=self.lat, lon=self.lon)

    def __unicode__(self):
        return u"Zip=%(zipcode)s City=%(city)s State=%(state)s Latitude=%(lat)s Longitude=%(lon)s" % dict(
            zipcode=self.zipcode, city=self.city, state=self.state, lat=self.lat, lon=self.lon)

zipcode_entries = []

contents = ''
with open_method('zipcodes/free-zipcode-database-Primary.csv') as zipcodes_file:
    contents = StringIO.StringIO(chunked_read(zipcodes_file))

zipcodes_reader = csv.reader(contents)
header = next(zipcodes_reader)  # skip header

for row in zipcodes_reader:
    zipcode_entries.append(
        ZipCodeEntry(zipcode=row[0], city=row[2], state=row[3],
                     lat=row[5], lon=row[6]))

db = shelve.open('zipcodes/zipcode_db.shelve', 'c')
try:
    for zipcode_entry in zipcode_entries:
        db[zipcode_entry.zipcode] = zipcode_entry
finally:
    db.close()

db = shelve.open('zipcodes/zipcode_db.shelve', 'r')
print db.get('98052', None)
