from intuginehelper import intudb
import datetime
import re

gmt_to_ist = datetime.timedelta(hours=5, minutes=30)


def date_range(start_date, end_date):
    for date in range(int((end_date - start_date).days)):
        yield start_date + datetime.timedelta(date)


def get_trips(user, client, startTime, endTime):
    database = intudb.get_database()
    collection = database['trips']
    start = datetime.datetime(startTime[2], startTime[1], startTime[0]) - gmt_to_ist
    end = datetime.datetime(endTime[2], endTime[1], endTime[0]) - gmt_to_ist

    query = {
        'user': user,
        'invoice': {'$nin': [re.compile("test", re.IGNORECASE)]},
        'truck_number': {'$nin': [re.compile("test", re.IGNORECASE)]},
        'vehicle': {'$nin': [re.compile("test", re.IGNORECASE)]},
        '$and': [{
            '$or': [{
                'startTime': {'$lte': end}
            }, {
                'startTime': {'$lte': end.isoformat()}
            }]
        }, {
            '$or': [{
                'endTime': {'$gte': start}
            }, {
                'endTime': {'$gte': start.isoformat()}
            }, {
                'running': True
            }]
        }
        ]}
    if client == '' or client is None:
        data = collection.find(query)
        return list(x for x in data)
    else:
        query['client_client'] = client
        data = collection.find(query)
        if isinstance(data, list):
            return data
        else:
            res = list()
            for x in data:
                res.append(x)
            return res
