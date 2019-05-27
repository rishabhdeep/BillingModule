from pymongo import MongoClient
import datetime


def date_range(start_date, end_date):
    for date in range(int((end_date - start_date).days)):
        yield start_date + datetime.timedelta(date)


def get_database():
    server, port = open('private', 'r').read().rsplit(':', 1)
    client = MongoClient(server, port=int(port))
    database = client['telenitytracking']
    return database


def get_trips(user, client, startTime, endTime):
    database = get_database()
    collection = database['trips']
    start = datetime.datetime(startTime[2], startTime[1], startTime[0])
    end = datetime.datetime(endTime[2], endTime[1], endTime[0])
    query = {
        'user': user, '$and': [{
            '$or': [{
                'startTime': {
                    '$gte': start
                }
            }, {
                'startTime': {
                    '$gte': start.isoformat()
                }
            }]
        }, {
            '$or': [{
                'startTime': {
                    '$lte': end
                }
            }, {
                'startTime': {
                    '$lte': end.isoformat()
                }
            }]
        }]}
    if client == '' or client is None:
        data = collection.find(query)
        return list(x for x in data)
    else:
        query['client_client'] = client
        data = collection.find(query)
        if isinstance(data, list):
            return data
        else:
            data = list(d for d in data)
            if isinstance(data, list):
                return data
            else:
                data = list(d for d in data)
                return data
