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
    if client == '' or client is None:
        data = collection.find({
            'user': user,
            'startTime': {
                '$gte': datetime.datetime(startTime[2], startTime[1], startTime[0]),
                '$lte': datetime.datetime(endTime[2], endTime[1], endTime[0])
            }
        })
        return list(x for x in data)
    else:
        data = collection.find({
            'user': user,
            'client_client': client,
            'startTime': {
                '$gte': datetime.datetime(startTime[2], startTime[1], startTime[0]),
                '$lte': datetime.datetime(endTime[2], endTime[1], endTime[0])
            }
        })
        return data
