from helper import get_database, get_trips, date_range
import dateutil.parser
import datetime
import xlsxwriter
import sys
import json


def get_pings(trips_list):
    database = get_database()
    collection = database['status']
    data = collection.aggregate([{
        '$match': {
            'tripId': {
                '$in': trips_list
            }
        }
    }, {
        '$group': {
            '_id': '$tripId', 'pings': {'$push': '$$ROOT'}
        }
    }])
    return list(x for x in data)


def get_index(x):
    cnt = 0
    for key, val in headings_dict.items():
        if x == key:
            return cnt
        cnt += 1
    return cnt


def add_new_sheet(data, name):
    name = name.replace('\\', ' ')
    name = name.replace('/', ' ')
    worksheet = workbook.add_worksheet(name)

    row, column = 0, 0
    for heading_key, heading_val in headings_dict.items():
        row = 0
        worksheet.write(row, column, heading_val, HEADING_FORMAT)
        worksheet.set_column(column, column, len(heading_val) + 7)
        row = 1
        for result in data:
            worksheet.write(row, column, str(result[str(heading_key)]))
            row += 1
        column += 1

    worksheet.set_column(0, 0, 25)

    n = len(data)
    for i in range(len(headings_dict.items())):
        worksheet.write(n + 1, i, '', TOTAL_FORMAT)
    worksheet.write(n + 1, 0, "TOTAL ", TOTAL_FORMAT)
    if 'pings' in headings_dict:
        worksheet.write_number(n + 1, get_index('pings'), sum(x['pings'] for x in data), TOTAL_FORMAT)
    if 'trackable' in headings_dict:
        worksheet.write_number(n + 1, get_index('trackable'), sum(1 for i in data if i['trackable'] == 'Y'),
                               TOTAL_FORMAT)
    if billing_type:
        worksheet.write_number(n + 1, get_index('trip_days'), sum(i['trip_days'] for i in data), TOTAL_FORMAT)


def operators_summary(ws, result):
    global has_client_client
    headings2 = ['Operators', 'Total Tracked Trips', 'Tracked Trips', 'Traced Percentage']
    for i in range(len(headings2)):
        ws.write(2, i + 8, headings2[i], HEADING_FORMAT)
    total_trips, total_tracked = 0, 0
    operators = trackable_operators
    operators_data = list()
    for opr in operators:
        temp = list()
        for trip in result:
            if trip['operator'] == opr:
                temp.append(trip['pings'])
        operators_data.append(temp)
    temp = list()
    for trip in result:
        if trip['operator'] not in trackable_operators:
            temp.append(trip['pings'])
    operators_data.append(temp)
    operators.append('Other Operator')
    total_operators = len(operators_data)
    for i in range(total_operators):
        length = len(operators_data[i])
        total_trips += length
        tracked_trips = sum(1 for x in operators_data[i] if x > 0)
        total_tracked += tracked_trips
        ws.write(i + 3, 0 + 8, operators[i])
        ws.write(i + 3, 1 + 8, length)
        ws.write(i + 3, 2 + 8, tracked_trips)
        try:
            ws.write(i + 3, 3 + 8, '{} %'.format(round(sum(1 for i in operators_data[i] if i > 0) / length * 100, 2)))
        except ZeroDivisionError:
            print("Data is Empty for operator \'{}\' ".format(operators_data[i]))
    ws.write(total_operators + 3, 8, 'TOTAL', TOTAL_FORMAT)
    ws.write(total_operators + 3, 9, total_trips, TOTAL_FORMAT)
    ws.write(total_operators + 3, 10, total_tracked, TOTAL_FORMAT)
    try:
        ws.write(total_operators + 3, 11, '{} %'.format(round(total_tracked / total_trips * 100, 2)), TOTAL_FORMAT)
    except ZeroDivisionError as e:
        print("Total Tracked Trips is Empty " + str(e))


def create_summary(result):
    global has_client_client
    ws = workbook.add_worksheet('Summary')
    operators_summary(ws, result)
    headings = ['Source Name', 'Total Trips', 'Trackable Operators', 'Other Operators', 'Total Tracked', 'Trackable %',
                'Tracked %']
    if has_client_client:
        headings[0] = 'Client Name'

    header = sorted(set(opr['client_client'] for opr in result)) if has_client_client else sorted(
        set(opr['source'] for opr in result))
    ws.write(3 + len(header), 0, 'TOTAL', TOTAL_FORMAT)
    total = [0] * 6
    for i in range(len(header)):
        ws.write(3 + i, 0, header[i])
        total_trips = list(
            trip for trip in result if trip['client_client'] == header[i]) if has_client_client else list(
            trip for trip in result if trip['source'] == header[i])
        total[0] += len(total_trips)
        ws.write(3 + i, 1, len(total_trips))
        trackable = sum(1 for trip in total_trips if trip['operator'] in trackable_operators)
        total_tracked = sum(1 for trip in total_trips if trip['pings'] > 0)
        total[1] += trackable
        total[3] += total_tracked
        total[2] += len(total_trips) - trackable
        ws.write(3 + i, 2, trackable)
        ws.write(3 + i, 3, len(total_trips) - trackable)
        ws.write(3 + i, 4, total_tracked)
        try:
            ws.write(3 + i, 5, "{} %".format(round(trackable / len(total_trips) * 100, 2)))
        except ZeroDivisionError as e:
            print("Total Trips are Zero " + str(e))
        try:
            ws.write(3 + i, 6, "{} %".format(round(total_tracked / trackable * 100, 2)))
        except ZeroDivisionError as e:
            print("Total Trackable trips are Zero " + str(e))

    for i in range(len(headings)):
        ws.set_column(i, i, 15)
        ws.write(2, i, headings[i], HEADING_FORMAT)
    try:
        total[4] = '{} %'.format(round(total[1] / total[0] * 100, 2))
        total[5] = '{} %'.format(round(total[3] / total[1] * 100, 2))
    except ZeroDivisionError as e:
        print("Division by Zero in calculating Total. " + str(e))

    for i in range(len(total)):
        ws.write(3 + len(header), 1 + i, total[i], TOTAL_FORMAT)

    trackable_operators_data = list()
    for oper in trackable_operators:
        cnt = sum(1 for values in result if values['operator'] == oper)
        trackable_operators_data.append({oper: cnt})
    ws.set_column(0, 13, 14)


def main():
    global has_client_client
    cnt = 0
    result = list()
    print("Process Starting ..........")
    gmt_to_ist = datetime.timedelta(hours=5, minutes=30)
    temp_trip = list(trip['_id'] for trip in get_trips(username, client, start_time, end_time))
    n = len(temp_trip)
    all_trips_pings = get_pings(temp_trip)
    trips = get_trips(username, client, start_time, end_time)
    # print(list(tt for tt in trips))
    try:
        has_client_client = 'client_client' in trips[0].keys()
    except ValueError:
        if len(client) > 0:
            has_client_client = False
        print("client_client Does not exists")
    if len(client) > 0:
        has_client_client = False
    for trip in trips:
        client_client = ''
        if has_client_client:
            client_client = trip['client_client'] if 'client_client' in trip else "NULL"
        trip_keys = trip.keys()
        running = trip['running']
        start = trip['startTime']
        if isinstance(start, str):
            start = dateutil.parser.parse(start) + gmt_to_ist
        truck_number = ''
        invoice = ''
        if 'truck_number' in trip_keys:
            truck_number = trip['truck_number']

        if 'invoice' in trip_keys:
            invoice = trip['invoice']
        source = ''
        if 'srcname' in trip_keys:
            source = trip['srcname']

        destination = ''
        if 'destname' in trip_keys:
            destination = trip['destname']

        if not running:
            if 'endTime' in trip_keys:
                end = trip['endTime']
            elif 'end_time' in trip_keys:
                end = trip['end_time']
            elif 'forcedEndTime' in trip_keys:
                end = trip['forcedEndTime']
            else:
                raise Exception('ERR = Parameter for End Time is not defined')
            end = end + gmt_to_ist
        else:
            end = datetime.datetime.now()
        tel = trip['tel'][0]
        trip_id = trip['_id']
        operator = ''
        if 'consent' in trip.keys():
            if 'result' in trip['consent']:
                if 'operator' in trip['consent']['result']:
                    operator = trip['consent']['result']['operator']
        if operator is None:
            operator = ''
        no_of_pings = 0
        trip_days = 0
        startTime = datetime.datetime(start.year, start.month, start.day, start.hour, start.month, start.day)
        endTime = datetime.datetime(end.year, end.month, end.day, end.hour, end.month, end.day)
        while startTime < endTime:
            nextTime = startTime + datetime.timedelta(1)
            ok = False
            for trip_pings in all_trips_pings:
                if trip_id == trip_pings['_id']:
                    no_of_pings += len(trip_pings['pings'])
                    for trip_ping in trip_pings['pings']:
                        ping = trip_ping['createdAt'] + gmt_to_ist
                        if startTime < ping < min(nextTime, endTime):
                            ok = True
            if ok:
                trip_days += 1
            startTime = nextTime
        result.append({
            '_id': trip['_id'],
            'trip_id': str(trip_id),
            'truck_number': truck_number,
            'invoice': invoice,
            'source': source,
            'destination': destination,
            'start_time': start.strftime("%d/%m/%Y %H:%M"),
            'end_time': end.strftime("%d/%m/%Y %H:%M"),
            'pings': no_of_pings,
            'tel': tel,
            'operator': operator,
            'trackable': 'Y' if bool(no_of_pings > trackable_threshold) else 'N',
            'trip_days': trip_days if bool(no_of_pings > trackable_threshold) else 0
        })
        if has_client_client:
            result[cnt]['client_client'] = client_client
        cnt += 1
        percentage = n // 13 if n > 13 else n
        if (cnt % percentage) == 0:
            print('Progress = {0:.2f} %'.format(round((cnt / n) * 100, 2)))
    print('Progress = 100 %')
    print('Started writing to the file ' + file_name)
    create_summary(result)
    add_new_sheet(result, 'All Trips')
    for src in set(data['client_client' if has_client_client else 'source'] for data in result):
        data = list(trips for trips in result if trips['client_client' if has_client_client else 'source'] == src)
        add_new_sheet(data, str(src))
    workbook.close()


if __name__ == '__main__':
    # Example Input 1 1 2019 1 2 2019 data TRIPDAYS lplogics
    # Input type Start <DDMMYY> END <DDMMYY> <FileName> <BillingType> <username> <client>
    start_time = list(map(int, sys.argv[1:4]))
    end_time = list(map(int, sys.argv[4:7]))
    file_name = sys.argv[7]
    # TRIP , TRIPDAYS
    billing_type = bool(sys.argv[8] == 'TRIPDAYS')
    username = sys.argv[9]
    if file_name is None or file_name == '':
        file_name = 'data'
    client = ''
    if len(sys.argv) > 10:
        client = sys.argv[10]
        file_name = '{0}-{1}-{2}_{3}-{4}-{5}_{6}'.format(start_time[0], start_time[1], start_time[2], end_time[0],
                                                         end_time[1], end_time[2], client)
    else:
        file_name = '{0}-{1}-{2}_{3}-{4}-{5}_{6}'.format(start_time[0], start_time[1], start_time[2], end_time[0],
                                                         end_time[1], end_time[2], file_name)
    workbook = xlsxwriter.Workbook('temp/' + file_name + '.xlsx')
    BOLD = workbook.add_format({'bold': True})
    HEADING_FORMAT = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#457DC0'})
    TOTAL_FORMAT = workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': '#00B050'})
    config_file = json.load(open('config.json', 'r'))
    trackable_operators = config_file['operators']
    trackable_threshold = config_file['trackable_threshold']
    headings_dict = config_file['headings']
    main()
    exit(0)
