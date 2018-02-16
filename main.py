import urllib.request
import urllib.parse
import http.cookiejar
import json
import pygsheets
import zipfile
import os
import fitparse
import shutil

ZONE1 = 130
ZONE2 = 148

ZONE1_MONEY = 0.5
ZONE2_MONEY = 1

gc = pygsheets.authorize()

# Open spreadsheet and then workseet
sh = gc.open_by_key('1HtkoRVZqJSomA2SoQ5tJC42rMtjmUpX2S2cdwmp7FgE')
wks_overview = sh.worksheet_by_title("Overview")
wks_earning = sh.worksheet_by_title("Earning_history")


# Maximum number of activities you can request at once.  Set and enforced by Garmin.

# URLs for various services.
url_gc_login = 'https://sso.garmin.com/sso/login?service=https%3A%2F%2Fconnect.garmin.com%2Fpost-auth%2' \
               'Flogin&webhost=olaxpw-connect04&source=https%3A%2F%2Fconnect.garmin.com%2Fen-US%2Fsignin&' \
               'redirectAfterAccountLoginUrl=https%3A%2F%2Fconnect.garmin.com%2Fpost-auth%2' \
               'Flogin&redirectAfterAccountCreationUrl=https%3A%2F%2Fconnect.garmin.com%2Fpost-auth%2' \
               'Flogin&gauthHost=https%3A%2F%2Fsso.garmin.com%2Fsso&locale=en_US&id=gauth-widget&cssUrl=' \
               'https%3A%2F%2Fstatic.garmincdn.com%2Fcom.garmin.connect%2Fui%2Fcss%2Fgauth-custom-v1.1-min.css&' \
               'clientId=GarminConnect&rememberMeShown=true&rememberMeChecked=false&createAccountShown=true&' \
               'openCreateAccount=false&usernameShown=false&displayNameShown=false&consumeServiceTicket=false&' \
               'initialFocus=true&embedWidget=false&generateExtraServiceTicket=false'
url_gc_post_auth = 'https://connect.garmin.com/post-auth/login?'
url_gc_search = 'http://connect.garmin.com/proxy/activity-search-service-1.0/json/activities?'
url_gc_gpx_activity = 'http://connect.garmin.com/proxy/activity-service-1.1/gpx/activity/'
url_gc_tcx_activity = 'http://connect.garmin.com/proxy/activity-service-1.1/tcx/activity/'
url_gc_original_activity = 'http://connect.garmin.com/proxy/download-service/files/activity/'


class GarminHandler:
    opener = None
    user = None

    def __init__(self, user):   # log in with user
        self.user = user
        cookie_jar = http.cookiejar.CookieJar()
        self.opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cookie_jar))

        self.http_req(url_gc_login)

        post_data = {'username': user['username'], 'password': user['password'],
                     'embed': 'true', 'lt': 'e1s1', '_eventId': 'submit',
                     'displayNameRequired': 'false'}  # Fields that are passed in a typical Garmin login.

        self.http_req(url_gc_login, post_data)

        login_ticket = None
        for cookie in cookie_jar:
            if cookie.name == 'CASTGC':
                login_ticket = cookie.value
                break

        if not login_ticket:
            raise Exception('Did not get a ticket cookie. Cannot log in. Did you enter the correct username and password?')

        # Chop of 'TGT-' off the beginning, prepend 'ST-0'.
        login_ticket = 'ST-0' + login_ticket[4:]

        self.http_req(url_gc_post_auth + 'ticket=' + login_ticket)

        print('login ok')

    def __del__(self):
        self.opener = None
        print('exit')

    # url is a string, post is a dictionary of POST parameters, headers is a dictionary of headers.
    def http_req(self, url, post=None, headers={}):
        request = urllib.request.Request(url)
        request.add_header('User-Agent', 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 '
                                         '(KHTML, like Gecko) Chrome/1337 Safari/537.36')
        # Tell Garmin we're some supported browser.
        for header_key, header_value in headers.items():
            request.add_header(header_key, header_value)
        if post:
            post = urllib.parse.urlencode(post).encode('utf-8')  # Convert dictionary to POST parameter string.
        response = self.opener.open(request, data=post)  # This line may throw a urllib2.HTTPError.

        # N.B. urllib2 will follow any 302 redirects.
        # Also, the "open" call above may throw a urllib2.HTTPError which is checked for below.

        if response.getcode() != 200:
            raise Exception('Bad return code (' + response.getcode() + ') for: ' + url)

        return response.read()

    def download_all_activities(self):
        total_to_download = 1
        total_downloaded = 0
        all_activities = []

        while total_downloaded < total_to_download:
            # Maximum of 100... 400 return status if over 100.  So download 100 or whatever remains if less than 100.
            if total_to_download - total_downloaded > 100:
                num_to_download = 100
            else:
                num_to_download = total_to_download - total_downloaded

            search_params = {'start': total_downloaded, 'limit': num_to_download}
            # Query Garmin Connect
            result = self.http_req(url_gc_search + urllib.parse.urlencode(search_params))
            result = result.decode('utf-8')
            json_results = json.loads(result)  # TODO: Catch possible exceptions here.

            search = json_results['results']['search']

            total_to_download = int(search['totalFound'])

            # Pull out just the list of activities.
            activities = json_results['results']['activities']

            for singleActivity in activities:
                all_activities.append(singleActivity)

            total_downloaded += num_to_download

        return all_activities

    def download_single_activity_file(self, activity):
        download_url = url_gc_original_activity + activity
        file_mode = 'wb'
        try:
            data = self.http_req(download_url)
        except urllib.request.HTTPError as e:
            # Handle expected (though unfortunate) error codes; die on unexpected ones.
            if e.code == 404:
                # For manual activities (i.e., entered in online without a file upload), there is no original file.
                # Write an empty file to prevent redownloading it.
                print('Writing empty file since there was no original activity data...')
                data = ''
            else:
                raise Exception('Failed. Got an unexpected HTTP error (' + str(e.code) + ').')

        tmp_dir = os.getcwd() + '/tmp'

        if not os.path.exists(tmp_dir):
            os.makedirs(tmp_dir)

        data_filename = tmp_dir + os.path.sep + activity + '.zip'
        save_file = open(data_filename, file_mode)
        save_file.write(data)
        save_file.close()

        return data_filename

    def process_activity_file(self, filepath):
        ret_dict = {}
        ext = os.path.splitext(filepath)[-1].lower()
        dirname = os.path.dirname(filepath)
        if ext == ".zip":
            activity = os.path.splitext(os.path.basename(filepath))[0]
            extract_path = dirname + os.path.sep + activity + os.path.sep

            zip_ref = zipfile.ZipFile(filepath, 'r')
            zip_ref.extractall(extract_path)
            zip_ref.close()

            fit_file = extract_path + activity + '.fit'

            if os.path.exists(fit_file):
                fitfile = fitparse.FitFile(fit_file)

                timestamp = []
                heartrate = []
                # Get all data messages that are of type record
                for record in fitfile.get_messages('record'):

                    # Go through all the data entries in this record

                    for record_data in record:
                        if record_data.name == "heart_rate":
                            heartrate.append(record_data.value)
                        elif record_data.name == "timestamp":
                            timestamp.append(record_data.value)

                timestamp_len = len(timestamp)
                heartrate_len = len(heartrate)

                cal_len = min(timestamp_len, heartrate_len)

                zone1_total_time_min = 0
                zone2_total_time_min = 0
                for i in range(cal_len - 1):
                    time_diff = timestamp[i + 1] - timestamp[i]
                    time_diff_sec = time_diff.total_seconds()

                    if ZONE1 <= heartrate[i] < ZONE2:
                        zone1_total_time_min += time_diff_sec / 60.0
                    elif heartrate[i] >= ZONE2:
                        zone2_total_time_min += time_diff_sec / 60.0

                zone1_total = zone1_total_time_min * ZONE1_MONEY
                zone2_total = zone2_total_time_min * ZONE2_MONEY

                money_earned = zone1_total + zone2_total

                ret_dict = {
                            'money_earned': money_earned,
                            'zone1_minutes': zone1_total_time_min,
                            'zone2_minutes': zone2_total_time_min,
                            'start_time': timestamp[0]
                            }

        return ret_dict

    def update_sheet(self):
        all_activities = self.download_all_activities()

        print('total activity={0}'.format(len(all_activities)))
        wk_activity_list = wks_earning.get_col(2)

        for singleActivity in all_activities:
            activityID = singleActivity['activity']['activityId']
            timestamp = singleActivity['activity']['uploadDate']['millis']
            timestamp = int(timestamp) / 1000
            if str(activityID) not in wk_activity_list and timestamp > 1516410868: # epoch time stamp > 1/20/2018
                file_path = self.download_single_activity_file(activityID)
                cal_data = self.process_activity_file(file_path)
                self.update_sheet_with_single_data(activityID, cal_data)

        # clear tmp folder
        tmp_dir = os.getcwd() + '/tmp'
        if os.path.exists(tmp_dir):
            shutil.rmtree(tmp_dir)

    def update_sheet_with_single_data(self, activityID, cal_data):
        current_row_cell = wks_earning.cell('I1')
        current_row = current_row_cell.value
        user_cell = wks_overview.find(self.user['name'])
        user_cell_label_prefix = ''
        if len(user_cell) == 1:
            user_cell_label = str(user_cell[0].label)
            user_cell_label_prefix = user_cell_label[0]

        current_zone1_min_cell = wks_overview.cell(user_cell_label_prefix + '2')
        current_zone2_min_cell = wks_overview.cell(user_cell_label_prefix + '3')
        current_earned_cell = wks_overview.cell(user_cell_label_prefix + '4')
        current_balance_cell = wks_overview.cell(user_cell_label_prefix + '7')

        current_zone1_min = float(current_zone1_min_cell.value)
        current_zone2_min = float(current_zone2_min_cell.value)
        current_earned = float(current_earned_cell.value)
        current_balance = float(current_balance_cell.value)

        wks_earning.cell('A' + str(current_row)).value = self.user['name']
        wks_earning.cell('B' + str(current_row)).value = activityID
        wks_earning.cell('C' + str(current_row)).value = str(cal_data['start_time'])
        wks_earning.cell('D' + str(current_row)).value = cal_data['zone1_minutes']
        wks_earning.cell('E' + str(current_row)).value = cal_data['zone2_minutes']
        wks_earning.cell('F' + str(current_row)).value = cal_data['money_earned']
        wks_earning.cell('G' + str(current_row)).value = float(current_balance + float(cal_data['money_earned']))

        current_zone1_min_cell.value = float(current_zone1_min + float(cal_data['zone1_minutes']))
        current_zone2_min_cell.value = float(current_zone2_min + float(cal_data['zone1_minutes']))
        current_earned_cell.value = float(current_earned + float(cal_data['money_earned']))
        current_balance_cell.value = float(current_balance + float(cal_data['money_earned']))


users = [
            {'username': 'XXXX',
             'password': 'XXXX',
             'name': 'xxxx'},
            {'username': 'XXXX',
             'password': 'XXXX',
             'name': 'XXXX'}
        ]

for singleUser in users:
    gmHandle = GarminHandler(singleUser)
    gmHandle.update_sheet()
    del gmHandle

# activity = download_all_activities(user_kangmin)
# print('numbers of activity = {0}'.format(len(activity)))
#
# for singleActivity in activity:
#     print('activity ID={0},timestamp={1}'.format(singleActivity['activity']['activityId'],
#                                                  singleActivity['activity']['uploadDate']['millis']))

