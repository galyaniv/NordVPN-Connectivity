import speedtest
import math
import json
import pingparsing
import win32com.client as win32
import datetime
import pathlib
import subprocess
import os.path
from os import path
import requests
import random
import schedule
import time
import sys

st = speedtest.Speedtest()
download = []
upload = []
ping = {"packet_receive": 0, "packet_loss_count": 0}
vpn = ''
servers_changed_count = 0
min_download_speed = 0
send_to = ''


class cd:
    def __init__(self, new_path):
        self.new_path = os.path.expanduser(new_path)

    def __enter__(self):
        self.saved_path = os.getcwd()
        os.chdir(self.new_path)

    def __exit__(self, type, value, traceback):
        os.chdir(self.saved_path)


def convert_size(test_result_in_bytes):
    if test_result_in_bytes == 0:
        return 0
    elif test_result_in_bytes < 1024 * 1024:
        return round(test_result_in_bytes / (1024 * 1024), 2)
    else:
        formatted_test_result = round(
            test_result_in_bytes / math.pow(1024, int(math.floor(math.log(test_result_in_bytes, 1024)))), 2)
        return round(formatted_test_result, 2)


def check_ping(host):
    ping_parser = pingparsing.PingParsing()
    transmitter = pingparsing.PingTransmitter()
    transmitter.destination = host
    transmitter.count = 10
    result = json.loads(json.dumps(ping_parser.parse(transmitter.ping()).as_dict(), indent=4))
    ping["packet_receive"] += result["packet_receive"]
    ping["packet_loss_count"] += result["packet_loss_count"]
    return (result["packet_receive"], result["packet_loss_count"])


def statistics():
    avg_downloads = round(sum(download) / len(download), 2)
    avg_upload = round(sum(upload) / len(upload), 2)
    total = ping["packet_receive"] + ping["packet_loss_count"]
    if total != 0:
        ping_statistics = round((ping["packet_receive"] / total) * 100, 2)
    else:
        ping_statistics = 0
    return [avg_downloads, avg_upload, ping_statistics]


def outlook_email_send(auto=True):
    global download
    global upload
    global ping
    global servers_changed_count
    if len(vpn.split(' ')) == 2:
        country_name = vpn.split(' ')[0]
    elif len(vpn.split(' ')) == 3:
        country_name = vpn.split(' ')[0] + ' ' + vpn.split(' ')[1]
    else:
        country_name = vpn
    statistics_info = statistics()
    message_text = "<html><body dir='ltr'><br>Country: {}<br>Current server name: {}<br>Count of servers changed: {}<br><br>Average download speed: {}<br><br>Average upload speed:" \
                   " {}<br><br>Package succeeded: {}<br>Package lost: {}<br>Packages statistics: {}% successes</body></html>".format(
        country_name, vpn, servers_changed_count - 1, statistics_info[0], statistics_info[1],
        ping["packet_receive"], ping["packet_loss_count"], statistics_info[2])

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.Recipients.Add(send_to)
    mail.Subject = 'Network connectivity report'
    mail.HtmlBody = message_text

    if auto:
        try:
            mail.send
        except Exception:
            print("Error raised while sending Email - Are your sure your email is valid?")
    else:
        mail.Display(True)

    download = []
    upload = []
    ping["packet_receive"] = 0
    ping["packet_loss_count"] = 0
    servers_changed_count = 1


def data_file(download_speed, upload_speed, ping_test):
    date = datetime.datetime.now().replace(microsecond=0).date()
    time = datetime.datetime.now().replace(microsecond=0).time()
    text = "{},{},{},{},{},{},{}".format(
        vpn,
        date,
        time,
        download_speed,
        upload_speed,
        ping_test[0],
        ping_test[1])
    dir_name = pathlib.Path(__file__).parent.absolute()
    file = os.path.join(dir_name, 'data' + "." + 'txt')
    if str(path.exists(file)):
        f = open('data.txt', 'a')
        f.close()
    with open("data.txt", 'r+') as f:
        content = f.read()
        f.seek(0, 0)
        f.write(text.rstrip('\r\n') + '\n' + content)
        f.close()


def speed_test():
    global download
    global min_download_speed
    download_speed = convert_size(st.download())
    upload_speed = convert_size(st.upload())
    ping_test = check_ping('google.com')
    download.append(download_speed)
    upload.append(upload_speed)
    data_file(download_speed, upload_speed, ping_test)

    if len(download) > 0 and len(download) % 3 == 0:
        download_speed_check = download[-3:]
        if round(sum(download_speed_check) / 3) < min_download_speed and all(
                d < (min_download_speed + 10) for d in download_speed_check):
            connect_to_recommended_nord_vpn_server()


def connect_to_recommended_nord_vpn_server():
    global vpn
    global servers_changed_count
    servers_changed_count += 1
    recomended_servers = requests.get(f'https://api.nordvpn.com/v1/servers/recommendations')
    recomended_servers_names = [recomended_server['name'] for recomended_server in recomended_servers.json()]
    try:
        if recomended_servers_names:
            random_server = random.randint(0, len(recomended_servers_names) - 1)
            vpn = recomended_servers_names[random_server]
            with cd("C:\\Program Files (x86)\\NordVPN\\"):
                subprocess.run(["nordvpn", "-c", "-n", "{}".format(recomended_servers_names[random_server])])
            return 1
        else:
            return 0
    except:
        return 0


if __name__ == '__main__':

    """command for running the program - python <program name>
    optional: <minimum download speed>
    optional: <email to send report>
    """

    if len(sys.argv) == 1:
        min_download_speed = 50

    elif len(sys.argv) == 2:
        try:
            min_download_speed = int(sys.argv[1])
        except ValueError:
            print('Error: Minimum download speed is an integer value')
            sys.exit()

    elif len(sys.argv) == 3:
        try:
            min_download_speed = int(sys.argv[1])
        except ValueError:
            print('Error: Minimum download speed is an integer value')
            sys.exit()
        send_to = sys.argv[2]
        schedule.every(180).minutes.do(outlook_email_send)
    else:
        print("command for running the program - python <program name>"
              " optional: <minimum download speed>"
              " optional: <email to send report>")
        sys.exit()

    if connect_to_recommended_nord_vpn_server():
        schedule.every(10).minutes.do(speed_test)
        while 1:
            schedule.run_pending()
            time.sleep(1)
