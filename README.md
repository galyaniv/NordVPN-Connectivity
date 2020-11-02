# NordVPN-Connectivity
Get statistics on your NordVPN server connectivity and change servers automatically when connection is bad.<br>
The test start every 10 minutes.<br>
The test results will be stored in data.txt file.<br>
Every 30 minutes if connection is bad, the server will change to one of the recommended servers (by NordVPN API).<br>
If an Email is specified, the statistic's report will be send every 3 hours through Outlook object model.<br>
Using python3.<br>
# Running 
For the default minimum 50 Mbps average download speed before changing server:<br><br>
C:\\> python NetworkStatistics.py <br><br>
For setting up your preferred average download speed before changing server:<br><br>
C:\\> python NetworkStatistics.py 100 <br><br>
For setting up your preferred average download speed before changing server and Email address (for report):<br><br>
C:\\> python NetworkStatistics.py 100 xyz@xyz.com
