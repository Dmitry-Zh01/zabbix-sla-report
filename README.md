Hello, everyone.<br>
The script allows to get average last month ICMP ping report for hostgroups or hosts and create Excel report.
Also, this script shows last month problems history data in the separate table which includes such columns as: Hostname, Event ID, Problem	Started date and time,	Resolved Event ID, Resolved date and time.

<strong>Example</strong><br>
The first worksheet with SLA info:
![image](https://user-images.githubusercontent.com/106164393/224988164-5cae33f5-9e11-475c-a265-5347aba010e8.png)

The second worksheet with Problems info:
![image](https://user-images.githubusercontent.com/106164393/224988569-02ce95ef-8ebb-46f2-bdf5-675ba9ceb28d.png)

<strong>Requirements</strong><br>
- pyzabbix
- getpass (optionally, you may type directly API user login and password)
- openpyxl

<strong>Important</strong><br>
It is necessary to set by -K argument Zabbix item filter variable according to the monitoring item ('icmpping[,5]' or 'icmpping' or another one), for example:<br>
<strong>-K icmpping[,5]</strong><br>
By default:<br>
<strong>itemFilter = {'key_': 'icmpping'}</strong>  

<strong>How to use?</strong><br>
Please enter hostgroup name and item filter key as argument:<br>
<strong>sla_report.py -G '<hostgroup_name>' -K '<item_key>'</strong><br>
or any number of hosts divided by comma and item filter key:<br>
<strong>sla_report.py -H '<host_name>,<host_name>,<host_name>' -K '<item_key></strong><br>

<strong>Example:</strong><br>
<strong>sla_report.py -G 'Zabbix servers' -K 'icmpping'</strong><br>

Have a nice day.
