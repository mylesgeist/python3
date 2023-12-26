#!/usr/bin/python
#print "|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||"
#print "###############################################################################################################"
#print "######                                                                                                   ######"
#print "######                       Begining Empty text box                                                     ######"
#print "######                                                                                                   ######"
#print "###############################################################################################################"
# Written in Python 2
import datetime
from dateutil import tz
import sys
import pycurl
import json
import certifi
from StringIO import StringIO
import time

script_start_time = datetime.datetime.utcnow().isoformat('T')
print("---------------------------------------------------------------------------------------------------------------")
print("")

import_id = sys.argv[1:]

import_id = str(import_id)
status = 'queued'
progress = '0'
script_start_time = str(datetime.datetime.utcnow().isoformat('T'))
start_time = str(datetime.datetime.utcnow().isoformat('T'))
end_time = str(datetime.datetime.utcnow().isoformat('T'))


def strip_variables():
	global import_id
	global status
	global progress
	global start_time
	global end_time

	if import_id.startswith("[") and import_id.endswith("]"):
		import_id = import_id[1:-1]
	if import_id.startswith("'") and import_id.endswith("'"):
		import_id = import_id[1:-1]
	if import_id.startswith('"') and import_id.endswith('"'):
		import_id = import_id[1:-1]

	if status.startswith("'") and status.endswith("'"):
		status = status[1:-1]
	if status.startswith('"') and status.endswith('"'):
		status = status[1:-1]
	if status.startswith("'") and status.endswith("'"):
		status = status[1:-1]
	if status.startswith('"') and status.endswith('"'):
		status = status[1:-1]

	if progress.startswith("'") and progress.endswith("'"):
		progress = status[1:-1]
	if progress.startswith('"') and progress.endswith('"'):
		progress = status[1:-1]
	if progress.startswith("'") and progress.endswith("'"):
		progress = status[1:-1]
	if progress.startswith('"') and progress.endswith('"'):
		progress = progress[1:-1]

	if start_time.startswith('"') and start_time.endswith('"'):
		start_time = start_time[1:-1]
	if start_time.startswith("'") and start_time.endswith("'"):
		start_time = start_time[1:-1]
	if start_time.endswith("Z"):
		start_time = start_time[0:-1]

	if end_time.startswith('"') and end_time.endswith('"'):
		end_time = end_time[1:-1]
	if end_time.startswith("'") and end_time.endswith("'"):
		end_time = end_time[1:-1]
	if end_time.endswith("Z"):
		end_time = end_time[0:-1]

		
		
strip_variables()
print("Running Script to Import " + str(import_id) + " at " + str(datetime.datetime.utcnow().isoformat('T')) + ' Universal Time')
retries = 0
json_buffer = StringIO()
c = pycurl.Curl()
c.setopt(c.URL, 'https://app.salsify.com/api/orgs/place-account-id-here/imports/' + str(import_id) + '/runs')
c.setopt(c.WRITEFUNCTION, json_buffer.write)
c.setopt(c.CAINFO, certifi.where())
c.setopt(pycurl.HTTPHEADER, ['Accept:application/json'])
c.setopt(pycurl.POST, 1)
c.setopt(pycurl.POSTFIELDS, 'auth_token=place-auth-token-here')
c.setopt(c.VERBOSE, True)
c.perform()
c.close()

trigger_data = json_buffer.getvalue()
trigger_json = json.loads(trigger_data)

# data for the current transfer
status = json.dumps(trigger_json["status"], indent=4, sort_keys=True)
import_id = json.dumps(trigger_json["import"]["id"], indent=4, sort_keys=True)
import_format = json.dumps(trigger_json["import"]["import_format"]["import_mode"], indent=4, sort_keys=True)
import_name = json.dumps(trigger_json["import"]["name"], indent=4, sort_keys=True)
import_source = json.dumps(trigger_json["import"]["import_source"]["file"], indent=4, sort_keys=True)
run_time = json.dumps(trigger_json["duration"], indent=4, sort_keys=True)
current_id = json.dumps(trigger_json["id"], indent=4, sort_keys=True)
status_summary = json.dumps(trigger_json["status_summary"], indent=4, sort_keys=True)
start_time = json.dumps(trigger_json["start_time"], indent=4, sort_keys=True)
end_time = json.dumps(trigger_json["end_time"], indent=4, sort_keys=True)
failure_reason = json.dumps(trigger_json["failure_reason"], indent=4, sort_keys=True)
progress = json.dumps(trigger_json["progress"], indent=4, sort_keys=True)

strip_variables()



import_queue_retries = 0
while status != 'completed' and retries <= 5:
	if retries >= 1:
		print "Status = " + str(status) + " after " + str(retries*120) + " seconds" 		
	else:
		print "Status = " + str(status)
	json_buffer2 = StringIO()
	c = pycurl.Curl()
	c.setopt(c.URL, 'https://app.salsify.com/api/orgs/s-place-account-id-here/imports/runs/' + current_id + '?auth_token=place-auth-token-here')
	c.setopt(c.WRITEFUNCTION, json_buffer2.write)
	c.setopt(c.CAINFO, certifi.where())
	c.perform()
	c.close()
	status_data = json_buffer2.getvalue()
	while_status_json = json.loads(status_data)
	time.sleep(1)
	status = json.dumps(while_status_json["status"], indent=4, sort_keys=True)
	status_summary = json.dumps(while_status_json["status_summary"], indent=4, sort_keys=True)
	start_time = json.dumps(while_status_json["start_time"], indent=4, sort_keys=True)
	end_time = json.dumps(while_status_json["end_time"], indent=4, sort_keys=True)
	failure_reason = json.dumps(while_status_json["failure_reason"], indent=4, sort_keys=True)
	progress = json.dumps(while_status_json["progress"], indent=4, sort_keys=True)
	strip_variables()
	if while_status_json['status'] != ('running' or 'completing') :
		break
	if retries >=6:
		break
	time.sleep(120)
	retries = retries + 1


end_time = str(datetime.datetime.utcnow().isoformat('T'))
strip_variables()
if start_time != None:
	import_start_time = datetime.datetime.strptime(start_time, "%Y-%m-%dT%H:%M:%S.%f")
else:
	import_start_time = script_start_time
	
if end_time != None:
	import_end_time = datetime.datetime.strptime(end_time, "%Y-%m-%dT%H:%M:%S.%f")
else:
	import_end_time = datetime.datetime.utcnow().isoformat('T')
	
if run_time != None:
	run_time = import_end_time - import_start_time
else:
	run_time = import_start_time - import_end_time
	
script_run_time = import_end_time - import_start_time


print ""
print "   - - - - - - - -  "
print 'Import ID is ' + str(import_id)
print 'Current Import Id is ' + str(current_id)
print 'Import Format is ' + str(import_format)
print 'Import Name is ' + str(import_name)
print 'Import Source is ' + str(import_source)
if status != 'completed':
	print 'Current Status is ' + str(status)
print 'Status Summary is ' + str(status_summary)
print 'Start Time was ' + str(start_time) + ' Universal Time'
print 'End Time was ' + str(end_time) + ' Universal Time'

print 'Script Run Time Was ' + str(script_run_time) + ' Universal Time'
print ""

if script_start_time > end_time:
	print "This import was run previously, it probably didn't import anything this time."
	
if status == 'running' and retries >= 20:
	print 'This Import is Still Running After 10 Minutes at ' + str(datetime.datetime.utcnow().isoformat('T')) + ' Universal Time'
	if failure_reason != 'null':
		print 'Failure Reason is ' + str(failure_reason)
	print 'exit 1'
elif status == 'failed':
	print 'This Import Failed at ' + str(datetime.datetime.utcnow().isoformat('T')) + ' Universal Time'
	print 'Failure Reason is ' + str(failure_reason)
	print 'exit 1'
elif status == 'failure':
	print 'This Import Failed at ' + str(datetime.datetime.utcnow().isoformat('T')) + ' Universal Time'
	print 'Failure Reason is ' + str(failure_reason)
	print 'exit 1'
elif status == 'completed' or status == 'completing':
	print 'Import Number ' + str(import_id) + ' Completed Successfully '
	print 'With a Status of ' + str(status) + ' at ' + str(end_time) + ' Universal Time'
	print 'With a Total Run Time of ' + str(run_time)
	print 'exit 0'
print "   - - - - - - - -  "

print ""
print "---------------------------------------------------------------------------------------------------------------"
#print "###############################################################################################################"
#print "######                                                                                                   ######"
#print "######                       Ending empty text box                                                       ######"
#print "######                                                                                                   ######"
#print "###############################################################################################################"
#print "|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||"



