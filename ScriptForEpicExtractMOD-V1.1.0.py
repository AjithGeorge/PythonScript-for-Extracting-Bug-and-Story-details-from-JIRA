from pprint import pprint
from atlassian import Jira  #pip install atlassian-python-api
import json
import jira
from pandas import DataFrame
import pandas
import xlsxwriter
import openpyxl
from openpyxl.styles import Font
import os
import datetime
import requests

Epic ='TEL-9870'

BugDetails=True
StoryDetails=True

StoryJQL = 'project = TEL AND issuetype = Story AND "Epic Link" = '+ Epic +' ORDER BY created DESC'
#BugJQL = 'project = TEL AND issuetype in ("Automation Bug", Bug, Problem) AND "Epic Link" = '+ Epic +' ORDER BY created DESC'
BugJQL = 'project = TEL AND issuetype in ("Automation Bug", Bug, Problem) AND status in ("Code Review", "Failed QA", "In Progress", "In Testing", "On Hold", "Ready for Merge", "Ready for QA", "To Do") AND "Epic Link" = '+ Epic +' ORDER BY created DESC'

d = str(datetime.datetime.now().date())

link = requests.get('https://xyz.atlassian.net/rest/greenhopper/1.0/integration/teamcalendars/sprint/list?jql=project=TEL AND"Epic Link"='+ Epic +'', auth=('ageorge@korewireless.com', 'Welcome#123'))
sprintData=link.json()
p=sprintData["sprints"][0]["start"]
SprintStartDate='     '+p[:2] +'-'+p[2:4]+'-'+p[4:8]

jira = Jira(
    url='https://xyz.atlassian.net/',
    username='ageorge@korewireless.com',
    password='Welcome#123')

if StoryDetails == True:

	data2 = jira.jql(StoryJQL)
	decoded2 = json.dumps(data2)
	
	jsonResponse2=json.loads(decoded2)
	jsondata2 = jsonResponse2["issues"]
	
	Count2=jsonResponse2['total']
	if Count2<=0:
		print("In Stories -NO Match Found for the given Filter ")
	else:
		issuelistArray=[]
		issuetypeArray=[]
		priorityArray=[]
		esttimeArray=[]
		acttimeArray=[]
		qaestimateArray=[]
		statusArray=[]
		testphaseArray=[]
		
		for item in jsondata2:
			issueid = item.get("key")
			issuelistArray.append(issueid)
			issuetype = item.get("fields").get("issuetype").get("name")
			issuetypeArray.append(issuetype)
			priority = item.get("fields").get("priority").get("name")
			priorityArray.append(priority)
			esttime = item.get("fields").get("aggregatetimeoriginalestimate")
			esttimeArray.append(esttime)
			acttime = item.get("fields").get("aggregatetimespent")
			acttimeArray.append(acttime)
			status = item.get("fields").get("status").get("name")
			statusArray.append(status)
			if issuetype == 'Story':
				qaestimate = item.get("fields").get("customfield_15401")
				qaestimateArray.append(qaestimate)
		
		
		df1= pandas.DataFrame(issuelistArray)
		df2= pandas.DataFrame(issuetypeArray)
		df3= pandas.DataFrame(priorityArray)
		df4= pandas.DataFrame(esttimeArray)
		df5= pandas.DataFrame(acttimeArray)
		df6= pandas.DataFrame(qaestimateArray)
		df7= pandas.DataFrame(statusArray)
		writer = pandas.ExcelWriter('Tempdata.xlsx')
		df1.to_excel(writer, startcol = 0, startrow = 1,index = False,header = False)
		df2.to_excel(writer, startcol = 1, startrow = 1,index = False,header = False)
		df3.to_excel(writer, startcol = 2, startrow = 1,index = False,header = False)
		df4.to_excel(writer, startcol = 3, startrow = 1,index = False,header = False)
		df5.to_excel(writer, startcol = 4, startrow = 1,index = False,header = False)
		df6.to_excel(writer, startcol = 5, startrow = 1,index = False,header = False)
		df7.to_excel(writer, startcol = 6, startrow = 1,index = False,header = False)
		
		writer.save()
		
		
		xfile = openpyxl.load_workbook('Tempdata.xlsx')
		sheet = xfile.get_sheet_by_name('Sheet1')
		sheet['A1'] = 'Issue Key'
		sheet['B1'] = 'Issue Type'
		sheet['C1'] = 'Issue Priority'
		sheet['D1'] = 'Original Estimate(Hrs)'
		sheet['E1'] = 'Time Spent(Hrs)'
		sheet['F1'] = 'QA. Est(Hrs)'
		sheet['G1'] = 'Status'
		xfile.save('Tempdata.xlsx')
		
		df = pandas.read_excel('Tempdata.xlsx')
		FORMAT = ['Issue Key', 'Original Estimate(Hrs)', 'Time Spent(Hrs)']
		df_selected = df[FORMAT]
		df1 = df['Original Estimate(Hrs)']
		df1= (df1.div(3600)).round(2)
		df2 = df['Time Spent(Hrs)']
		df2= (df2.div(3600)).round(2)
		df3=df2.subtract(df1)
		df3=DataFrame({'Deviation':df3})
		df4=df['QA. Est(Hrs)']
		df5=df['Status']
		TE= df1.sum()
		TA= df2.sum()
		TD= df3.sum()
		TE=DataFrame({'Total Estimated(Hrs)':TE},index=[0])
		TA=DataFrame({'Total Actuals(Hrs)':TA},index=[0])
		TD=DataFrame({'Total Deviation(Hrs)':TD})
		SD=DataFrame({'1st Sprint Start Date':SprintStartDate},index=[0])
		l=len(df2.index)
		TS=l
		TS=DataFrame({'Total Story Tickets':TS},index=[0])
		l=l+3
		writer = pandas.ExcelWriter(Epic +' ('+ d +') '+'-StoryDetails.xlsx')
		df_selected.to_excel(writer,'Sheet1',index = False)
		
		df1.to_excel(writer, startcol = 1, startrow = 0,index = False)
		df2.to_excel(writer, startcol = 2, startrow = 0,index = False)
		df3.to_excel(writer, startcol = 3, startrow = 0,index = False,header = True)
		df4.to_excel(writer, startcol = 4, startrow = 0,index = False,header = True)
		df5.to_excel(writer, startcol = 5, startrow = 0,index = False,header = True)
		TD.to_excel(writer, startcol = 3, startrow = l,index = False,header = True)
		TA.to_excel(writer, startcol = 2, startrow = l,index = False,header = True)
		TE.to_excel(writer, startcol = 1, startrow = l,index = False,header = True)
		TS.to_excel(writer, startcol = 0, startrow = l,index = False,header = True)
		SD.to_excel(writer, startcol = 4, startrow = l,index = False,header = True)
		writer.save()
		os.remove('Tempdata.xlsx')
else:
	print("Story Details are Opted Out In Code")



	
if BugDetails == True:
	
	data2 = jira.jql(BugJQL)
	decoded2 = json.dumps(data2)
	
	jsonResponse2=json.loads(decoded2)
	jsondata2 = jsonResponse2["issues"]
	
	Count2=jsonResponse2['total']
	if Count2<=0:
		print("In Bugs -NO Match Found for the given Filter ")
	else:
		issuelistArray1=[]
		issuetypeArray1=[]
		priorityArray1=[]
		acttimeArray1=[]
		statusArray1=[]
		testphaseArray1=[]
		rootcauseArray=[]
		outwardissueArray=[]
		tempArray=[]
		
		for item in jsondata2:
		
			issueid = item.get("key")
			issuelistArray1.append(issueid)
			issuetype = item.get("fields").get("issuetype").get("name")
			issuetypeArray1.append(issuetype)
			priority = item.get("fields").get("priority").get("name")
			priorityArray1.append(priority)
			acttime = item.get("fields").get("aggregatetimespent")
			acttimeArray1.append(acttime)
			status = item.get("fields").get("status").get("name")
			statusArray1.append(status)
			if item.get("fields").get("customfield_14900") is None:
				testphaseArray1.append('NotAdded')
			else:
				testphase = item.get("fields").get("customfield_14900").get("value")
				testphaseArray1.append(testphase)
			if item.get("fields").get("customfield_13500")is None:
				rootcauseArray.append('NotAdded')
			else:
				failedreason = item.get("fields").get("customfield_13500").get("value")
				rootcauseArray.append(failedreason)
			for issuelinks in item.get("fields").get("issuelinks"):			
				if 'inwardIssue' in issuelinks:
					linkedissue1 = issuelinks.get("inwardIssue").get("key")
					tempArray.append(linkedissue1)
				else:
					if 'outwardIssue' in issuelinks:
						linkedissue2 = issuelinks.get("outwardIssue").get("key")
						tempArray.append(linkedissue2)
			outwardissueArray.append(tempArray)
			tempArray=[]
				
		df1= pandas.DataFrame(issuelistArray1)
		df2= pandas.DataFrame(issuetypeArray1)
		df3= pandas.DataFrame(priorityArray1)
		df4= pandas.DataFrame(statusArray1)
		df5= pandas.DataFrame(acttimeArray1)
		df5= (df5.div(3600)).round(2)
		df6= pandas.DataFrame(rootcauseArray)
		df7= pandas.DataFrame(testphaseArray1)
		df8= pandas.DataFrame(outwardissueArray)
		writer = pandas.ExcelWriter(Epic +' ('+ d +') '+'-BugDetails.xlsx')
		df1.to_excel(writer, startcol = 0, startrow = 1,index = False,header = False)
		df2.to_excel(writer, startcol = 1, startrow = 1,index = False,header = False)
		df3.to_excel(writer, startcol = 2, startrow = 1,index = False,header = False)
		df4.to_excel(writer, startcol = 3, startrow = 1,index = False,header = False)
		df5.to_excel(writer, startcol = 4, startrow = 1,index = False,header = False)
		df6.to_excel(writer, startcol = 5, startrow = 1,index = False,header = False)
		df7.to_excel(writer, startcol = 6, startrow = 1,index = False,header = False)
		df8.to_excel(writer, startcol = 7, startrow = 1,index = False,header = False)
		
		writer.save()
		
		xfile = openpyxl.load_workbook(Epic +' ('+ d +') '+'-BugDetails.xlsx')
		sheet = xfile.get_sheet_by_name('Sheet1')
		sheet['A1'] = 'Issue Key'
		sheet['B1'] = 'Issue Type'
		sheet['C1'] = 'Issue Priority'
		sheet['D1'] = 'Status'
		sheet['E1'] = 'Time Spent(Hrs)'
		sheet['F1'] = 'Root Cause'
		sheet['G1'] = 'Test Phase'
		sheet['H1'] = 'Linked Stories'
		
		for Title in ('A1','B1','C1','D1','E1','F1','G1','H1'):
			s1=sheet[Title]
			s1.font=Font(bold=True)
		xfile.save(Epic +' ('+ d +') '+'-BugDetails.xlsx')
else:
	print("Bug Details are Opted Out In Code")
	