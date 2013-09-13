#!/usr/bin/env python
# -*- coding: utf-8 -*- 

import json
import urllib

def CalculateDistance(Origin = False,Destination = False, Method = False,TimeUnits = False,DistUnits = False):
	
	#this is the start of a distnace matrix url
	base = "http://maps.googleapis.com/maps/api/distancematrix/json?"
	
	#Converts the variables to the required format
	urlorigin = "origins=%s&".encode('utf-8') %(Origin)
	urldestination = "destinations=%s&".encode('utf-8') %(Destination)
	urlmethod = "mode=%s&" %(Method)
	if DistUnits == "Kilometers" or DistUnits == "Meters":
		urlunits = "units=metric&"
	else:
		urlunits = "units=imperial&"
	#constructs the completed url
	url = base.decode('utf-8') + urlorigin.decode('utf-8') + urldestination.decode('utf-8') + urlmethod.decode('utf-8') + urlunits.decode('utf-8') + "language=en-EN&sensor=false".decode('utf-8')
	
	#Interprets the json data recieved
	try:
		result= json.load(urllib.urlopen(url))
	except:
		return 'ERROR','ERROR'
	
	#Reads the status code and takes the appropriate action
	if result["status"] == "OK":
		if result["rows"][0]["elements"][0]["status"] == "OK":
			time =  result["rows"][0]["elements"][0]["duration"]["value"]
			distance = result["rows"][0]["elements"][0]["distance"]["value"]
			
			if TimeUnits == "Minutes":
				time = time/60.0
			elif TimeUnits == "Hours":
				time = time/3600.0
				
			if DistUnits == "Kilometres":
				distance = distance/1000.0
			elif DistUnits == "Yards":
				distance = distance*1.0936133
			elif DistUnits == "Miles":
				distance = distance*0.000621371192
					
			return time,distance
		else:
			return result["rows"][0]["elements"][0]["status"],result["rows"][0]["elements"][0]["status"]
	else:
		return result["status"]

def Geoencoding(PostCode = False):
	
	#This is the base of the Geoencoder
	base = "https://maps.googleapis.com/maps/api/geocode/json?"
	urlpostcode = "address=%s&" %(PostCode)
	
	#creates the full url
	url = base + urlpostcode + "language=en-EN&sensor=false"
	
	#Interprates the json code
	result= json.load(urllib.urlopen(url))
	
	#Reads the status code and takes the appropriate action
	if result["status"] == "OK":
		return result["status"],result["results"][0]["geometry"]["location"]["lat"],result["results"][0]["geometry"]["location"]["lng"]
	else:
		return result["status"]

def StaticMaps(DataS,DataD,fileloc):
	urlbase = "http://maps.googleapis.com/maps/api/staticmap?size=640x640&scale=2&"
	urlservicemarkers = "markers=color:0x0000FF|"
	
	for e in DataS:
		urlservicemarkers += "%s|" %(e)
	urlservicemarkers += "&"
	
	urldemandmarkers = "markers=color:0xFFFF00|"
	for e in DataD:
		urldemandmarkers += "%s|" %(e)
		if len(urlbase + urlservicemarkers + urldemandmarkers ) > 1900:
			break
	urldemandmarkers += "&"

	url = urlbase + urlservicemarkers + urldemandmarkers + "sensor=false"
	
	urllib.urlretrieve(url, fileloc + ".jpg")
	
def StaticMapRadius(Sloc,DataDW,DataDO,fileloc):
	urlbase = "http://maps.googleapis.com/maps/api/staticmap?size=640x640&scale=2&"
	
	urlservicemarkers = "markers=color:0x0000FF|"
	
	urlservicemarkers += "%s|" %(Sloc)
	urlservicemarkers += "&"
	
	urldemandmarkers = "markers=color:0x00CC33|"
	for e in DataDW:
		urldemandmarkers += "%s|" %(e)
		if len(urlbase + urlservicemarkers + urldemandmarkers ) > 1900:
			break
	urldemandmarkers += "&"
	
	urldemandmarkers += "markers=color:0xFF0000|"
	for e in DataDO:
		urldemandmarkers += "%s|" %(e)
		if len(urlbase + urlservicemarkers + urldemandmarkers ) > 1900:
			break
	urldemandmarkers += "&"

	url = urlbase + urlservicemarkers + urldemandmarkers + "sensor=false"
	
	urllib.urlretrieve(url, fileloc + ".jpg")
	
def StaticMapHeat(DataS,DataD,hexdict,fileloc):
	urlbase = "http://maps.googleapis.com/maps/api/staticmap?size=640x640&scale=2&"
	
	urlservicemarkers = "markers=color:0x0000FF|"
	for e in DataS:
		urlservicemarkers += "%s|" %(e)
	urlservicemarkers += "&"
	
	urldemandmarkers = ""
	for e in DataD:
		if not DataD[e] == []:
			urldemandmarkers += "markers=color:0x%s|" %(hexdict[e])
			for i in DataD[e]:

				urldemandmarkers += "%s|" %(i)
				
				if len(urlbase + urlservicemarkers + urldemandmarkers ) > 1900:
					break
			urldemandmarkers += "&"
	

	url = urlbase + urlservicemarkers + urldemandmarkers + "sensor=false"
	print url
	urllib.urlretrieve(url, fileloc + ".jpg")